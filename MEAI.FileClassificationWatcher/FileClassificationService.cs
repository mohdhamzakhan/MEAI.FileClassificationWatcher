using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Security.Principal;

namespace MEAI.FileClassificationWatcher
{
    // Runs in the logged-on user's session (NOT the SystemMonitorWorker service — that
    // still can't show UI). Watches configured folders and handles classification for
    // EVERY monitored file type, including Word/Excel/PowerPoint — there is no separate
    // VSTO add-in anymore; Office files are classified via OfficeDocumentClassifier's
    // headless COM automation, triggered from the same close-detection flow as plain files.
    //
    // "Closed" isn't a real filesystem event for arbitrary apps, so it's inferred: after a
    // Changed event (or a Created event for Office formats — see OnCreated), poll until the
    // file can be opened with FileShare.None (i.e. no process still holds it), which is the
    // same technique Windows uses to know what to restart during updates. It's a reasonable
    // proxy for "the user is done with it for now," not a guaranteed close signal.
    //
    // TRADE-OFF vs. the old VSTO approach: without an add-in living inside Office, there is
    // no way to intercept/cancel a close while the document is still open in the user's own
    // Word/Excel/PowerPoint window. The classification prompt now appears once the file has
    // already been released, and OfficeDocumentClassifier briefly reopens it invisibly to
    // write the result — there's no "keep editing" recovery path if the user backs out of a
    // mandatory prompt the way there was with DocumentBeforeClose's cancel flag.
    public class FileClassificationService
    {
        private WatcherConfig _config;
        private List<string> _excludedFolderPrefixes;
        private HashSet<string> _excludedFolderNames;
        private readonly ClassificationApiClient _api = new();
        private readonly List<FileSystemWatcher> _watchers = new();

        // One entry per file currently being tracked after a Changed event, so rapid-fire
        // saves reset the same timer instead of spawning duplicate watchers.
        private readonly ConcurrentDictionary<string, CancellationTokenSource> _pendingCloseWatches = new();
        private readonly ConcurrentDictionary<string, DateTime> _selfAppliedUtc =
            new(StringComparer.OrdinalIgnoreCase);
        private static readonly TimeSpan SelfWriteSuppressWindow = TimeSpan.FromSeconds(38);
        private readonly ConcurrentDictionary<string, byte> _inFlight =
                        new(StringComparer.OrdinalIgnoreCase);
        // Tracks whichever ClassificationPromptForm is currently open for a given path, so
        // a Renamed event for that same path can force-close ("supersede") a stale prompt
        // instead of leaving it open alongside a fresh one for the new name.
        private readonly ConcurrentDictionary<string, ClassificationPromptForm> _activePrompts =
                        new(StringComparer.OrdinalIgnoreCase);
        public FileClassificationService(WatcherConfig config)
        {
            _config = config;
            (_excludedFolderPrefixes, _excludedFolderNames) = config.ResolveExclusions();
        }

        // Called when SettingsSyncService detects a newer settings row in the DB. Tears
        // down and rebuilds the watchers with the new folder/extension/exclusion lists —
        // no app restart needed, so a central settings change reaches every machine on its
        // own without anyone touching the endpoint.
        public void Reconfigure(WatcherConfig newConfig)
        {
            Stop();
            _config = newConfig;
            (_excludedFolderPrefixes, _excludedFolderNames) = newConfig.ResolveExclusions();
            Start();
        }

        public void Start()
        {
            foreach (var folder in _config.WatchedFolders)
            {
                if (!Directory.Exists(folder)) continue;

                var watcher = new FileSystemWatcher(folder)
                {
                    IncludeSubdirectories = true,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
                    EnableRaisingEvents = true
                };
                watcher.Created += (_, e) => Task.Run(() => OnCreated(e.FullPath));
                watcher.Changed += (_, e) => OnChanged(e.FullPath);
                watcher.Deleted += (_, e) => Task.Run(() => OnDeletedAsync(e.FullPath));
                watcher.Renamed += (_, e) =>
                {
                    // Cancel any in-flight watch for the OLD path — it's about to point at a
                    // file that no longer exists (classic Explorer "New file" flow: create,
                    // then immediately rename while still in the inline-edit box). Without
                    // this, the stale watch runs to completion and OfficeDocumentClassifier
                    // fails with "couldn't find <old name>" once the popup/Apply cycle
                    // catches up to a file that's already been renamed away.
                    if (_pendingCloseWatches.TryRemove(e.OldFullPath, out var staleCts))
                        staleCts.Cancel();
                    // If a prompt is already open for the old name (mid-classification when
                    // the rename landed), force-close it instead of leaving it sitting there
                    // alongside a fresh prompt that's about to start for the new name.
                    if (_activePrompts.TryRemove(e.OldFullPath, out var staleForm))
                        staleForm.Supersede(); // safe from any thread — marshals internally
                    OnChanged(e.FullPath);
                };
                _watchers.Add(watcher);
            }
        }

        public void Stop()
        {
            foreach (var w in _watchers) w.Dispose();
            _watchers.Clear();
        }

        private bool ShouldIgnore(string path)
        {
            var name = Path.GetFileName(path);
            if (string.IsNullOrEmpty(name)) return true;
            if (ClassificationSidecar.IsSidecarFile(path)) return true;
            if (name.StartsWith("~$") || name.StartsWith(".~")) return true; // Office/editor temp locks
            if (Directory.Exists(path)) return true; // it's a folder, not a file

            // NEW — a classification is already in progress for this exact file;
            // any event that arrives while we're mid-Apply() is noise from our own
            // open/write/save cycle, not a genuine new user edit.
            if (_inFlight.ContainsKey(path)) return true;

            // NEW: ignore the echo Changed event caused by our own headless
            // COM save in OfficeDocumentClassifier.Apply — otherwise the tool
            // re-prompts itself in an infinite loop on the file it just wrote.
            // In ShouldIgnore:
            if (_selfAppliedUtc.TryGetValue(path, out var selfWriteTime)
                && File.GetLastWriteTimeUtc(path) == selfWriteTime)
                return true; // this Changed event is exactly the write we just made

            // WHITELIST: only act on extensions explicitly opted in.
            var ext = Path.GetExtension(path);
            if (!_config.MonitoredExtensions.Contains(ext)) return true;

            if (IsUnderExcludedFolder(path)) return true;

            // Watching a shared network folder means EVERY user's watcher instance sees
            // EVERY file event in that folder — not just their own. Without this check,
            // one person creating a file on a shared drive pops a classification dialog
            // on every other user who happens to be watching the same folder, which is
            // exactly the "very irritating" behavior this was reported as. Only react to
            // files the current Windows user actually owns (created) — NTFS sets the
            // creator as owner by default. If ownership can't be determined at all (e.g.
            // restrictive share permissions), fail safe by skipping rather than risking
            // prompting someone about a file that isn't theirs.
            if (!IsOwnedByCurrentUser(path)) return true;

            return false;
        }

        // NOTE: this means a file someone else created is only ever handled by ITS
        // creator's own watcher instance — if a second person later edits that same
        // shared file, their own watcher will not prompt them for it either, since
        // ownership doesn't change on edit. That's an intentional trade-off: it's the
        // simplest fix for the reported noise problem, but it does mean genuinely
        // collaborative shared documents won't get a fresh confirmation from every editor,
        // only from whoever originally created the file.
        private static bool IsOwnedByCurrentUser(string path)
        {
            try
            {
                var security = new FileInfo(path).GetAccessControl();
                var owner = security.GetOwner(typeof(NTAccount)) as NTAccount;
                var currentUser = WindowsIdentity.GetCurrent().Name; // DOMAIN\username
                var result = owner != null && string.Equals(owner.Value, currentUser, StringComparison.OrdinalIgnoreCase);

                // DIAGNOSTIC: on some network shares (NAS/Samba, or under certain "default
                // owner" GPOs) the NTFS owner metadata returned to a client doesn't reflect
                // the real per-file creator — some fall back to reporting the querying user
                // for every file, which would make this check always match regardless of who
                // actually created the file. This line lets us confirm whether that's
                // happening here before redesigning the ownership approach.
                Logger.LogInfo($"IsOwnedByCurrentUser('{path}'): FileOwner='{owner?.Value ?? "<null>"}', CurrentUser='{currentUser}', Match={result}");

                return result;
            }
            catch (Exception ex)
            {
                // Can't read the ACL (common on locked-down shares for non-owners) —
                // treat as "not mine" rather than risk prompting for someone else's file.
                Logger.LogInfo($"IsOwnedByCurrentUser('{path}'): ACL read failed ({ex.Message}) — treating as not mine.");
                return false;
            }
        }

        private bool IsUnderExcludedFolder(string path)
        {
            var fullPath = Path.GetFullPath(path);

            // Absolute-path exclusions: skip if the file lives under that tree.
            foreach (var prefix in _excludedFolderPrefixes)
            {
                if (fullPath.StartsWith(prefix + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    || fullPath.Equals(prefix, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            // Bare-name exclusions: skip if any directory segment matches, e.g. a file
            // anywhere under "...\node_modules\..." or "...\.git\..." regardless of depth.
            if (_excludedFolderNames.Count > 0)
            {
                var dir = Path.GetDirectoryName(fullPath);
                while (!string.IsNullOrEmpty(dir))
                {
                    var segment = Path.GetFileName(dir);
                    if (!string.IsNullOrEmpty(segment) && _excludedFolderNames.Contains(segment))
                        return true;
                    dir = Path.GetDirectoryName(dir);
                }
            }

            return false;
        }

        private async Task OnCreated(string path)
        {
            try
            {
                if (ShouldIgnore(path)) return;

                // Give the writing app a moment to finish the initial write before we act.
                await Task.Delay(500);
                if (!File.Exists(path)) return; // e.g. a temp file that got renamed/deleted already

                if (OfficeDocumentClassifier.IsOfficeFormat(path))
                {
                    // Office files can't be touched via COM automation while the app that
                    // just created them still has the file open — a brand-new .docx is
                    // open in the user's own Word window the instant it exists. Route
                    // through the same close-detection queue used for edits instead of
                    // prompting immediately; the prompt happens once the file is actually
                    // released, same as HandleOfficeFileClosedAsync for a later edit.
                    QueueCloseWatch(path);
                    return;
                }

                var existing = ClassificationSidecar.Load(path);
                if (existing != null) return; // already tagged (e.g. moved from elsewhere with its sidecar)

                var level = PromptMandatory(path, "This is a new file.\nSelect a classification level:");
                var sidecar = new ClassificationSidecar
                {
                    DocumentGuid = ClassificationSidecar.DeterministicDocumentGuid(path),
                    Classification = level.ToString(),
                    LastContentHash = TryHash(path),
                    LastConfirmed = DateTime.Now
                };
                sidecar.Save(path);

                await _api.LogAsync(new ClassificationLogEntry
                {
                    Application = "File",
                    DocumentName = Path.GetFileName(path),
                    DocumentPath = path,
                    DocumentGuid = sidecar.DocumentGuid,
                    ActionType = "Created",
                    Classification = level.ToString()
                });
            }
            catch (Exception ex)
            {
                // This runs via fire-and-forget Task.Run — without logging, a failure here
                // (locked file, permission denied, disk full, etc.) would vanish silently
                // and it would just look like "nothing happened."
                Logger.LogError($"OnCreated failed for '{path}'", ex);
            }
        }

        private void OnChanged(string path)
        {
            if (ShouldIgnore(path)) return;
            if (!File.Exists(path)) return;
            QueueCloseWatch(path);
        }

        // The file is already gone by the time this fires, so we can't read anything off
        // disk — we can only ask the API what the LAST KNOWN classification was for this
        // path's GUID. That's also why this deliberately skips ShouldIgnore's extension
        // whitelist/ownership checks except where they're still meaningful post-delete
        // (exclusion folders, sidecar files): a Secret/TopSecret file being deleted is
        // worth flagging regardless of who currently "owns" the now-nonexistent file.
        private async Task OnDeletedAsync(string path)
        {
            try
            {
                if (ClassificationSidecar.IsSidecarFile(path)) return;
                var ext = Path.GetExtension(path);
                if (!_config.MonitoredExtensions.Contains(ext)) return;
                if (IsUnderExcludedFolder(path)) return;

                var documentGuid = ClassificationSidecar.DeterministicDocumentGuid(path);
                var currentFromApi = await _api.GetLatestClassificationAsync(documentGuid);
                var level = ClassificationLevelExtensions.Parse(currentFromApi);

                if (level != ClassificationLevel.Secret && level != ClassificationLevel.TopSecret)
                    return; // not classified, or classified below Secret — nothing to flag

                Logger.LogInfo($"OnDeletedAsync: a {level} file was deleted: '{path}'.");

                await _api.LogAsync(new ClassificationLogEntry
                {
                    Application = Path.GetExtension(path).ToLowerInvariant() switch
                    {
                        ".docx" or ".doc" => "Word",
                        ".xlsx" or ".xls" => "Excel",
                        ".pptx" or ".ppt" => "PowerPoint",
                        _ => "File"
                    },
                    DocumentName = Path.GetFileName(path),
                    DocumentPath = path,
                    DocumentGuid = documentGuid,
                    ActionType = "DeletedWhileClassified",
                    Classification = level.Value.ToString(),
                });

                ShowDeletionAlert(path, level.Value);
            }
            catch (Exception ex)
            {
                Logger.LogError($"OnDeletedAsync failed for '{path}'", ex);
            }
        }

        // A plain, informational, non-blocking alert — this is notifying someone that a
        // sensitive file was removed, not asking them to make a decision, so unlike the
        // classification prompt there's nothing to make mandatory and no reason to block
        // the calling thread's caller any longer than the dialog needs to be shown.
        private static void ShowDeletionAlert(string path, ClassificationLevel level)
        {
            var thread = new Thread(() =>
            {
                MessageBox.Show(
                    $"A {level.ToDisplayName()} file was deleted:\n\n\"{path}\"\n\nDeleted by: {WindowsIdentity.GetCurrent().Name}\nTime: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                    "Classified File Deleted",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            // Deliberately not joining here — this is a fire-and-forget notification, not
            // a mandatory decision point like the classification prompt, so it shouldn't
            // hold up OnDeletedAsync's caller (a Task.Run off the FileSystemWatcher thread).
        }

        // Reset any existing watch for this file so multiple saves in a row only result
        // in one confirmation prompt, after things settle down. Shared by both OnCreated
        // (for Office formats) and OnChanged (everything).
        private void QueueCloseWatch(string path)
        {
            if (_inFlight.ContainsKey(path)) return; // already mid-classification; nothing to (re)queue

            if (_pendingCloseWatches.TryRemove(path, out var oldCts))
                oldCts.Cancel();

            var cts = new CancellationTokenSource();
            _pendingCloseWatches[path] = cts;
            _ = WatchForReleaseAsync(path, cts.Token);
        }

        private async Task WatchForReleaseAsync(string path, CancellationToken token)
        {
            var deadline = DateTime.Now.AddMinutes(_config.MaxWatchMinutes);
            try
            {
                if (OfficeDocumentClassifier.IsOfficeFormat(path))
                {
                    // Give Explorer's inline rename box time to finish before the very first
                    // release-check below. A fresh "New > Office Document" leaves the
                    // filename editable for the user to type a real name, and Explorer often
                    // also fires a Changed event (writing the template bytes) shortly after
                    // Created — both routes land here via QueueCloseWatch, so the delay has
                    // to live in this one shared method, not duplicated in each caller.
                    // Since nobody has the file locked yet, IsFileReleased() would otherwise
                    // return true almost instantly and pop the prompt against the placeholder
                    // name before the rename ever lands on disk. Using the token here means a
                    // rename during the wait (which cancels and requeues via QueueCloseWatch)
                    // aborts this wait cleanly instead of running it out pointlessly.
                    await Task.Delay(TimeSpan.FromSeconds(4), token);
                    if (!File.Exists(path)) return; // renamed/deleted during the settle window
                                                    // — whatever handler fired for its new
                                                    // name (if any) already queued its own watch
                }

                while (!token.IsCancellationRequested && DateTime.Now < deadline)
                {
                    await Task.Delay(TimeSpan.FromSeconds(_config.LockPollIntervalSeconds), token);
                    if (!File.Exists(path)) return; // deleted/renamed while we watched

                    if (IsFileReleased(path))
                    {
                        if (OfficeDocumentClassifier.IsOfficeFormat(path))
                            await HandleOfficeFileClosedAsync(path);
                        else
                            await HandlePossibleCloseAsync(path);
                        return;
                    }
                }
            }
            catch (TaskCanceledException)
            {
                // superseded by a newer Changed event — the newer watch will handle it
            }
            catch (Exception ex)
            {
                Logger.LogError($"WatchForReleaseAsync failed for '{path}'", ex);
            }
            finally
            {
                _pendingCloseWatches.TryRemove(path, out _);
            }
        }

        private static bool IsFileReleased(string path)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return true;
            }
            catch (IOException)
            {
                return false; // still open elsewhere
            }
        }

        private async Task HandlePossibleCloseAsync(string path)
        {
            // Marked BEFORE the dialog is shown, not after it returns — a second Changed
            // event for this exact path while the dialog is still open (e.g. two quick
            // saves) would otherwise sail straight through ShouldIgnore/QueueCloseWatch's
            // _inFlight check and spawn a second, fully independent popup for the same
            // file. Removed in `finally` so it covers the dialog AND everything after it.
            if (!_inFlight.TryAdd(path, 0)) return;
            try
            {
                var sidecar = ClassificationSidecar.Load(path) ?? new ClassificationSidecar();
                var documentGuid = ClassificationSidecar.DeterministicDocumentGuid(path);
                var newHash = TryHash(path);

                // Nothing actually changed since we last confirmed (e.g. app just touched
                // the file's timestamp) — don't bother the user. Only applies when we have a
                // sidecar from earlier in this same session; if it was deleted after a prior
                // close, we fall through and always confirm on the first Changed event of a
                // fresh edit session.
                if (newHash == sidecar.LastContentHash && !string.IsNullOrEmpty(sidecar.Classification))
                    return;

                // The sidecar is just a same-session cache — the authoritative "current
                // classification" comes from the API/DB, looked up by the file's deterministic
                // GUID. This matters once the sidecar has been deleted after a previous close:
                // without this, a fresh edit session would have no idea what the file was last
                // classified as. If the API can't be reached, default to Confidential rather
                // than leaving it unclassified or blocking on retries.
                var currentFromApi = await _api.GetLatestClassificationAsync(documentGuid);
                var current = ClassificationLevelExtensions.Parse(currentFromApi)
                              ?? ClassificationLevelExtensions.Parse(sidecar.Classification)
                              ?? ClassificationLevel.Confidential;

                ClassificationLevel? selected = null;
                var thread = new Thread(() =>
                {
                    using var form = new ClassificationPromptForm(
                        Path.GetFileName(path),
                        "Confirm the classification for this updated file.",
                        current,
                        allowCancel: false);
                    _activePrompts[path] = form;
                    if (form.ShowDialog() == DialogResult.OK)
                        selected = form.SelectedLevel;
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                _activePrompts.TryRemove(path, out _);

                if (selected == null)
                {
                    // Reachable now via Supersede() (renamed mid-classification — the old
                    // path is gone, so its sidecar is stale) or, extremely rarely, an OS
                    // force-kill of the dialog process. Either way, clean up the sidecar for
                    // THIS path rather than leaving an orphaned .json sidecar behind — a
                    // fresh cycle for a renamed file's new name creates its own sidecar, and
                    // deleting one that's already gone is a harmless no-op.
                    ClassificationSidecar.Delete(path);
                    return;
                }

                var previous = current.ToString();

                await _api.LogAsync(new ClassificationLogEntry
                {
                    Application = "File",
                    DocumentName = Path.GetFileName(path),
                    DocumentPath = path,
                    DocumentGuid = documentGuid,
                    ActionType = selected.Value.ToString() == previous ? "ConfirmedOnClose" : "ChangedOnClose",
                    PreviousClassification = previous,
                    Classification = selected.Value.ToString()
                });

                // The file is closed and its classification is confirmed/logged to the DB —
                // the sidecar's only job (avoiding re-prompts mid-edit) is done, so clean it up
                // rather than leaving a hidden file behind once the user is finished with it.
                ClassificationSidecar.Delete(path);
            }
            finally
            {
                _activePrompts.TryRemove(path, out _); // no-op if already removed above
                _inFlight.TryRemove(path, out _);
            }
        }

        // Runs once a Word/Excel/PowerPoint file has been released (closed, or saved-and-
        // let-go). No sidecar is used for Office files at all — "is this a genuinely new/
        // unclassified file" and "what is its current level" both come straight from the
        // API, keyed by the file's deterministic path-based GUID. This is what replaces
        // both the VSTO add-ins' Created/Open prompt AND their BeforeClose confirm — both
        // now collapse into this single post-close step, since that's the earliest point
        // this watcher can safely open the file itself.
        private async Task HandleOfficeFileClosedAsync(string path)
        {
            if (!File.Exists(path)) return; // renamed/deleted since release was detected —
                                            // the correct cycle for its new name (if any)
                                            // is already queued separately

            // Marked BEFORE either dialog is shown, not after both return — a second
            // release-detection for this exact path while a prompt is still open (two
            // quick saves, or a large file emitting several Changed events as it flushes)
            // would otherwise sail straight through ShouldIgnore/QueueCloseWatch's
            // _inFlight check and spawn a second, fully independent popup for the same
            // file. Wrapping the whole method — both dialogs plus Apply/LogAsync — in one
            // try/finally is what makes that guard actually cover the window where it matters.
            if (!_inFlight.TryAdd(path, 0))
                return; // another cycle for this exact path is already running — let it finish
            try
            {
                var documentGuid = ClassificationSidecar.DeterministicDocumentGuid(path);
                var currentFromApi = await _api.GetLatestClassificationAsync(documentGuid);
                var currentLevel = ClassificationLevelExtensions.Parse(currentFromApi);
                var isNew = currentLevel == null;

                var headline = isNew
                    ? "This file has no classification yet.\nSelect a classification level:"
                    : "Confirm the classification for this updated file.";

                ClassificationLevel? selected = null;
                var thread = new Thread(() =>
                {
                    using var form = new ClassificationPromptForm(
                        Path.GetFileName(path), headline,
                        currentLevel, allowCancel: false); // no Cancel option, ever — always mandatory
                    _activePrompts[path] = form;
                    if (form.ShowDialog() == DialogResult.OK)
                        selected = form.SelectedLevel;
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                _activePrompts.TryRemove(path, out _);

                // allowCancel is now always false — every prompt is mandatory, and
                // ClassificationPromptForm blocks any force-close that isn't the Confirm
                // button or a deliberate Supersede() (renamed mid-classification), so
                // selected should always be set here except in that supersede case. Kept
                // as a fail-safe rather than an assumption: if a mandatory prompt ever
                // still comes back null, fall back to the file's existing classification
                // if it had one, or Confidential for a genuinely new file — either way
                // avoids silently leaving it unclassified or skipping the write.
                if (selected == null)
                {
                    if (!File.Exists(path))
                    {
                        // Superseded by a rename — this path is gone, there's nothing left
                        // to classify here, and a fresh cycle already started for whatever
                        // the file is now called. Just stop, don't apply a fallback level
                        // to a file that no longer exists at this path.
                        Logger.LogInfo($"HandleOfficeFileClosedAsync: prompt for '{path}' was superseded (file renamed/removed) — skipping.");
                        return;
                    }
                    selected = currentLevel ?? ClassificationLevel.Confidential;
                    Logger.LogInfo($"HandleOfficeFileClosedAsync: mandatory prompt for '{path}' closed without a selection — defaulting to {selected}.");
                }

                var newLevel = selected.Value;
                var levelChanged = newLevel.ToString() != currentFromApi;

                // Tri-state password action — see OfficeDocumentClassifier for the convention.
                // Only ask the user to type a password when the level is NEWLY entering
                // Secret/TopSecret; staying at the same Secret/TopSecret level leaves whatever
                // password is already embedded in the file untouched (we never asked the user
                // to retype a password they already set moments/days ago).
                var knownPassword = PasswordStore.TryGet(documentGuid);
                bool fileAlreadyPassworded = knownPassword != null;
                string? passwordAction = null;
                bool needsSecretLevel = newLevel == ClassificationLevel.Secret || newLevel == ClassificationLevel.TopSecret;

                if (needsSecretLevel && !fileAlreadyPassworded)
                {
                    passwordAction = PromptForPassword(path, newLevel, allowCancel: false); // no Cancel, ever
                    if (passwordAction == null)
                    {
                        // Should be unreachable now that the dialog can't be force-closed
                        // without a valid password, but keep a safe fallback rather than
                        // silently abandoning the whole classification if it somehow happens.
                        Logger.LogError($"HandleOfficeFileClosedAsync: mandatory password prompt for '{path}' returned null unexpectedly.",
                            new Exception("PromptForPassword returned null despite allowCancel: false."));
                        return;
                    }
                    PasswordStore.Save(documentGuid, passwordAction);
                }
                else if (!needsSecretLevel && fileAlreadyPassworded)
                {
                    passwordAction = string.Empty; // downgrading below Secret — clear it
                    PasswordStore.Delete(documentGuid);
                }

                _selfAppliedUtc[path] = DateTime.UtcNow; // mark BEFORE Apply's save can fire its echo
                var applied = OfficeDocumentClassifier.Apply(path, newLevel, passwordAction, knownPassword);

                if (!applied)
                {
                    Logger.LogError($"HandleOfficeFileClosedAsync: failed to apply classification to '{path}' — NOT logging this as a success to the API, so the next close will re-prompt (mandatory, if it was still unclassified) instead of silently treating it as done.",
                        new Exception("OfficeDocumentClassifier.Apply returned false — see prior log entry for the underlying exception."));
                    return; // do not record a classification that was never actually written
                }

                await _api.LogAsync(new ClassificationLogEntry
                {
                    Application = Path.GetExtension(path).ToLowerInvariant() switch
                    {
                        ".docx" or ".doc" => "Word",
                        ".xlsx" or ".xls" => "Excel",
                        ".pptx" or ".ppt" => "PowerPoint",
                        "pdf" => "PDF",
                        _ => "File"
                    },
                    DocumentName = Path.GetFileName(path),
                    DocumentPath = path,
                    DocumentGuid = documentGuid,
                    ActionType = isNew ? "Created" : (levelChanged ? "ChangedOnClose" : "ConfirmedOnClose"),
                    PreviousClassification = currentFromApi,
                    Classification = newLevel.ToString()
                });
            }
            finally
            {
                _activePrompts.TryRemove(path, out _); // no-op if already removed above
                _inFlight.TryRemove(path, out _);
            }
        }

        // Returns the password to set (non-empty string), or null if the user cancelled
        // (only possible when allowCancel is true — see caller).
        private static string? PromptForPassword(string path, ClassificationLevel level, bool allowCancel)
        {
            string? result = null;
            var thread = new Thread(() =>
            {
                using var form = new PasswordEntryForm(Path.GetFileName(path), level.ToDisplayName(), allowCancel);
                if (form.ShowDialog() == DialogResult.OK)
                    result = form.EnteredPassword;
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        private ClassificationLevel PromptMandatory(string path, string headline)
        {
            ClassificationLevel? selected = null;
            var thread = new Thread(() =>
            {
                using var form = new ClassificationPromptForm(Path.GetFileName(path), headline, allowCancel: false);
                form.ShowDialog();
                selected = form.SelectedLevel;
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return selected ?? ClassificationLevel.Confidential;
        }

        private static string TryHash(string path)
        {
            try
            {
                using var sha = SHA256.Create();
                using var stream = File.OpenRead(path);
                return Convert.ToHexString(sha.ComputeHash(stream));
            }
            catch
            {
                return string.Empty; // file locked/unreadable at hash time — treated as "changed" next check
            }
        }
    }
}