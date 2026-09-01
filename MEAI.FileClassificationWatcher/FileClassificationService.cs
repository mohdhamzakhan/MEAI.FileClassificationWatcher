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
        private static readonly TimeSpan SelfWriteSuppressWindow = TimeSpan.FromSeconds(8);
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
                watcher.Renamed += (_, e) => OnChanged(e.FullPath);
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
                return owner != null && string.Equals(owner.Value, currentUser, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                // Can't read the ACL (common on locked-down shares for non-owners) —
                // treat as "not mine" rather than risk prompting for someone else's file.
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

        // Reset any existing watch for this file so multiple saves in a row only result
        // in one confirmation prompt, after things settle down. Shared by both OnCreated
        // (for Office formats) and OnChanged (everything).
        private void QueueCloseWatch(string path)
        {
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
                    allowCancel: true);
                if (form.ShowDialog() == DialogResult.OK)
                    selected = form.SelectedLevel;
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (selected == null) return; // user dismissed — leave classification as-is, don't log a change

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

        // Runs once a Word/Excel/PowerPoint file has been released (closed, or saved-and-
        // let-go). No sidecar is used for Office files at all — "is this a genuinely new/
        // unclassified file" and "what is its current level" both come straight from the
        // API, keyed by the file's deterministic path-based GUID. This is what replaces
        // both the VSTO add-ins' Created/Open prompt AND their BeforeClose confirm — both
        // now collapse into this single post-close step, since that's the earliest point
        // this watcher can safely open the file itself.
        private async Task HandleOfficeFileClosedAsync(string path)
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
                    currentLevel, allowCancel: !isNew); // mandatory only when genuinely unclassified
                if (form.ShowDialog() == DialogResult.OK)
                    selected = form.SelectedLevel;
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            // isNew => allowCancel was false, so selected is always set by the form
            // (defaults to Confidential internally); a null here only happens when the
            // user backed out of a re-confirmation on an already-classified file.
            if (selected == null) return;

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
                passwordAction = PromptForPassword(path, newLevel, allowCancel: !isNew);
                if (passwordAction == null) return;
                PasswordStore.Save(documentGuid, passwordAction);
            }
            else if (!needsSecretLevel && fileAlreadyPassworded)
            {
                passwordAction = string.Empty; // downgrading below Secret — clear it
                PasswordStore.Delete(documentGuid);
            }

            var applied = OfficeDocumentClassifier.Apply(path, newLevel, passwordAction, knownPassword);
            _selfAppliedUtc[path] = File.GetLastWriteTimeUtc(path);   // NEW — stops the echo before it can queue another watch
            if (!applied)
            {
                Logger.LogError($"HandleOfficeFileClosedAsync: failed to apply classification to '{path}'",
                    new Exception("OfficeDocumentClassifier.Apply returned false — see prior log entry for the underlying exception."));
                // Still log the intended classification below so there's an audit trail
                // even though the in-file property/password write failed.
            }

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
                ActionType = isNew ? "Created" : (levelChanged ? "ChangedOnClose" : "ConfirmedOnClose"),
                PreviousClassification = currentFromApi,
                Classification = newLevel.ToString()
            });
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