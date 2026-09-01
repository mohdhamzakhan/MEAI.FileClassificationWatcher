using System.Runtime.InteropServices;

namespace MEAI.FileClassificationWatcher
{
    // Applies classification to a Word/Excel/PowerPoint file by briefly opening it in a
    // hidden, non-interactive instance of the app via COM automation, then saving and
    // closing it — this is what replaces the VSTO add-ins. It ONLY runs after
    // FileClassificationService has already confirmed the file is released (not open in
    // the user's own Office session) — attempting this while the user still has the file
    // open would collide with their lock and likely force a read-only open that silently
    // fails to save anything.
    //
    // passwordAction convention (tri-state, since "leave unchanged" and "clear" are both
    // real states a caller needs to express):
    //   null            -> don't touch the existing password
    //   ""  (empty)     -> remove password protection
    //   non-empty value -> set/replace the password with this value
    public static class OfficeDocumentClassifier
    {
        private const int MaxOpenRetries = 3;
        private static readonly TimeSpan RetryDelay = TimeSpan.FromSeconds(1);

        public static bool Apply(string path, ClassificationLevel level, string? passwordAction, string? openPassword = null)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".docx" or ".doc" => ApplyToWord(path, level, passwordAction, openPassword),
                ".xlsx" or ".xls" => ApplyToExcel(path, level, passwordAction, openPassword),
                ".pptx" or ".ppt" => ApplyToPowerPoint(path, level, passwordAction),
                _ => false,
            };
        }

        public static bool IsOfficeFormat(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".docx" or ".doc" or ".xlsx" or ".xls" or ".pptx" or ".ppt";
        }

        private static bool ApplyToWord(string path, ClassificationLevel level, string? passwordAction, string? openPassword = null)
        {
            Microsoft.Office.Interop.Word.Application? app = null;
            Microsoft.Office.Interop.Word.Document? doc = null;
            try
            {
                if (!File.Exists(path)) return false; // renamed/deleted since the close was detected
                ClearReadOnlyAttribute(path);

                app = new Microsoft.Office.Interop.Word.Application
                {
                    Visible = false,
                    DisplayAlerts = 0,
                    AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow
                };
                doc = OpenWithRetry(() => app.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false,
                             Visible: false, PasswordDocument: openPassword ?? ""));
                if (doc == null) return false;

                WriteCustomProperty(doc.CustomDocumentProperties, level.ToString());
                WriteCategory(doc.BuiltInDocumentProperties, level);
                WriteWordHeader(doc, level);
                ApplyPassword(passwordAction, v => doc.Password = v);

                doc.Saved = false;

                SaveWithRetry(path, doc.Save, () => doc.ProtectionType.ToString());
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"OfficeDocumentClassifier.ApplyToWord failed for '{path}'", ex);
                return false;
            }
            finally
            {
                CloseAndQuit(doc, d => d.Close(SaveChanges: false), app, a => a.Quit(SaveChanges: false));
            }
        }

        private static bool ApplyToExcel(string path, ClassificationLevel level, string? passwordAction, string? openPassword = null)
        {
            Microsoft.Office.Interop.Excel.Application? app = null;
            Microsoft.Office.Interop.Excel.Workbook? wb = null;
            try
            {
                if (!File.Exists(path)) return false; // renamed/deleted since the close was detected
                ClearReadOnlyAttribute(path);

                app = new Microsoft.Office.Interop.Excel.Application { Visible = false, DisplayAlerts = false, AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow };
                wb = OpenWithRetry(() => app.Workbooks.Open(path, ReadOnly: false, UpdateLinks: false,
                                    Password: openPassword ?? ""));
                if (wb == null) return false;

                WriteCustomProperty(wb.CustomDocumentProperties, level.ToString());
                dynamic excelWb = wb;
                WriteCategory(excelWb.BuiltInDocumentProperties, level);
                WriteExcelHeader(wb, level);
                ApplyPassword(passwordAction, () => wb.Password, v => wb.Password = v);

                wb.Saved = false;

                SaveWithRetry(path, wb.Save, () => "n/a");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"OfficeDocumentClassifier.ApplyToExcel failed for '{path}'", ex);
                return false;
            }
            finally
            {
                CloseAndQuit(wb, w => w.Close(SaveChanges: false), app, a => a.Quit());
            }
        }

        private static bool ApplyToPowerPoint(string path, ClassificationLevel level, string? passwordAction)
        {
            Microsoft.Office.Interop.PowerPoint.Application? app = null;
            Microsoft.Office.Interop.PowerPoint.Presentation? pres = null;
            try
            {
                if (!File.Exists(path)) return false; // renamed/deleted since the close was detected
                ClearReadOnlyAttribute(path);

                // PowerPoint's Application object doesn't reliably support a headless
                // Visible=false the way Word/Excel do — some PowerPoint versions ignore it,
                // treat it as read-only, or briefly flash a window regardless. Opening with
                // WithWindow=msoFalse below is what actually keeps the presentation window
                // from appearing; Application.Visible is left at its default rather than
                // relying on a property that isn't consistently honored.
                app = new Microsoft.Office.Interop.PowerPoint.Application();
                pres = OpenWithRetry(() => app.Presentations.Open(path, ReadOnly: Microsoft.Office.Core.MsoTriState.msoFalse,
                    Untitled: Microsoft.Office.Core.MsoTriState.msoFalse, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse));
                if (pres == null) return false;

                WriteCustomProperty(pres.CustomDocumentProperties, level.ToString());
                WriteCategory(pres.BuiltInDocumentProperties, level);
                WritePowerPointHeader(pres, level);
                ApplyPassword(passwordAction, () => pres.Password, v => pres.Password = v);

                pres.Saved = Microsoft.Office.Core.MsoTriState.msoFalse;

                SaveWithRetry(path, pres.Save, () => "n/a");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"OfficeDocumentClassifier.ApplyToPowerPoint failed for '{path}'", ex);
                return false;
            }
            finally
            {
                CloseAndQuit(pres, p => p.Close(), app, a => a.Quit());
            }
        }

        private static void ApplyPassword(string? passwordAction, Func<string> getCurrent, Action<string> setPassword)
        {
            if (passwordAction == null) return; // leave whatever password state the file already has
            setPassword(passwordAction); // "" clears it, non-empty sets/replaces it
        }

        private static void ApplyPassword(
    string? passwordAction,
    Action<string> setPassword)
        {
            if (passwordAction == null)
                return;

            setPassword(passwordAction);
        }

        private static void WriteCategory(object builtInPropsObj, ClassificationLevel level)
        {
            try
            {
                dynamic props = builtInPropsObj;
                props["Category"].Value = level.ToDisplayName();
            }
            catch (Exception ex)
            {
                Logger.LogError("WriteCategory failed to set Category property", ex);
            }
        }

        private static void WriteCustomProperty(object customPropsObj, string value)
        {
            const string propName = "MEAI_Classification";
            try
            {
                dynamic props = customPropsObj;
                try { props[propName].Delete(); } catch { /* didn't exist yet */ }
                props.Add(propName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value);
            }
            catch (Exception ex)
            {
                Logger.LogError("WriteCustomProperty failed to set MEAI_Classification property", ex);
            }
        }

        // Files created via Explorer's "New > Word/Excel/PowerPoint Document" menu sometimes
        // inherit the NTFS Read-only attribute from the ShellNew template Explorer copies
        // from. Word/Excel/PowerPoint's own ReadOnly:false open parameter only controls the
        // app's internal editing lock — it does NOT clear this OS-level attribute — so the
        // file opens fine but .Save() throws (0x800A11FD for Word) unless we clear it first.
        private static void ClearReadOnlyAttribute(string path)
        {
            try
            {
                var attrs = File.GetAttributes(path);
                if (attrs.HasFlag(FileAttributes.ReadOnly))
                    File.SetAttributes(path, attrs & ~FileAttributes.ReadOnly);
            }
            catch
            {
                // best effort — if we can't read/clear it, the Open/Save call below will
                // surface a clearer error than swallowing this silently would.
            }
        }

        // A single clear-then-open wasn't enough to fix the "document is read-only" save
        // failure on brand-new files under Desktop — the leading theory is a corporate
        // OneDrive-redirected Desktop briefly re-asserting the Read-only attribute on a
        // just-created file while it registers it for sync, which can happen AFTER our one-
        // time ClearReadOnlyAttribute() call but BEFORE Save() actually runs. So: retry the
        // save itself, re-clearing the attribute on every attempt, and if it still fails
        // after all retries, log enough state (file attribute + doc protection, where
        // available) to tell definitively whether this is really an attribute race or
        // something else (e.g. Restrict Editing / Protected View) — the prior log entries
        // only ever showed the COM exception, not the underlying file/doc state.
        private const int MaxSaveRetries = 5;
        private static readonly TimeSpan SaveRetryDelay = TimeSpan.FromMilliseconds(750);

        private static void SaveWithRetry(string path, Action save, Func<string> describeProtection)
        {
            for (int attempt = 1; attempt <= MaxSaveRetries; attempt++)
            {
                ClearReadOnlyAttribute(path);
                try
                {
                    save();
                    return;
                }
                catch (COMException) when (attempt < MaxSaveRetries)
                {
                    System.Threading.Thread.Sleep(SaveRetryDelay);
                }
                catch (COMException)
                {
                    string attrs = "unknown", protection = "unknown";
                    try { attrs = File.GetAttributes(path).ToString(); } catch { }
                    try { protection = describeProtection(); } catch { }
                    Logger.LogError(
                        $"SaveWithRetry: still read-only after {MaxSaveRetries} attempts for '{path}'. " +
                        $"FileAttributes='{attrs}', DocProtectionType='{protection}'",
                        new Exception("See FileAttributes/DocProtectionType above for root cause."));
                    throw;
                }
            }
        }

        private static T? OpenWithRetry<T>(Func<T> openAction) where T : class
        {
            for (int attempt = 1; attempt <= MaxOpenRetries; attempt++)
            {
                try
                {
                    return openAction();
                }
                catch (COMException) when (attempt < MaxOpenRetries)
                {
                    // Transient sharing violation right after the owning app released the
                    // file — the OS hasn't fully let go yet. Brief backoff and retry.
                    System.Threading.Thread.Sleep(RetryDelay);
                }
            }
            return null;
        }

        // Explicit close + quit + COM release, in that order, wrapped so a failure in one
        // step doesn't skip the others. Orphaned WINWORD.EXE/EXCEL.EXE/POWERPNT.EXE
        // background processes are the single most common bug in this kind of automation
        // if any of these steps get skipped.
        private static void CloseAndQuit<TDoc, TApp>(TDoc? doc, Action<TDoc> closeDoc, TApp? app, Action<TApp> quitApp)
            where TDoc : class where TApp : class
        {
            try { if (doc != null) closeDoc(doc); } catch { /* already gone/inaccessible */ }
            try { if (app != null) quitApp(app); } catch { /* already gone/inaccessible */ }

            if (doc != null) Marshal.FinalReleaseComObject(doc);
            if (app != null) Marshal.FinalReleaseComObject(app);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void WriteWordHeader(Microsoft.Office.Interop.Word.Document doc, ClassificationLevel level)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    var header = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    header.Range.Text = $"Classification: {level.ToDisplayName()}";
                    header.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                    header.Range.Font.Bold = 1;
                    header.Range.Font.Size = 10;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("WriteWordHeader failed to set header text", ex);
            }
        }

        private static void WriteExcelHeader(Microsoft.Office.Interop.Excel.Workbook wb, ClassificationLevel level)
        {
            try
            {
                string headerText = $"&B&U&10Classification: {level.ToDisplayName()}"; // &B = Bold, &10 = 10pt font

                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                {
                    sheet.PageSetup.CenterHeader = headerText;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("WriteExcelHeader failed to set header text", ex);
            }
        }

        private static void WritePowerPointHeader(Microsoft.Office.Interop.PowerPoint.Presentation pres, ClassificationLevel level)
        {
            const string bannerName = "MEAI_ClassificationHeader";
            string text = $"Classification: {level.ToDisplayName()}";

            try
            {
                // Iterate through all design slide masters in the presentation
                for (int d = 1; d <= pres.Designs.Count; d++)
                {
                    var master = pres.Designs[d].SlideMaster;

                    // Remove any previously applied classification banner
                    for (int s = master.Shapes.Count; s >= 1; s--)
                    {
                        if (master.Shapes[s].Name == bannerName)
                        {
                            master.Shapes[s].Delete();
                        }
                    }

                    // Create a full-width header banner at the top (Y = 0)
                    float slideWidth = pres.PageSetup.SlideWidth;
                    var banner = master.Shapes.AddTextbox(
                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                        Left: 0,
                        Top: 0,
                        Width: slideWidth,
                        Height: 18);

                    banner.Name = bannerName;
                    banner.TextFrame.TextRange.Text = text;
                    banner.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    banner.TextFrame.TextRange.Font.Size = 10;
                    banner.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("WritePowerPointHeader failed to set master banner", ex);
            }
        }
    }
}