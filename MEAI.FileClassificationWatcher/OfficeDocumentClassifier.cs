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
                ApplyPassword(passwordAction, v => doc.Password = v);

                doc.Save();
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
                app = new Microsoft.Office.Interop.Excel.Application { Visible = false, DisplayAlerts = false, AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow };
                wb = OpenWithRetry(() => app.Workbooks.Open(path, ReadOnly: false, UpdateLinks: false,
                                    Password: openPassword ?? ""));
                if (wb == null) return false;

                WriteCustomProperty(wb.CustomDocumentProperties, level.ToString());
                dynamic excelWb = wb;
                WriteCategory(excelWb.BuiltInDocumentProperties, level);
                ApplyPassword(passwordAction, () => wb.Password, v => wb.Password = v);

                wb.Save();
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
                ApplyPassword(passwordAction, () => pres.Password, v => pres.Password = v);

                pres.Save();
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
    }
}