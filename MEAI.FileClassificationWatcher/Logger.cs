namespace MEAI.FileClassificationWatcher
{
    // OnCreated/OnChanged run via fire-and-forget Task.Run — an unhandled exception there
    // just disappears (no crash, no message, the file simply never gets tagged). This gives
    // that failure somewhere to land instead of vanishing silently.
    internal static class Logger
    {
        private static readonly string _logFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MEAI", "FileClassificationWatcher");
        private static readonly string _logFile = Path.Combine(_logFolder, "log.txt");
        private static readonly object _lock = new();

        public static void LogError(string message, Exception ex)
        {
            try
            {
                Directory.CreateDirectory(_logFolder);
                lock (_lock)
                {
                    File.AppendAllText(_logFile,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}: {ex}\n\n");
                }
            }
            catch
            {
                // if we can't even write the log, there's nowhere left to report this
            }
        }
    }
}