using System.Text.Json;
using System.Windows.Forms;

namespace MEAI.FileClassificationWatcher
{
    // Runs as a normal user-mode app (Startup folder or a per-user Scheduled Task set to
    // "Run only when user is logged on") — NOT as a service, so it can show UI. Sits in the
    // system tray with no visible window; SystemMonitorWorker keeps doing its own thing
    // unrelated to this.
    internal static class Program
    {
        private static FileClassificationService? _service;
        private static readonly CancellationTokenSource _pollCts = new();

        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            var bundledDefaults = LoadBundledDefaults();
            var settingsSync = new SettingsSyncService(bundledDefaults);

            // Block briefly on startup so the very first run uses whatever's in the DB
            // (or the local cache from last time) rather than the bundled defaults.
            var initialConfig = settingsSync.GetInitialConfigAsync().GetAwaiter().GetResult();

            _service = new FileClassificationService(initialConfig);
            _service.Start();

            // Central settings changes (an admin updates the DB row) reach this machine
            // on the next poll and take effect immediately — no restart, no per-machine edit.
            settingsSync.SettingsChanged += config => _service.Reconfigure(config);
            _ = settingsSync.RunPollingLoopAsync(_pollCts.Token);

            // No context menu / Exit option here on purpose: this is a compliance tool,
            // so a user shouldn't be able to right-click the tray icon and quit their way
            // out of the classification prompts. It can still be stopped via Task Manager
            // or by an admin disabling the Scheduled Task, but there's no casual one-click
            // exit for the end user.
            using var trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Shield,
                Visible = true,
                Text = "MEAI Document Classification"
            };

            Application.Run();

            _pollCts.Cancel();
            _service.Stop();
        }

        // Only used as a last resort: first run on a machine with no network reachable
        // and no local cache yet from a previous run.
        private static WatcherConfig LoadBundledDefaults()
        {
            var path = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            if (!File.Exists(path)) return new WatcherConfig();

            try
            {
                var config = JsonSerializer.Deserialize<WatcherConfig>(File.ReadAllText(path)) ?? new WatcherConfig();
                config.WatchedFolders = config.WatchedFolders
                    .Select(Environment.ExpandEnvironmentVariables)
                    .ToList();
                return config;
            }
            catch
            {
                return new WatcherConfig();
            }
        }
    }
}