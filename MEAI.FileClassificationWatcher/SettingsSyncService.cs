using System;
using System.Configuration; // Required for reading App.config
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MEAI.FileClassificationWatcher
{
    // Fetches WatcherConfig from SystemMonitorAPI (backed by SMM_DOC_CLASSIFICATION_SETTINGS)
    // so config is administered centrally instead of editing appsettings.json on every
    // machine. Falls back to a local cache if the API is unreachable, and to the bundled
    // appsettings.json defaults if there's no cache yet either (first run, offline).
    public class SettingsSyncService
    {
        private const string DefaultApiBaseUrl = "http://10.235.20.49:5295/api";
        private readonly string _apiBaseUrl;

        private const string Profile = "Global";
        private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(15) };
        private static readonly string _cacheFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MEAI", "ClassificationConfig");
        private static readonly string _cacheFile = Path.Combine(_cacheFolder, "settings.cache.json");

        private DateTime? _lastKnownUpdatedAt;
        private readonly PeriodicTimer _timer;
        private readonly WatcherConfig _fallbackDefaults;

        public event Action<WatcherConfig>? SettingsChanged;

        public SettingsSyncService(WatcherConfig fallbackDefaults, TimeSpan? pollInterval = null, string? apiBaseUrl = null)
        {
            _fallbackDefaults = fallbackDefaults;
            _timer = new PeriodicTimer(pollInterval ?? TimeSpan.FromMinutes(15));

            // 1. If a URL is passed explicitly, use it.
            if (!string.IsNullOrWhiteSpace(apiBaseUrl))
            {
                _apiBaseUrl = apiBaseUrl.TrimEnd('/');
            }
            else
            {
                // 2. Otherwise, try to read it from App.config
                var appSettingsUrl = ConfigurationManager.AppSettings["ClassificationApiBaseUrl"];

                // 3. Fallback to the default if App.config is missing the key
                _apiBaseUrl = !string.IsNullOrWhiteSpace(appSettingsUrl)
                    ? appSettingsUrl.TrimEnd('/')
                    : DefaultApiBaseUrl;
            }
        }

        // Call once at startup to get the config to run with immediately.
        public async Task<WatcherConfig> GetInitialConfigAsync()
        {
            var (config, updatedAt) = await TryFetchFromApiAsync();
            if (config != null)
            {
                _lastKnownUpdatedAt = updatedAt;
                SaveCache(config, updatedAt);
                return config;
            }

            var cached = LoadCache();
            if (cached != null)
            {
                _lastKnownUpdatedAt = cached.Value.UpdatedAt;
                return cached.Value.Config;
            }

            // Nothing in the DB yet and no local cache — run with the built-in defaults
            // until an admin publishes a settings row.
            return _fallbackDefaults;
        }

        // Runs forever (until cancelled), checking for a newer settings row and raising
        // SettingsChanged when one shows up, so the running watcher can reconfigure itself
        // without needing a restart.
        public async Task RunPollingLoopAsync(CancellationToken token)
        {
            while (await _timer.WaitForNextTickAsync(token))
            {
                var (config, updatedAt) = await TryFetchFromApiAsync();
                if (config == null) continue; // API unreachable this cycle — keep running with current config

                if (_lastKnownUpdatedAt == null || updatedAt > _lastKnownUpdatedAt)
                {
                    _lastKnownUpdatedAt = updatedAt;
                    SaveCache(config, updatedAt);
                    SettingsChanged?.Invoke(config);
                }
            }
        }

        private async Task<(WatcherConfig? Config, DateTime UpdatedAt)> TryFetchFromApiAsync()
        {
            try
            {
                // FIXED: Now uses _apiBaseUrl instead of the hardcoded ApiBaseUrl constant
                var response = await _http.GetAsync($"{_apiBaseUrl}/documentclassification/settings?profile={Profile}");
                if (!response.IsSuccessStatusCode) return (null, default);

                var body = await response.Content.ReadFromJsonAsync<JsonElement>();
                var configJson = body.GetProperty("configJson").GetString();
                var updatedAt = body.GetProperty("updatedAt").GetDateTime();

                if (string.IsNullOrWhiteSpace(configJson)) return (null, default);

                var config = JsonSerializer.Deserialize<WatcherConfig>(configJson);
                if (config == null) return (null, default);

                ExpandEnvironmentPaths(config);
                return (config, updatedAt);
            }
            catch
            {
                return (null, default); // network/API down — caller falls back to cache/defaults
            }
        }

        private static void ExpandEnvironmentPaths(WatcherConfig config)
        {
            config.WatchedFolders = config.WatchedFolders
                .Select(Environment.ExpandEnvironmentVariables)
                .ToList();
        }

        private void SaveCache(WatcherConfig config, DateTime updatedAt)
        {
            try
            {
                Directory.CreateDirectory(_cacheFolder);
                var payload = new CachedSettings { Config = config, UpdatedAt = updatedAt };
                File.WriteAllText(_cacheFile, JsonSerializer.Serialize(payload));
            }
            catch
            {
                // best-effort cache; if this fails we just re-fetch from the API next time
            }
        }

        private static CachedSettings? LoadCache()
        {
            try
            {
                if (!File.Exists(_cacheFile)) return null;
                return JsonSerializer.Deserialize<CachedSettings>(File.ReadAllText(_cacheFile));
            }
            catch
            {
                return null;
            }
        }

        private struct CachedSettings
        {
            public WatcherConfig Config { get; set; }
            public DateTime UpdatedAt { get; set; }
        }
    }
}