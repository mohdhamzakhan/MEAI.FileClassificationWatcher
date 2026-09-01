using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading.Tasks;
using System.Configuration; // Required for App.config

public class ClassificationLogEntry
{
    public string ClientEventId { get; set; } = Guid.NewGuid().ToString();
    public string Hostname { get; set; } = Environment.MachineName;
    public string Username { get; set; } = Environment.UserName;
    public string Application { get; set; } = string.Empty;
    public string DocumentName { get; set; } = string.Empty;
    public string? DocumentPath { get; set; }
    public string? DocumentGuid { get; set; }
    public string ActionType { get; set; } = string.Empty;
    public string? PreviousClassification { get; set; }
    public string Classification { get; set; } = string.Empty;
    public DateTime EventTime { get; set; } = DateTime.Now;
}

public class ClassificationApiClient
{
    private const string DefaultApiBaseUrl = "http://10.235.20.49:5295/api";

    private readonly string _apiBaseUrl;
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(15) };
    private static readonly string _queueFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MEAI", "ClassificationQueue");
    private static readonly object _queueLock = new();

    public ClassificationApiClient(string? apiBaseUrl = null)
    {
        // 1. If a URL is passed explicitly via code, use it.
        if (!string.IsNullOrWhiteSpace(apiBaseUrl))
        {
            _apiBaseUrl = apiBaseUrl.TrimEnd('/');
            return;
        }

        // 2. Otherwise, try to read it from App.config
        var appSettingsUrl = ConfigurationManager.AppSettings["ClassificationApiBaseUrl"];

        // 3. Fallback to the default if App.config is missing the key
        _apiBaseUrl = !string.IsNullOrWhiteSpace(appSettingsUrl)
            ? appSettingsUrl.TrimEnd('/')
            : DefaultApiBaseUrl;
    }

    public async Task LogAsync(ClassificationLogEntry entry)
    {
        Directory.CreateDirectory(_queueFolder);

        if (!await TrySendAsync(entry))
        {
            var file = Path.Combine(_queueFolder, $"{entry.ClientEventId}.json");
            lock (_queueLock)
            {
                File.WriteAllText(file, JsonSerializer.Serialize(entry));
            }
        }

        await FlushQueueAsync();
    }

    private async Task<bool> TrySendAsync(ClassificationLogEntry entry)
    {
        try
        {
            var response = await _http.PostAsJsonAsync($"{_apiBaseUrl}/documentclassification/log", entry);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public async Task FlushQueueAsync()
    {
        if (!Directory.Exists(_queueFolder)) return;

        foreach (var file in Directory.GetFiles(_queueFolder, "*.json"))
        {
            try
            {
                var json = File.ReadAllText(file);
                var entry = JsonSerializer.Deserialize<ClassificationLogEntry>(json);
                if (entry == null) { File.Delete(file); continue; }

                if (await TrySendAsync(entry))
                    File.Delete(file);
            }
            catch
            {
                // Leave the file in place; we'll retry on the next event.
            }
        }
    }

    public async Task<string?> GetLatestClassificationAsync(string documentGuid)
    {
        try
        {
            var response = await _http.GetAsync(
                $"{_apiBaseUrl}/documentclassification/latest?documentGuid={Uri.EscapeDataString(documentGuid)}");
            if (!response.IsSuccessStatusCode) return null;

            var json = await response.Content.ReadFromJsonAsync<JsonElement>();
            return json.TryGetProperty("classification", out var val) ? val.GetString() : null;
        }
        catch
        {
            return null;
        }
    }
}