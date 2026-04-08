using System.IO;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class SessionViewerSettingsService
{
    private readonly string _settingsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".codex",
            "session_viewer_settings.json");

    public AppLanguage LoadLanguage()
    {
        if (!File.Exists(_settingsPath))
        {
            return AppLanguage.English;
        }

        try
        {
            var json = File.ReadAllText(_settingsPath);
            var settings = JsonSerializer.Deserialize<SettingsDto>(json);

            return settings?.Language?.ToLowerInvariant() switch
            {
                "ru" => AppLanguage.Russian,
                _ => AppLanguage.English
            };
        }
        catch (Exception exception) when (exception is JsonException or IOException)
        {
            return AppLanguage.English;
        }
    }

    public void SaveLanguage(AppLanguage language)
    {
        var directoryPath = Path.GetDirectoryName(_settingsPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var settings = new SettingsDto
        {
            Language = language == AppLanguage.Russian ? "ru" : "en"
        };

        var json = JsonSerializer.Serialize(
            settings,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_settingsPath, json);
    }

    private sealed class SettingsDto
    {
        public string Language { get; set; } = "en";
    }
}
