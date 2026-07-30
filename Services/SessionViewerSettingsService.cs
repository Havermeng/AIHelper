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

    public SessionViewerSettings LoadSettings()
    {
        if (!File.Exists(_settingsPath))
        {
            return new SessionViewerSettings();
        }

        try
        {
            var json = File.ReadAllText(_settingsPath);
            var settings = JsonSerializer.Deserialize<SettingsDto>(json);

            return new SessionViewerSettings
            {
                Language = settings?.Language?.ToLowerInvariant() switch
                {
                    "ru" => AppLanguage.Russian,
                    _ => AppLanguage.English
                },
                DefaultDangerousFullAccess = settings?.DefaultDangerousFullAccess ?? false,
                PhotoPasteFixEnabled = settings?.PhotoPasteFixEnabled ?? false,
                BeginnerModeEnabled = settings?.BeginnerModeEnabled ?? true,
                HasCompletedBeginnerOnboarding = settings?.HasCompletedBeginnerOnboarding ?? false
            };
        }
        catch (Exception exception) when (exception is JsonException or IOException)
        {
            return new SessionViewerSettings();
        }
    }

    public AppLanguage LoadLanguage()
    {
        return LoadSettings().Language;
    }

    public bool LoadDefaultDangerousFullAccess()
    {
        return LoadSettings().DefaultDangerousFullAccess;
    }

    public bool LoadPhotoPasteFixEnabled()
    {
        return LoadSettings().PhotoPasteFixEnabled;
    }

    public void SaveSettings(SessionViewerSettings settings)
    {
        var directoryPath = Path.GetDirectoryName(_settingsPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var dto = new SettingsDto
        {
            Language = settings.Language == AppLanguage.Russian ? "ru" : "en",
            DefaultDangerousFullAccess = settings.DefaultDangerousFullAccess,
            PhotoPasteFixEnabled = settings.PhotoPasteFixEnabled,
            BeginnerModeEnabled = settings.BeginnerModeEnabled,
            HasCompletedBeginnerOnboarding = settings.HasCompletedBeginnerOnboarding
        };

        var json = JsonSerializer.Serialize(
            dto,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_settingsPath, json);
    }

    public void SaveLanguage(AppLanguage language)
    {
        var settings = LoadSettings();
        settings.Language = language;
        SaveSettings(settings);
    }

    public void SaveDefaultDangerousFullAccess(bool enabled)
    {
        var settings = LoadSettings();
        settings.DefaultDangerousFullAccess = enabled;
        SaveSettings(settings);
    }

    public void SavePhotoPasteFixEnabled(bool enabled)
    {
        var settings = LoadSettings();
        settings.PhotoPasteFixEnabled = enabled;
        SaveSettings(settings);
    }

    public void SaveBeginnerModeEnabled(bool enabled)
    {
        var settings = LoadSettings();
        settings.BeginnerModeEnabled = enabled;
        SaveSettings(settings);
    }

    public void SaveHasCompletedBeginnerOnboarding(bool completed)
    {
        var settings = LoadSettings();
        settings.HasCompletedBeginnerOnboarding = completed;
        SaveSettings(settings);
    }

    private sealed class SettingsDto
    {
        public string Language { get; set; } = "en";

        public bool DefaultDangerousFullAccess { get; set; }

        public bool PhotoPasteFixEnabled { get; set; }

        public bool? BeginnerModeEnabled { get; set; }

        public bool? HasCompletedBeginnerOnboarding { get; set; }
    }
}
