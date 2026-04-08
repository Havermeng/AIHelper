using System.IO;
using System.Net;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class DnsPresetSettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private static readonly IReadOnlyList<DnsPreset> SeededCustomPresets =
    [
        new DnsPreset
        {
            Name = "COMSS DNS",
            PrimaryDns = "83.220.169.155",
            SecondaryDns = "212.109.195.93",
            Description = "Публичный DNS-сервер (COMSS)",
            EnableDoh = true,
            DohTemplate = "https://dns.comss.one/dns-query",
            IsCustom = true,
            IsAutomaticPreset = false
        },
        new DnsPreset
        {
            Name = "Xbox DNS (Smart DNS)",
            PrimaryDns = "111.88.96.50",
            SecondaryDns = "111.88.96.51",
            Description = "Публичный DNS-сервер",
            EnableDoh = true,
            DohTemplate = "https://xbox-dns.ru/dns-query",
            IsCustom = true,
            IsAutomaticPreset = false
        }
    ];

    private readonly AppLogService _logService = new();

    public string PresetsPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "dns-presets.json");

    public IReadOnlyList<DnsPreset> EnsureAndLoadCustomPresets()
    {
        if (!File.Exists(PresetsPath))
        {
            var seededPresets = CreateInitialCustomPresets();
            SaveCustomPresets(seededPresets);
            return seededPresets;
        }

        var loadedPresets = LoadCustomPresets().ToList();
        var changed = UpgradeSeededPresets(loadedPresets);

        if (changed)
        {
            SaveCustomPresets(loadedPresets);
        }

        return loadedPresets;
    }

    public IReadOnlyList<DnsPreset> LoadCustomPresets()
    {
        if (!File.Exists(PresetsPath))
        {
            return [];
        }

        try
        {
            var json = File.ReadAllText(PresetsPath);
            var presets = JsonSerializer.Deserialize<List<DnsPreset>>(json) ?? [];

            return presets
                .Where(preset => preset.IsCustom)
                .Where(preset => !string.IsNullOrWhiteSpace(preset.Name))
                .Select(NormalizeCustomPreset)
                .ToList();
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(DnsPresetSettingsService), "Failed to load custom DNS presets.", exception);
            return [];
        }
    }

    public IReadOnlyList<DnsPreset> LoadAllPresets(LocalizationService strings)
    {
        return DnsPresetCatalog.CreateDefaultPresets(strings)
            .Concat(
                EnsureAndLoadCustomPresets()
                    .Select(preset => preset.Clone())
                    .OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase))
            .ToList();
    }

    public void SaveCustomPresets(IEnumerable<DnsPreset> presets)
    {
        var customPresets = presets
            .Where(preset => preset.IsCustom)
            .Select(
                preset => new DnsPreset
                {
                    Name = preset.Name.Trim(),
                    PrimaryDns = preset.PrimaryDns.Trim(),
                    SecondaryDns = preset.SecondaryDns.Trim(),
                    Description = (preset.Description ?? string.Empty).Trim(),
                    EnableDoh = preset.EnableDoh,
                    DohTemplate = preset.EnableDoh ? (preset.DohTemplate ?? string.Empty).Trim() : string.Empty,
                    IsCustom = true,
                    IsAutomaticPreset = false
                })
            .OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var directory = Path.GetDirectoryName(PresetsPath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(PresetsPath, JsonSerializer.Serialize(customPresets, JsonOptions));
        _logService.Info(nameof(DnsPresetSettingsService), $"Saved {customPresets.Count} custom DNS presets.");
    }

    public IReadOnlyList<DnsPreset> ImportCustomPresets(string filePath)
    {
        var json = File.ReadAllText(filePath);
        var presets = JsonSerializer.Deserialize<List<DnsPreset>>(json) ?? [];

        var imported = presets
            .Where(preset => !string.IsNullOrWhiteSpace(preset.Name))
            .Where(preset => IPAddress.TryParse(preset.PrimaryDns?.Trim(), out _))
            .Where(preset => string.IsNullOrWhiteSpace(preset.SecondaryDns) || IPAddress.TryParse(preset.SecondaryDns.Trim(), out _))
            .Where(
                preset => !preset.EnableDoh ||
                          (Uri.TryCreate(preset.DohTemplate?.Trim(), UriKind.Absolute, out var uri) &&
                           string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            .Select(
                preset => new DnsPreset
                {
                    Name = preset.Name.Trim(),
                    PrimaryDns = preset.PrimaryDns.Trim(),
                    SecondaryDns = preset.SecondaryDns.Trim(),
                    Description = (preset.Description ?? string.Empty).Trim(),
                    EnableDoh = preset.EnableDoh,
                    DohTemplate = preset.EnableDoh ? (preset.DohTemplate ?? string.Empty).Trim() : string.Empty,
                    IsCustom = true,
                    IsAutomaticPreset = false
                })
            .ToList();

        _logService.Info(nameof(DnsPresetSettingsService), $"Imported {imported.Count} DNS presets from {filePath}.");
        return imported;
    }

    public void ExportCustomPresets(string filePath, IEnumerable<DnsPreset> presets)
    {
        var customPresets = presets
            .Where(preset => preset.IsCustom)
            .Select(NormalizeCustomPreset)
            .OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var directory = Path.GetDirectoryName(filePath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(filePath, JsonSerializer.Serialize(customPresets, JsonOptions));
        _logService.Info(nameof(DnsPresetSettingsService), $"Exported {customPresets.Count} DNS presets to {filePath}.");
    }

    private static DnsPreset NormalizeCustomPreset(DnsPreset preset)
    {
        return new DnsPreset
        {
            Name = preset.Name.Trim(),
            PrimaryDns = preset.PrimaryDns.Trim(),
            SecondaryDns = preset.SecondaryDns.Trim(),
            Description = (preset.Description ?? string.Empty).Trim(),
            EnableDoh = preset.EnableDoh,
            DohTemplate = preset.EnableDoh ? (preset.DohTemplate ?? string.Empty).Trim() : string.Empty,
            IsCustom = true,
            IsAutomaticPreset = false
        };
    }

    private static IReadOnlyList<DnsPreset> CreateInitialCustomPresets()
    {
        return SeededCustomPresets
            .Select(preset => preset.Clone())
            .ToList();
    }

    private static bool UpgradeSeededPresets(IList<DnsPreset> presets)
    {
        var changed = false;

        foreach (var seededPreset in SeededCustomPresets)
        {
            var existingPreset = presets.FirstOrDefault(
                preset => string.Equals(preset.Name, seededPreset.Name, StringComparison.OrdinalIgnoreCase));

            if (existingPreset is null)
            {
                presets.Add(seededPreset.Clone());
                changed = true;
                continue;
            }

            if (string.IsNullOrWhiteSpace(existingPreset.Description) ||
                string.Equals(existingPreset.Description, "РџСѓР±Р»РёС‡РЅС‹Р№ DNS-СЃРµСЂРІРµСЂ", StringComparison.Ordinal))
            {
                existingPreset.Description = seededPreset.Description;
                changed = true;
            }

            if (!existingPreset.EnableDoh)
            {
                existingPreset.EnableDoh = true;
                changed = true;
            }

            if (!string.Equals(existingPreset.DohTemplate, seededPreset.DohTemplate, StringComparison.OrdinalIgnoreCase))
            {
                existingPreset.DohTemplate = seededPreset.DohTemplate;
                changed = true;
            }
        }

        return changed;
    }
}
