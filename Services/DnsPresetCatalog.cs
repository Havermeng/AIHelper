using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public static class DnsPresetCatalog
{
    public static IReadOnlyList<DnsPreset> CreateDefaultPresets(LocalizationService strings)
    {
        return
        [
            new DnsPreset
            {
                Name = strings["DnsPresetAutomatic"],
                PrimaryDns = string.Empty,
                SecondaryDns = string.Empty,
                Description = strings["DnsPresetAutomaticDescription"],
                EnableDoh = false,
                DohTemplate = string.Empty,
                IsCustom = false,
                IsAutomaticPreset = true
            },
            new DnsPreset
            {
                Name = "Cloudflare",
                PrimaryDns = "1.1.1.1",
                SecondaryDns = "1.0.0.1",
                Description = strings["DnsPresetCloudflareDescription"],
                EnableDoh = false,
                DohTemplate = string.Empty,
                IsCustom = false
            },
            new DnsPreset
            {
                Name = "Google",
                PrimaryDns = "8.8.8.8",
                SecondaryDns = "8.8.4.4",
                Description = strings["DnsPresetGoogleDescription"],
                EnableDoh = false,
                DohTemplate = string.Empty,
                IsCustom = false
            },
            new DnsPreset
            {
                Name = "Quad9",
                PrimaryDns = "9.9.9.9",
                SecondaryDns = "149.112.112.112",
                Description = strings["DnsPresetQuad9Description"],
                EnableDoh = false,
                DohTemplate = string.Empty,
                IsCustom = false
            },
            new DnsPreset
            {
                Name = strings["DnsPresetCustom"],
                PrimaryDns = string.Empty,
                SecondaryDns = string.Empty,
                Description = strings["DnsPresetCustomDescription"],
                EnableDoh = false,
                DohTemplate = string.Empty,
                IsCustom = false,
                IsAutomaticPreset = false
            }
        ];
    }
}
