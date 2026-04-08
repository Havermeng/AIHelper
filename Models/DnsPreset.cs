using System.Text.Json.Serialization;

namespace LaptopSessionViewer.Models;

public sealed class DnsPreset
{
    public string Name { get; set; } = string.Empty;
    public string PrimaryDns { get; set; } = string.Empty;
    public string SecondaryDns { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public bool EnableDoh { get; set; }
    public string DohTemplate { get; set; } = string.Empty;
    public bool IsCustom { get; set; }
    public bool IsAutomaticPreset { get; set; }

    [JsonIgnore]
    public bool IsAutomatic => IsAutomaticPreset;

    public DnsPreset Clone()
    {
        return new DnsPreset
        {
            Name = Name,
            PrimaryDns = PrimaryDns,
            SecondaryDns = SecondaryDns,
            Description = Description,
            EnableDoh = EnableDoh,
            DohTemplate = DohTemplate,
            IsCustom = IsCustom,
            IsAutomaticPreset = IsAutomaticPreset
        };
    }
}
