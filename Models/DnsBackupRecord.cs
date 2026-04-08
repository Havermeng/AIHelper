namespace LaptopSessionViewer.Models;

public sealed class DnsBackupRecord
{
    public required int InterfaceIndex { get; init; }
    public required string InterfaceAlias { get; init; }
    public required List<string> DnsServers { get; init; }
    public bool EnableDoh { get; init; }
    public string DohTemplate { get; init; } = string.Empty;
    public required DateTime SavedAtUtc { get; init; }
}
