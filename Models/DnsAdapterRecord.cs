namespace LaptopSessionViewer.Models;

public sealed class DnsAdapterRecord
{
    public required int InterfaceIndex { get; init; }
    public required string InterfaceAlias { get; init; }
    public required string Description { get; init; }
    public required string Status { get; init; }
    public required IReadOnlyList<string> DnsServers { get; init; }
    public required bool IsAutomatic { get; init; }
    public bool HasSavedBackup { get; init; }

    public string DisplayName => $"{InterfaceAlias} ({Status})";

    public string DnsServersText =>
        DnsServers.Count == 0 ? "Automatic" : string.Join(", ", DnsServers);
}
