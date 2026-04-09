namespace LaptopSessionViewer.Models;

public sealed class CreativeAiToolOption
{
    public required string Name { get; init; }
    public required string PackageId { get; init; }
    public required string CoverageLabel { get; init; }
    public required string Description { get; init; }
    public required bool IsInstalled { get; init; }
    public required string InstalledStatusText { get; init; }
    public required string InstalledStatusBrush { get; init; }
    public required string InstalledDetailText { get; init; }
}
