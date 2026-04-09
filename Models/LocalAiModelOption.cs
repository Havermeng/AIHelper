namespace LaptopSessionViewer.Models;

public sealed class LocalAiModelOption
{
    public required string Name { get; init; }
    public required string ModelTag { get; init; }
    public required string SizeLabel { get; init; }
    public required string Description { get; init; }
    public required bool IsInstalled { get; init; }
    public required string InstalledStatusText { get; init; }
    public required string InstalledStatusBrush { get; init; }
    public required string InstalledSizeText { get; init; }
}
