namespace LaptopSessionViewer.Models;

public sealed class SetupCheckItem
{
    public required string Title { get; init; }
    public required string Status { get; init; }
    public required string Detail { get; init; }
    public required string AccentBrush { get; init; }
}
