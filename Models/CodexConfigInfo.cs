namespace LaptopSessionViewer.Models;

public sealed class CodexConfigInfo
{
    public required string DefaultModel { get; init; }
    public required IReadOnlyList<string> AvailableModels { get; init; }
    public required IReadOnlyList<string> Profiles { get; init; }
}
