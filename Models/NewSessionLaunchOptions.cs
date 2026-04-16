namespace LaptopSessionViewer.Models;

public sealed class NewSessionLaunchOptions
{
    public required string Prompt { get; init; }
    public required string WorkingDirectory { get; init; }
    public IReadOnlyList<string> ImagePaths { get; init; } = [];
    public required string Model { get; init; }
    public required string Profile { get; init; }
    public required string SandboxMode { get; init; }
    public required string ApprovalPolicy { get; init; }
    public required string LocalProvider { get; init; }
    public required bool UseSearch { get; init; }
    public required bool UseOss { get; init; }
    public required bool UseDangerousBypass { get; init; }
}
