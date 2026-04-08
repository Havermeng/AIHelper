namespace LaptopSessionViewer.Models;

public sealed class CodexEnvironmentSnapshot
{
    public required bool WingetAvailable { get; init; }
    public required string WingetVersion { get; init; }
    public required bool NodeAvailable { get; init; }
    public required string NodeVersion { get; init; }
    public required bool NpmAvailable { get; init; }
    public required string NpmVersion { get; init; }
    public required bool GitAvailable { get; init; }
    public required string GitVersion { get; init; }
    public required bool CodexAvailable { get; init; }
    public required string CodexVersion { get; init; }
    public required bool LoggedIn { get; init; }
    public required string LoginStatus { get; init; }
    public required bool SessionsFolderExists { get; init; }
    public required string SessionsFolderPath { get; init; }
}
