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
    public required bool OllamaAvailable { get; init; }
    public required string OllamaDetail { get; init; }
    public required bool LmStudioAvailable { get; init; }
    public required string LmStudioDetail { get; init; }
    public required bool ComfyUiDesktopAvailable { get; init; }
    public required string ComfyUiDesktopDetail { get; init; }
    public required bool PinokioAvailable { get; init; }
    public required string PinokioDetail { get; init; }
    public required bool OpenClawAvailable { get; init; }
    public required string OpenClawDetail { get; init; }
    public required IReadOnlyDictionary<string, string> InstalledOllamaModels { get; init; }
}
