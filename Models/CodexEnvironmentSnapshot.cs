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
    public required bool CodexDesktopAppAvailable { get; init; }
    public required string CodexDesktopAppDetail { get; init; }
    public required bool CodexAvailable { get; init; }
    public required string CodexVersion { get; init; }
    public required bool LoggedIn { get; init; }
    public required string LoginStatus { get; init; }
    public required bool SessionsFolderExists { get; init; }
    public required string SessionsFolderPath { get; init; }
    public required bool OllamaAvailable { get; init; }
    public required bool OllamaCommandVisible { get; init; }
    public required string OllamaExecutablePath { get; init; }
    public required bool OllamaAppAvailable { get; init; }
    public required string OllamaAppPath { get; init; }
    public required string OllamaAppDetail { get; init; }
    public required bool OllamaServerRunning { get; init; }
    public required bool OllamaTrayRunning { get; init; }
    public required int OllamaModelCount { get; init; }
    public required string OllamaModelsSummary { get; init; }
    public required string OllamaDetail { get; init; }
    public required bool LmStudioAvailable { get; init; }
    public required string LmStudioDetail { get; init; }
    public required bool ComfyUiDesktopAvailable { get; init; }
    public required string ComfyUiDesktopDetail { get; init; }
    public required bool PinokioAvailable { get; init; }
    public required string PinokioDetail { get; init; }
    public required bool OpenClawAvailable { get; init; }
    public required string OpenClawDetail { get; init; }
    public required bool OpenClawConfigExists { get; init; }
    public required string OpenClawConfigPath { get; init; }
    public required string OpenClawPrimaryModel { get; init; }
    public required string OpenClawToolProfile { get; init; }
    public required bool OpenClawWebSearchEnabled { get; init; }
    public required bool OpenClawTelegramConfigured { get; init; }
    public required bool OpenClawNodeInstalled { get; init; }
    public required string OpenClawNodeDetail { get; init; }
    public required bool OpenClawBrowserCliAvailable { get; init; }
    public required bool OpenClawBrowserReady { get; init; }
    public required string OpenClawBrowserDetail { get; init; }
    public required IReadOnlyDictionary<string, string> InstalledOllamaModels { get; init; }
}
