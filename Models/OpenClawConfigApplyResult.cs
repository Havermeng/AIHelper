namespace LaptopSessionViewer.Models;

public sealed class OpenClawConfigApplyResult
{
    public required string ModeName { get; init; }
    public required string ConfigPath { get; init; }
    public string BackupPath { get; init; } = string.Empty;
    public required string PrimaryModel { get; init; }
    public required string ToolsProfile { get; init; }
}
