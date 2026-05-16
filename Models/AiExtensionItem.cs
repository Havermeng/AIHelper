namespace LaptopSessionViewer.Models;

public sealed class AiExtensionItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Name { get; set; } = string.Empty;

    public AiExtensionKind Kind { get; set; } = AiExtensionKind.Plugin;

    public string Description { get; set; } = string.Empty;

    public string CommandOrUri { get; set; } = string.Empty;

    public bool IsPreset { get; set; }

    public bool IsInstalled { get; set; }

    public bool IsEnabled { get; set; } = true;

    public bool IsCustom => !IsPreset;

    public string KindLabel => Kind == AiExtensionKind.Mcp ? "MCP" : "Plugin";

    public string SourceLabel => IsPreset ? "Preset" : "Custom";

    public string SourceDisplayLabel { get; set; } = string.Empty;

    public string InstallStateLabel { get; set; } = string.Empty;

    public AiExtensionItem Clone()
    {
        return new AiExtensionItem
        {
            Id = Id,
            Name = Name,
            Kind = Kind,
            Description = Description,
            CommandOrUri = CommandOrUri,
            IsPreset = IsPreset,
            IsInstalled = IsInstalled,
            IsEnabled = IsEnabled,
            SourceDisplayLabel = SourceDisplayLabel,
            InstallStateLabel = InstallStateLabel
        };
    }
}
