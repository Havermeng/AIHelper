namespace LaptopSessionViewer.Models;

public sealed class AiExtensionItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Name { get; set; } = string.Empty;

    public AiExtensionKind Kind { get; set; } = AiExtensionKind.Plugin;

    public string TargetApp { get; set; } = "Codex";

    public string Description { get; set; } = string.Empty;

    public string CommandOrUri { get; set; } = string.Empty;

    public string DetectionPath { get; set; } = string.Empty;

    public bool IsPreset { get; set; }

    public bool IsDetected { get; set; }

    public bool IsInstalled { get; set; }

    public bool IsEnabled { get; set; } = true;

    public bool IsVerified { get; set; }

    public bool HasVerificationError { get; set; }

    public bool IsBusy { get; set; }

    public bool CanProvision { get; set; }

    public bool CanUninstall { get; set; }

    public string RequestedAccess { get; set; } = string.Empty;

    public string PackageVersion { get; set; } = string.Empty;

    public string VerificationDetail { get; set; } = string.Empty;

    public string ManagementKind { get; set; } = string.Empty;

    public bool IsCustom => !IsPreset && !IsDetected;

    public bool IsActive
    {
        get => IsInstalled && IsEnabled;
        set
        {
            IsEnabled = value;
            if (value)
            {
                IsInstalled = true;
            }
        }
    }

    public string KindLabel => Kind switch
    {
        AiExtensionKind.Mcp => "MCP",
        AiExtensionKind.Skill => "Skill",
        _ => "Plugin"
    };

    public string SourceLabel => IsDetected ? "Detected" : IsPreset ? "Preset" : "Custom";

    public string KindDisplayLabel { get; set; } = string.Empty;

    public string SourceDisplayLabel { get; set; } = string.Empty;

    public string TargetAppDisplayLabel { get; set; } = string.Empty;

    public string InstallStateLabel { get; set; } = string.Empty;

    public AiExtensionItem Clone()
    {
        return new AiExtensionItem
        {
            Id = Id,
            Name = Name,
            Kind = Kind,
            TargetApp = TargetApp,
            Description = Description,
            CommandOrUri = CommandOrUri,
            DetectionPath = DetectionPath,
            IsPreset = IsPreset,
            IsDetected = IsDetected,
            IsInstalled = IsInstalled,
            IsEnabled = IsEnabled,
            IsVerified = IsVerified,
            HasVerificationError = HasVerificationError,
            IsBusy = IsBusy,
            CanProvision = CanProvision,
            CanUninstall = CanUninstall,
            RequestedAccess = RequestedAccess,
            PackageVersion = PackageVersion,
            VerificationDetail = VerificationDetail,
            ManagementKind = ManagementKind,
            KindDisplayLabel = KindDisplayLabel,
            SourceDisplayLabel = SourceDisplayLabel,
            TargetAppDisplayLabel = TargetAppDisplayLabel,
            InstallStateLabel = InstallStateLabel
        };
    }
}
