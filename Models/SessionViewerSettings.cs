namespace LaptopSessionViewer.Models;

public sealed class SessionViewerSettings
{
    public AppLanguage Language { get; set; } = AppLanguage.English;

    public bool DefaultDangerousFullAccess { get; set; }

    public bool PhotoPasteFixEnabled { get; set; }

    public bool BeginnerModeEnabled { get; set; } = true;

    public bool HasCompletedBeginnerOnboarding { get; set; }
}
