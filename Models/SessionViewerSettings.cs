namespace LaptopSessionViewer.Models;

public sealed class SessionViewerSettings
{
    public AppLanguage Language { get; set; } = AppLanguage.English;

    public bool DefaultDangerousFullAccess { get; set; }
}
