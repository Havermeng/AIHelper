namespace LaptopSessionViewer.Models;

public sealed class AppUpdateSnapshot
{
    public required string CurrentVersion { get; init; }
    public required string CurrentVersionDisplay { get; init; }
    public required string LatestVersion { get; init; }
    public required string LatestVersionDisplay { get; init; }
    public string ReleaseTitle { get; init; } = string.Empty;
    public string ReleasePageUrl { get; init; } = string.Empty;
    public string InstallerDownloadUrl { get; init; } = string.Empty;
    public DateTimeOffset? PublishedAtUtc { get; init; }
    public bool IsUpdateAvailable { get; init; }

    public bool HasInstallerAsset => !string.IsNullOrWhiteSpace(InstallerDownloadUrl);
}
