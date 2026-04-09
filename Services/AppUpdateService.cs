using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class AppUpdateService
{
    private const string LatestReleaseApiUrl = "https://api.github.com/repos/Havermeng/AIHelper/releases/latest";
    private const string LatestReleasePageUrl = "https://github.com/Havermeng/AIHelper/releases/latest";
    private const string ReleaseDownloadUrlTemplate = "https://github.com/Havermeng/AIHelper/releases/download/{0}/AIHelper-Setup.exe";
    private const string SetupAssetName = "AIHelper-Setup.exe";
    private static readonly Regex HtmlTitleRegex =
        new("<title>\\s*(?<title>[^<]+?)\\s*</title>", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly AppLogService _logService = new();

    public string CurrentVersion => GetCurrentVersion();

    public string CurrentVersionDisplay => $"v{CurrentVersion}";

    public string ReleasePageUrl => LatestReleasePageUrl;

    public async Task<AppUpdateSnapshot> GetLatestReleaseAsync(CancellationToken cancellationToken = default)
    {
        var currentVersion = CurrentVersion;

        try
        {
            using var client = CreateHttpClient();
            return await TryGetLatestReleaseFromApiAsync(client, currentVersion, cancellationToken);
        }
        catch (Exception apiException)
        {
            _logService.Error(
                nameof(AppUpdateService),
                "GitHub API update check failed. Falling back to the public release page.",
                apiException);

            using var fallbackClient = CreateHttpClient();
            var snapshot = await TryGetLatestReleaseFromPageAsync(fallbackClient, currentVersion, cancellationToken);
            _logService.Info(
                nameof(AppUpdateService),
                $"Checked updates through release page fallback. Current={snapshot.CurrentVersionDisplay}, Latest={snapshot.LatestVersionDisplay}, Available={snapshot.IsUpdateAvailable}.");
            return snapshot;
        }
    }

    public async Task DownloadInstallerAsync(
        string installerUrl,
        string destinationFilePath,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(installerUrl))
        {
            throw new InvalidOperationException("Installer URL is empty.");
        }

        var directory = Path.GetDirectoryName(destinationFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var tempFilePath = destinationFilePath + ".download";

        try
        {
            using var client = CreateHttpClient();
            using var response = await client.GetAsync(
                installerUrl,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken);
            response.EnsureSuccessStatusCode();

            await using (var input = await response.Content.ReadAsStreamAsync(cancellationToken))
            await using (var output = File.Create(tempFilePath))
            {
                await input.CopyToAsync(output, cancellationToken);
            }

            File.Move(tempFilePath, destinationFilePath, overwrite: true);
            _logService.Info(nameof(AppUpdateService), $"Downloaded installer to {destinationFilePath}.");
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(AppUpdateService), "Failed to download the installer.", exception);

            try
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
            catch
            {
            }

            throw;
        }
    }

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(10)
        };

        client.DefaultRequestHeaders.UserAgent.ParseAdd("AIHelper-Updater/1.0");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    private async Task<AppUpdateSnapshot> TryGetLatestReleaseFromApiAsync(
        HttpClient client,
        string currentVersion,
        CancellationToken cancellationToken)
    {
        using var response = await client.GetAsync(
            LatestReleaseApiUrl,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        var root = document.RootElement;
        var latestVersion = NormalizeVersion(ReadString(root, "tag_name"));

        if (string.IsNullOrWhiteSpace(latestVersion))
        {
            latestVersion = currentVersion;
        }

        var releasePageUrl = ReadString(root, "html_url");
        var snapshot = new AppUpdateSnapshot
        {
            CurrentVersion = currentVersion,
            CurrentVersionDisplay = $"v{currentVersion}",
            LatestVersion = latestVersion,
            LatestVersionDisplay = $"v{latestVersion}",
            ReleaseTitle = ReadString(root, "name"),
            ReleasePageUrl = string.IsNullOrWhiteSpace(releasePageUrl) ? LatestReleasePageUrl : releasePageUrl,
            InstallerDownloadUrl = GetInstallerDownloadUrl(root),
            PublishedAtUtc = ReadDateTimeOffset(root, "published_at"),
            IsUpdateAvailable = CompareVersions(currentVersion, latestVersion) < 0
        };

        _logService.Info(
            nameof(AppUpdateService),
            $"Checked updates through GitHub API. Current={snapshot.CurrentVersionDisplay}, Latest={snapshot.LatestVersionDisplay}, Available={snapshot.IsUpdateAvailable}.");

        return snapshot;
    }

    private async Task<AppUpdateSnapshot> TryGetLatestReleaseFromPageAsync(
        HttpClient client,
        string currentVersion,
        CancellationToken cancellationToken)
    {
        using var response = await client.GetAsync(
            LatestReleasePageUrl,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);
        response.EnsureSuccessStatusCode();

        var finalUri = response.RequestMessage?.RequestUri;
        var resolvedUrl = finalUri?.AbsoluteUri ?? LatestReleasePageUrl;
        var releaseTag = GetLastUriSegment(finalUri) ?? $"v{currentVersion}";
        var latestVersion = NormalizeVersion(releaseTag);

        if (string.IsNullOrWhiteSpace(latestVersion))
        {
            latestVersion = currentVersion;
            releaseTag = $"v{currentVersion}";
        }

        var html = await response.Content.ReadAsStringAsync(cancellationToken);
        var releaseTitle = ParseHtmlTitle(html);

        return new AppUpdateSnapshot
        {
            CurrentVersion = currentVersion,
            CurrentVersionDisplay = $"v{currentVersion}",
            LatestVersion = latestVersion,
            LatestVersionDisplay = $"v{latestVersion}",
            ReleaseTitle = string.IsNullOrWhiteSpace(releaseTitle) ? releaseTag : releaseTitle,
            ReleasePageUrl = resolvedUrl,
            InstallerDownloadUrl = string.Format(ReleaseDownloadUrlTemplate, releaseTag),
            PublishedAtUtc = null,
            IsUpdateAvailable = CompareVersions(currentVersion, latestVersion) < 0
        };
    }

    private static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        var normalizedInformationalVersion = NormalizeVersion(informationalVersion);

        if (!string.IsNullOrWhiteSpace(normalizedInformationalVersion))
        {
            return normalizedInformationalVersion;
        }

        var executablePath = Environment.ProcessPath;

        if (!string.IsNullOrWhiteSpace(executablePath) && File.Exists(executablePath))
        {
            var productVersion = NormalizeVersion(FileVersionInfo.GetVersionInfo(executablePath).ProductVersion);

            if (!string.IsNullOrWhiteSpace(productVersion))
            {
                return productVersion;
            }
        }

        var version = assembly.GetName().Version;
        return version is null
            ? "0.0.0"
            : $"{version.Major}.{Math.Max(version.Minor, 0)}.{Math.Max(version.Build, 0)}";
    }

    private static string GetInstallerDownloadUrl(JsonElement root)
    {
        if (!root.TryGetProperty("assets", out var assetsElement) || assetsElement.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        foreach (var asset in assetsElement.EnumerateArray())
        {
            var name = ReadString(asset, "name");
            if (string.Equals(name, SetupAssetName, StringComparison.OrdinalIgnoreCase))
            {
                return ReadString(asset, "browser_download_url");
            }
        }

        foreach (var asset in assetsElement.EnumerateArray())
        {
            var name = ReadString(asset, "name");
            if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) &&
                name.Contains("setup", StringComparison.OrdinalIgnoreCase))
            {
                return ReadString(asset, "browser_download_url");
            }
        }

        return string.Empty;
    }

    private static string ReadString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property) || property.ValueKind == JsonValueKind.Null)
        {
            return string.Empty;
        }

        return property.GetString() ?? string.Empty;
    }

    private static DateTimeOffset? ReadDateTimeOffset(JsonElement element, string propertyName)
    {
        var value = ReadString(element, propertyName);
        return DateTimeOffset.TryParse(value, out var parsed) ? parsed : null;
    }

    private static int CompareVersions(string left, string right)
    {
        return TryParseVersion(left, out var leftVersion) && TryParseVersion(right, out var rightVersion)
            ? leftVersion.CompareTo(rightVersion)
            : StringComparer.OrdinalIgnoreCase.Compare(left, right);
    }

    private static bool TryParseVersion(string value, out Version version)
    {
        var normalized = NormalizeVersion(value);

        if (Version.TryParse(normalized, out version!))
        {
            return true;
        }

        var parts = normalized.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 1 && int.TryParse(parts[0], out var major))
        {
            version = new Version(major, 0, 0);
            return true;
        }

        if (parts.Length == 2 && int.TryParse(parts[0], out major) && int.TryParse(parts[1], out var minor))
        {
            version = new Version(major, minor, 0);
            return true;
        }

        version = new Version(0, 0, 0);
        return false;
    }

    private static string NormalizeVersion(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim();
        if (normalized.StartsWith('v') || normalized.StartsWith('V'))
        {
            normalized = normalized[1..];
        }

        var separatorIndex = normalized.IndexOfAny(new[] { '-', '+' });
        if (separatorIndex >= 0)
        {
            normalized = normalized[..separatorIndex];
        }

        return normalized.Trim();
    }

    private static string ParseHtmlTitle(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var match = HtmlTitleRegex.Match(html);
        if (!match.Success)
        {
            return string.Empty;
        }

        var title = match.Groups["title"].Value.Trim();
        return title
            .Replace("· Havermeng/AIHelper · GitHub", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private static string? GetLastUriSegment(Uri? uri)
    {
        if (uri is null)
        {
            return null;
        }

        var segments = uri.AbsolutePath
            .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return segments.Length == 0 ? null : segments[^1];
    }
}



