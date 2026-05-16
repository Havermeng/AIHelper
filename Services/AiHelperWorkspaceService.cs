using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace LaptopSessionViewer.Services;

public static class AiHelperWorkspaceService
{
    private const string WorkspaceRootFolderName = "AIHelper Projects";

    public static string ResolveSafeWorkspace(
        string? requestedWorkingDirectory,
        string sessionId,
        string title,
        out bool usedFallback)
    {
        var normalizedDirectory = NormalizeDirectory(requestedWorkingDirectory);

        if (!string.IsNullOrWhiteSpace(normalizedDirectory) &&
            Directory.Exists(normalizedDirectory) &&
            !IsUnsafeWorkspace(normalizedDirectory))
        {
            usedFallback = false;
            return normalizedDirectory;
        }

        usedFallback = true;
        var workspace = BuildDesktopWorkspacePath(sessionId, title);
        Directory.CreateDirectory(workspace);
        EnsureReadme(workspace, requestedWorkingDirectory, sessionId, title);
        return workspace;
    }

    public static bool IsUnsafeWorkspace(string? directory)
    {
        var normalizedDirectory = NormalizeDirectory(directory);
        if (string.IsNullOrWhiteSpace(normalizedDirectory))
        {
            return true;
        }

        var systemDirectory = NormalizeDirectory(Environment.SystemDirectory);
        var windowsDirectory = NormalizeDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
        var programFiles = NormalizeDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
        var programFilesX86 = NormalizeDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
        var driveRoot = Path.GetPathRoot(normalizedDirectory);

        return IsSameOrChild(normalizedDirectory, systemDirectory) ||
               IsSameOrChild(normalizedDirectory, windowsDirectory) ||
               IsSameOrChild(normalizedDirectory, programFiles) ||
               IsSameOrChild(normalizedDirectory, programFilesX86) ||
               string.Equals(normalizedDirectory.TrimEnd(Path.DirectorySeparatorChar), driveRoot?.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildDesktopWorkspacePath(string sessionId, string title)
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var normalizedSessionId = string.IsNullOrWhiteSpace(sessionId)
            ? string.Empty
            : Regex.Replace(sessionId, "[^a-zA-Z0-9]", string.Empty);
        var shortId = string.IsNullOrWhiteSpace(normalizedSessionId)
            ? Guid.NewGuid().ToString("N")[..8]
            : normalizedSessionId[..Math.Min(8, normalizedSessionId.Length)];
        var safeTitle = Slugify(string.IsNullOrWhiteSpace(title) ? "Codex Session" : title);
        return Path.Combine(desktop, WorkspaceRootFolderName, $"{safeTitle}-{shortId}");
    }

    private static void EnsureReadme(string workspace, string? originalWorkingDirectory, string sessionId, string title)
    {
        var readmePath = Path.Combine(workspace, "README_AIHELPER_PROJECT.txt");
        if (File.Exists(readmePath))
        {
            return;
        }

        var text = string.Join(
            Environment.NewLine,
            "AIHelper project workspace",
            string.Empty,
            "This folder was created because the original session working directory was unsafe or too broad for AI tools.",
            $"Original working directory: {NormalizeDirectory(originalWorkingDirectory) ?? "-"}",
            $"Codex session ID: {sessionId}",
            $"Session title: {title}",
            string.Empty,
            "Put project files here if you want Codex/OpenCode to work with a clean, small workspace.");

        File.WriteAllText(readmePath, text, Encoding.UTF8);
    }

    private static string Slugify(string value)
    {
        var normalized = Regex.Replace(value.Trim().ToLowerInvariant(), "[^a-z0-9а-яё]+", "-").Trim('-');
        if (string.IsNullOrWhiteSpace(normalized))
        {
            normalized = "codex-session";
        }

        return normalized.Length <= 48 ? normalized : normalized[..48].Trim('-');
    }

    private static string? NormalizeDirectory(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory) || directory.Trim() == "-")
        {
            return null;
        }

        try
        {
            return Path.GetFullPath(directory.Trim()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch
        {
            return null;
        }
    }

    private static bool IsSameOrChild(string directory, string? parent)
    {
        if (string.IsNullOrWhiteSpace(parent))
        {
            return false;
        }

        var normalizedParent = parent.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return string.Equals(directory, normalizedParent, StringComparison.OrdinalIgnoreCase) ||
               directory.StartsWith(normalizedParent + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }
}
