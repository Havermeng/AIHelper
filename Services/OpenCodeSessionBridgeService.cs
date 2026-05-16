using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class OpenCodeSessionBridgeService
{
    private readonly AppLogService _logService;

    public OpenCodeSessionBridgeService(AppLogService logService)
    {
        _logService = logService;
    }

    public string OpenCodeCommandPath =>
        ResolveOpenCodeCommandPath() ??
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm", "opencode.cmd");

    public bool IsOpenCodeAvailable =>
        IsOpenCodeDesktopAvailable ||
        !string.IsNullOrWhiteSpace(ResolveOpenCodeCommandPath());

    public bool IsOpenCodeDesktopAvailable =>
        !string.IsNullOrWhiteSpace(ResolveOpenCodeDesktopPath());

    public OpenCodeSessionLinkRecord CreateBridge(CodexSessionConversation conversation)
    {
        _logService.Info(
            nameof(OpenCodeSessionBridgeService),
            $"CreateBridge requested. Codex={conversation.SessionId}; Messages={conversation.Messages.Count}; Cwd={conversation.WorkingDirectory}");

        if (conversation.Messages.Count == 0)
        {
            throw new InvalidOperationException("The selected Codex session has no user or assistant messages to bridge.");
        }

        var handoffId = BuildHandoffId(conversation.SessionId);
        var sessionTitle = string.IsNullOrWhiteSpace(conversation.Title)
            ? $"Codex handoff {conversation.SessionId[..Math.Min(12, conversation.SessionId.Length)]}"
            : conversation.Title.Trim();
        var workingDirectory = AiHelperWorkspaceService.ResolveSafeWorkspace(
            conversation.WorkingDirectory,
            conversation.SessionId,
            sessionTitle,
            out var usedFallbackWorkspace);
        var handoffPath = BuildHandoffPath(conversation.SessionId, handoffId);
        var directoryPath = Path.GetDirectoryName(handoffPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        File.WriteAllText(
            handoffPath,
            BuildHandoffMarkdown(conversation, sessionTitle, workingDirectory),
            Encoding.UTF8);

        _logService.Info(
            nameof(OpenCodeSessionBridgeService),
            $"Created OpenCode handoff file. Codex={conversation.SessionId}; Handoff={handoffId}; Cwd={workingDirectory}; FallbackWorkspace={usedFallbackWorkspace}; File={handoffPath}");

        return new OpenCodeSessionLinkRecord
        {
            CodexSessionId = conversation.SessionId,
            OpenCodeSessionId = handoffId,
            OpenCodeTitle = sessionTitle,
            WorkingDirectory = workingDirectory,
            HandoffPath = handoffPath,
            LinkedAtUtc = DateTime.UtcNow,
            CodexUpdatedAtUtc = conversation.UpdatedAtUtc.UtcDateTime,
            CodexMessageCount = conversation.Messages.Count
        };
    }

    public void LaunchSession(OpenCodeSessionLinkRecord linkRecord)
    {
        _logService.Info(
            nameof(OpenCodeSessionBridgeService),
            $"LaunchSession requested. Handoff={linkRecord.OpenCodeSessionId}; Cwd={linkRecord.WorkingDirectory}; File={linkRecord.HandoffPath ?? "-"}");

        if (string.IsNullOrWhiteSpace(linkRecord.HandoffPath) || !File.Exists(linkRecord.HandoffPath))
        {
            throw new FileNotFoundException("OpenCode handoff file was not found. Refresh the OpenCode bridge first.", linkRecord.HandoffPath);
        }

        var workingDirectory = AiHelperWorkspaceService.ResolveSafeWorkspace(
            linkRecord.WorkingDirectory,
            linkRecord.CodexSessionId,
            linkRecord.OpenCodeTitle,
            out var usedFallbackWorkspace);
        var prompt = BuildHandoffPrompt(linkRecord);
        var deepLink = BuildNewSessionDeepLink(workingDirectory, prompt);

        if (usedFallbackWorkspace)
        {
            _logService.Info(
                nameof(OpenCodeSessionBridgeService),
                $"OpenCode launch redirected to safe desktop workspace: {workingDirectory}");
        }

        if (!TryLaunchOpenCodeDesktop(deepLink, workingDirectory))
        {
            throw new FileNotFoundException(
                "OpenCode Desktop was not found. Install the native OpenCode app to continue Codex sessions there.",
                ResolveOpenCodeDesktopPath() ?? "OpenCode.exe");
        }
    }

    private string BuildHandoffPrompt(OpenCodeSessionLinkRecord linkRecord)
    {
        return string.Join(
            Environment.NewLine,
            "AIHelper prepared a Codex session context handoff file.",
            $"Read this file completely first: \"{linkRecord.HandoffPath}\"",
            "Restore the previous context from it, then continue the work from the latest user request.",
            "Do not answer that you cannot see the previous context until you have read the file.");
    }

    private static string BuildHandoffMarkdown(
        CodexSessionConversation conversation,
        string sessionTitle,
        string workingDirectory)
    {
        var builder = new StringBuilder();

        builder.AppendLine("# AIHelper Codex Session Handoff");
        builder.AppendLine();
        builder.AppendLine("This file was generated by AIHelper so OpenCode can continue a Codex CLI session without AIHelper modifying OpenCode's internal database.");
        builder.AppendLine();
        builder.AppendLine("## Instructions for OpenCode");
        builder.AppendLine();
        builder.AppendLine("- Read this file completely before answering.");
        builder.AppendLine("- Treat the transcript below as the previous conversation context.");
        builder.AppendLine("- Continue from the latest user request in the transcript.");
        builder.AppendLine("- If a file path is mentioned, inspect the local workspace before making changes.");
        builder.AppendLine();
        builder.AppendLine("## Metadata");
        builder.AppendLine();
        builder.AppendLine($"- Codex session ID: {conversation.SessionId}");
        builder.AppendLine($"- Title: {sessionTitle}");
        builder.AppendLine($"- Working directory: {workingDirectory}");
        builder.AppendLine($"- Model provider: {NormalizeMetadataValue(conversation.ModelProvider)}");
        builder.AppendLine($"- Started UTC: {FormatTimestamp(conversation.StartedAtUtc)}");
        builder.AppendLine($"- Updated UTC: {FormatTimestamp(conversation.UpdatedAtUtc)}");
        builder.AppendLine($"- Messages: {conversation.Messages.Count}");
        builder.AppendLine();

        var latestUserMessage = conversation.Messages
            .LastOrDefault(message => string.Equals(message.Role, "user", StringComparison.OrdinalIgnoreCase));

        if (latestUserMessage is not null)
        {
            builder.AppendLine("## Latest User Request");
            builder.AppendLine();
            builder.AppendLine(latestUserMessage.Text.Trim());
            builder.AppendLine();
        }

        builder.AppendLine("## Full Transcript");
        builder.AppendLine();

        for (var index = 0; index < conversation.Messages.Count; index++)
        {
            var message = conversation.Messages[index];
            var role = string.IsNullOrWhiteSpace(message.Role) ? "message" : message.Role.Trim();
            builder.AppendLine($"### {index + 1}. {role} ({FormatTimestamp(message.Timestamp)})");
            builder.AppendLine();
            builder.AppendLine(string.IsNullOrWhiteSpace(message.Text) ? "[empty message]" : message.Text.Trim());
            builder.AppendLine();
        }

        return builder.ToString();
    }

    private bool TryLaunchOpenCodeDesktop(string deepLink, string workingDirectory)
    {
        try
        {
            var desktopPath = ResolveOpenCodeDesktopPath();
            if (string.IsNullOrWhiteSpace(desktopPath))
            {
                return false;
            }

            if (!IsOpenCodeDesktopRunning())
            {
                Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = desktopPath,
                        WorkingDirectory = Path.GetDirectoryName(desktopPath) ?? workingDirectory,
                        UseShellExecute = true
                    });

                _ = Task.Run(async () =>
                {
                    await Task.Delay(2500).ConfigureAwait(false);
                    TryLaunchOpenCodeDeepLink(deepLink);
                });

                _logService.Info(
                    nameof(OpenCodeSessionBridgeService),
                    $"OpenCode Desktop started. Handoff deep link will be sent after startup: {deepLink}");
                return true;
            }

            return TryLaunchOpenCodeDeepLink(deepLink);
        }
        catch (Exception ex)
        {
            _logService.Error(
                nameof(OpenCodeSessionBridgeService),
                $"OpenCode Desktop launch failed: {ex.Message}",
                ex);
            return false;
        }
    }

    private bool TryLaunchOpenCodeDeepLink(string deepLink)
    {
        try
        {
            Process.Start(
                new ProcessStartInfo
                {
                    FileName = deepLink,
                    UseShellExecute = true
                });

            _logService.Info(
                nameof(OpenCodeSessionBridgeService),
                $"OpenCode Desktop deep link launched: {deepLink}");
            return true;
        }
        catch (Exception ex)
        {
            _logService.Error(
                nameof(OpenCodeSessionBridgeService),
                $"OpenCode Desktop deep link launch failed: {ex.Message}",
                ex);
            return false;
        }
    }

    private static string BuildNewSessionDeepLink(string workingDirectory, string prompt)
    {
        return string.Concat(
            "opencode://new-session?directory=",
            Uri.EscapeDataString(workingDirectory),
            "&prompt=",
            Uri.EscapeDataString(prompt));
    }

    private static bool IsOpenCodeDesktopRunning()
    {
        try
        {
            foreach (var process in Process.GetProcessesByName("OpenCode"))
            {
                using (process)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    private static string BuildHandoffId(string codexSessionId)
    {
        return $"handoff_{codexSessionId.Replace('-', '_')}_{Guid.NewGuid().ToString("N")[..8]}";
    }

    private static string BuildHandoffPath(string codexSessionId, string handoffId)
    {
        var safeCodexId = Regex.Replace(codexSessionId, "[^a-zA-Z0-9_-]", "_");
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AIHelper",
            "opencode-bridge",
            $"{safeCodexId}-{handoffId}.md");
    }

    private static string NormalizeMetadataValue(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
    }

    private static string FormatTimestamp(DateTimeOffset? value)
    {
        return value is null ? "unknown" : value.Value.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss");
    }

    private static string? ResolveOpenCodeCommandPath()
    {
        var candidates = new[]
        {
            ResolveExecutableOnPath("opencode.cmd"),
            ResolveExecutableOnPath("opencode.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm", "opencode.cmd"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "opencode", "opencode-cli.exe")
        };

        return candidates.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path));
    }

    private static string? ResolveOpenCodeDesktopPath()
    {
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenCode", "OpenCode.exe"),
            ResolveExecutableOnPath("OpenCode.exe")
        };

        return candidates.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path));
    }

    private static string? ResolveExecutableOnPath(string executableName)
    {
        var pathVariable = Environment.GetEnvironmentVariable("PATH");

        if (string.IsNullOrWhiteSpace(pathVariable))
        {
            return null;
        }

        foreach (var directory in pathVariable.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            try
            {
                var fullPath = Path.Combine(directory, executableName);

                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }
            catch
            {
            }
        }

        return null;
    }
}
