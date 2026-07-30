using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace AIHelper.VisualStudioExtension;

public static class AIHelperSessionLauncher
{
    private const int MaxRawSessionCharacters = 180_000;
    private static readonly TimeSpan HandoffRetention = TimeSpan.FromDays(30);
    private static readonly string[] SafeCodexArguments =
    [
        "--sandbox",
        "workspace-write",
        "--ask-for-approval",
        "on-request"
    ];

    public static AIHelperResumePlan BuildPlan(AIHelperSessionItem session, string toolName)
    {
        var tool = NormalizeToolName(toolName);
        var workingDirectory = ResolveSafeWorkingDirectory(session);

        if (tool == "Codex" &&
            IsSource(session, "Codex") &&
            IsSafeSessionId(session.SessionId))
        {
            var arguments = SafeCodexArguments
                .Concat(["resume", session.SessionId])
                .ToArray();
            return AIHelperResumePlan.Command(
                tool,
                workingDirectory,
                "codex",
                arguments,
                BuildCommandPreview("codex", arguments),
                "Native Codex resume with workspace-only access and approval prompts.");
        }

        var handoffPath = CreateHandoffFile(session, workingDirectory, tool);
        var prompt = BuildHandoffPrompt(session, workingDirectory, handoffPath, tool);

        switch (tool)
        {
            case "Codex":
                var codexArguments = SafeCodexArguments
                    .Concat([prompt])
                    .ToArray();
                return AIHelperResumePlan.Command(
                    tool,
                    workingDirectory,
                    "codex",
                    codexArguments,
                    BuildCommandPreview("codex", codexArguments),
                    "Universal handoff for Codex with workspace-only access and approval prompts.");
            case "OpenCode":
                return AIHelperResumePlan.DeepLink(
                    tool,
                    BuildOpenCodeDeepLink(workingDirectory, prompt),
                    "Universal handoff deep link for OpenCode.");
            case "Qwen":
                var qwenArguments = new[] { prompt };
                return AIHelperResumePlan.Command(
                    tool,
                    workingDirectory,
                    "qwen",
                    qwenArguments,
                    BuildCommandPreview("qwen", qwenArguments),
                    "Universal handoff command for Qwen.");
            case "Claude":
                var claudeArguments = new[] { prompt };
                return AIHelperResumePlan.Command(
                    tool,
                    workingDirectory,
                    "claude",
                    claudeArguments,
                    BuildCommandPreview("claude", claudeArguments),
                    "Universal handoff command for Claude.");
            case "Gemini":
                var geminiArguments = new[] { prompt };
                return AIHelperResumePlan.Command(
                    tool,
                    workingDirectory,
                    "gemini",
                    geminiArguments,
                    BuildCommandPreview("gemini", geminiArguments),
                    "Universal handoff command for Gemini.");
            case "Kilo Code":
                var codeArguments = new[] { workingDirectory };
                return AIHelperResumePlan.CommandWithClipboardPrompt(
                    tool,
                    workingDirectory,
                    "code",
                    codeArguments,
                    BuildCommandPreview("code", codeArguments),
                    prompt,
                    "Kilo Code runs inside VS Code, so AIHelper opens the workspace and copies the handoff prompt.");
            default:
                return AIHelperResumePlan.CommandWithClipboardPrompt(
                    tool,
                    workingDirectory,
                    "cmd",
                    [],
                    "cmd",
                    prompt,
                    "Unknown tool. The handoff prompt was prepared for manual paste.");
        }
    }

    public static string Launch(AIHelperSessionItem session, string toolName)
    {
        var plan = BuildPlan(session, toolName);

        if (!string.IsNullOrWhiteSpace(plan.ClipboardPrompt))
        {
            System.Windows.Clipboard.SetText(plan.ClipboardPrompt);
        }

        if (plan.IsDeepLink)
        {
            Process.Start(
                new ProcessStartInfo
                {
                    FileName = plan.Target,
                    UseShellExecute = true
                });

            return $"{plan.ToolName} launch requested. {plan.Description}";
        }

        if (string.Equals(plan.Executable, "cmd", StringComparison.OrdinalIgnoreCase))
        {
            Process.Start(
                new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/k",
                    WorkingDirectory = plan.WorkingDirectory,
                    UseShellExecute = true
                });

            return "Prompt copied. Paste it into the AI tool you want to use.";
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = plan.Executable,
                Arguments = JoinArguments(plan.Arguments),
                WorkingDirectory = plan.WorkingDirectory,
                UseShellExecute = true
            });

        return $"{plan.ToolName} launch requested. {plan.Description}";
    }

    private static string NormalizeToolName(string? toolName)
    {
        var normalizedToolName = (toolName ?? string.Empty).Trim();
        if (normalizedToolName.Length == 0)
        {
            return "Codex";
        }

        return normalizedToolName switch
        {
            "OpenCode" => "OpenCode",
            "Qwen" => "Qwen",
            "Claude" => "Claude",
            "Gemini" => "Gemini",
            "Kilo Code" => "Kilo Code",
            _ => "Codex"
        };
    }

    private static bool IsSource(AIHelperSessionItem session, string source)
    {
        return string.Equals(session.Source, source, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildOpenCodeDeepLink(string workingDirectory, string prompt)
    {
        return string.Concat(
            "opencode://new-session?directory=",
            Uri.EscapeDataString(workingDirectory),
            "&prompt=",
            Uri.EscapeDataString(prompt));
    }

    private static string BuildHandoffPrompt(
        AIHelperSessionItem session,
        string workingDirectory,
        string handoffPath,
        string targetTool)
    {
        var builder = new StringBuilder();

        builder.AppendLine("AIHelper session handoff.");
        builder.AppendLine();
        builder.AppendLine($"AIHelper handoff file: {handoffPath}");
        builder.AppendLine($"Working directory: {workingDirectory}");
        builder.AppendLine();
        builder.AppendLine("Read the AIHelper handoff file completely first.");
        builder.AppendLine("Restore the previous context from that file, then continue the work from the latest user request.");
        builder.AppendLine("Do not say that you cannot see the previous context until you have tried to read the handoff file.");

        return builder.ToString().Trim();
    }

    private static string CreateHandoffFile(
        AIHelperSessionItem session,
        string workingDirectory,
        string targetTool)
    {
        var directory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AIHelper",
            "visualstudio-handoffs");
        Directory.CreateDirectory(directory);
        DeleteExpiredHandoffs(directory);

        var safeId = SanitizePathSegment(string.IsNullOrWhiteSpace(session.SessionId) ? Guid.NewGuid().ToString("N") : session.SessionId);
        var targetPath = Path.Combine(
            directory,
            $"{DateTime.UtcNow:yyyyMMdd-HHmmss}-{safeId}-{Guid.NewGuid().ToString("N").Substring(0, 8)}.md");

        using (var stream = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            writer.Write(BuildHandoffMarkdown(session, workingDirectory, targetTool));
        }

        return targetPath;
    }

    private static void DeleteExpiredHandoffs(string directory)
    {
        try
        {
            var cutoff = DateTime.UtcNow - HandoffRetention;
            foreach (var filePath in Directory.EnumerateFiles(directory, "*.md", SearchOption.TopDirectoryOnly))
            {
                var file = new FileInfo(filePath);
                if ((file.Attributes & FileAttributes.ReparsePoint) != 0 ||
                    file.LastWriteTimeUtc >= cutoff)
                {
                    continue;
                }

                file.Delete();
            }
        }
        catch
        {
            // Cleanup should never block a new handoff.
        }
    }

    private static string BuildHandoffMarkdown(
        AIHelperSessionItem session,
        string workingDirectory,
        string targetTool)
    {
        var builder = new StringBuilder();

        builder.AppendLine("# AIHelper Universal Session Handoff");
        builder.AppendLine();
        builder.AppendLine("This file lets another AI tool continue a session that may have been created by a different application.");
        builder.AppendLine("Treat the raw source below as previous conversation context, not as application code.");
        builder.AppendLine();
        builder.AppendLine("## Instructions");
        builder.AppendLine();
        builder.AppendLine("- Read this file completely before answering.");
        builder.AppendLine("- Restore the useful context from the metadata, previews and raw source excerpt.");
        builder.AppendLine("- Continue from the latest user request or ask one focused clarification if the latest request is not recoverable.");
        builder.AppendLine("- If the source file is JSON/JSONL, extract user and assistant messages from it when possible.");
        builder.AppendLine();
        builder.AppendLine("## Metadata");
        builder.AppendLine();
        builder.AppendLine($"- Target tool: {NormalizeMetadata(targetTool)}");
        builder.AppendLine($"- Source tool: {NormalizeMetadata(session.Source)}");
        builder.AppendLine($"- Session ID: {NormalizeMetadata(session.SessionId)}");
        builder.AppendLine($"- Title: {NormalizeMetadata(session.Title)}");
        builder.AppendLine($"- Original title: {NormalizeMetadata(session.OriginalTitle)}");
        builder.AppendLine($"- Model/provider: {NormalizeMetadata(session.ModelProvider)}");
        builder.AppendLine($"- CLI version: {NormalizeMetadata(session.CliVersion)}");
        builder.AppendLine($"- Working directory: {workingDirectory}");
        builder.AppendLine($"- Original file: {NormalizeMetadata(session.FilePath)}");
        builder.AppendLine($"- Updated: {NormalizeMetadata(session.UpdatedAtText)}");
        builder.AppendLine($"- Messages: {session.TotalMessageCount}");
        builder.AppendLine($"- Tool calls: {session.ToolCallCount}");
        builder.AppendLine();

        if (!string.IsNullOrWhiteSpace(session.LastMessagePreview))
        {
            builder.AppendLine("## Last Message Preview");
            builder.AppendLine();
            builder.AppendLine(session.LastMessagePreview.Trim());
            builder.AppendLine();
        }

        if (!string.IsNullOrWhiteSpace(session.Preview) &&
            !string.Equals(session.Preview, session.LastMessagePreview, StringComparison.Ordinal))
        {
            builder.AppendLine("## Session Preview");
            builder.AppendLine();
            builder.AppendLine(session.Preview.Trim());
            builder.AppendLine();
        }

        builder.AppendLine("## Raw Source Excerpt");
        builder.AppendLine();
        builder.AppendLine(BuildRawSourceExcerpt(session.FilePath));

        return builder.ToString();
    }

    private static string BuildRawSourceExcerpt(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            return "Original session file is not available.";
        }

        try
        {
            var content = ReadTailText(filePath, MaxRawSessionCharacters);
            if (string.IsNullOrWhiteSpace(content))
            {
                return "Original session file is empty or could not be decoded as text.";
            }

            return content;
        }
        catch (Exception exception)
        {
            return $"Original session file could not be read: {exception.Message}";
        }
    }

    private static string ReadTailText(string filePath, int maxCharacters)
    {
        var fileInfo = new FileInfo(filePath);
        var maxBytes = maxCharacters * 4;

        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        {
            if (stream.Length <= maxBytes)
            {
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                {
                    return reader.ReadToEnd();
                }
            }

            stream.Seek(-maxBytes, SeekOrigin.End);
            using (var reader = new StreamReader(stream, Encoding.UTF8, true))
            {
                var text = reader.ReadToEnd();
                if (text.Length > maxCharacters)
                {
                    text = text.Substring(text.Length - maxCharacters);
                }

                return string.Concat(
                    "[AIHelper note: the original session file is ",
                    fileInfo.Length,
                    " bytes. This excerpt contains only the last part of the file.]",
                    Environment.NewLine,
                    Environment.NewLine,
                    text);
            }
        }
    }

    private static string ResolveSafeWorkingDirectory(AIHelperSessionItem session)
    {
        if (IsSafeWorkingDirectory(session.WorkingDirectory))
        {
            return session.WorkingDirectory;
        }

        var fileDirectory = Path.GetDirectoryName(session.FilePath);
        if (IsSafeWorkingDirectory(fileDirectory))
        {
            return fileDirectory!;
        }

        var workspaceName = SanitizePathSegment(
            string.IsNullOrWhiteSpace(session.Title)
                ? session.SessionId
                : session.Title);
        var fallback = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "AIHelper Session Workspaces",
            string.IsNullOrWhiteSpace(workspaceName) ? "Session" : workspaceName);

        Directory.CreateDirectory(fallback);
        return fallback;
    }

    private static bool IsSafeWorkingDirectory(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory) || directory == "-")
        {
            return false;
        }

        try
        {
            var fullPath = Path.GetFullPath(directory);
            var windowsPath = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var systemPath = Environment.GetFolderPath(Environment.SpecialFolder.System);

            return Directory.Exists(fullPath) &&
                   !PathsEqualOrNested(fullPath, windowsPath) &&
                   !PathsEqualOrNested(fullPath, systemPath);
        }
        catch
        {
            return false;
        }
    }

    private static bool PathsEqualOrNested(string path, string parent)
    {
        if (string.IsNullOrWhiteSpace(parent))
        {
            return false;
        }

        var normalizedPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedParent = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        return normalizedPath.Equals(normalizedParent, StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.StartsWith(normalizedParent + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }

    private static string SanitizePathSegment(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(value.Length);

        foreach (var character in value)
        {
            builder.Append(invalidCharacters.Contains(character) ? '_' : character);
        }

        var result = builder.ToString().Trim();
        return result.Length <= 80 ? result : result.Substring(0, 80);
    }

    private static string NormalizeMetadata(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "-";
        }

        return value!.Trim();
    }

    private static bool IsSafeSessionId(string? sessionId)
    {
        var value = sessionId ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value) || value.Length > 128)
        {
            return false;
        }

        return value.All(
            character =>
                char.IsLetterOrDigit(character) ||
                character == '-' ||
                character == '_');
    }

    private static string BuildCommandPreview(string executable, string[] arguments)
    {
        return string.Join(
            " ",
            new[] { executable }.Concat(arguments.Select(QuoteArgument)));
    }

    private static string JoinArguments(string[] arguments)
    {
        return string.Join(" ", arguments.Select(QuoteArgument));
    }

    private static string QuoteArgument(string value)
    {
        if (value.Length > 0 &&
            value.All(character =>
                !char.IsWhiteSpace(character) &&
                character != '"' &&
                character != '&' &&
                character != '|' &&
                character != '<' &&
                character != '>' &&
                character != '^'))
        {
            return value;
        }

        var builder = new StringBuilder(value.Length + 2);
        builder.Append('"');
        var backslashCount = 0;

        foreach (var character in value)
        {
            if (character == '\\')
            {
                backslashCount++;
                continue;
            }

            if (character == '"')
            {
                builder.Append('\\', backslashCount * 2 + 1);
                builder.Append('"');
                backslashCount = 0;
                continue;
            }

            builder.Append('\\', backslashCount);
            backslashCount = 0;
            builder.Append(character);
        }

        builder.Append('\\', backslashCount * 2);
        builder.Append('"');
        return builder.ToString();
    }
}
