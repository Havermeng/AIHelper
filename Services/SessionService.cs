using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class SessionService
{
    private static readonly string CodexHomePath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex");

    private static readonly string SessionsRootPath = Path.Combine(CodexHomePath, "sessions");
    private static readonly string SessionIndexPath = Path.Combine(CodexHomePath, "session_index.jsonl");
    private static readonly string HistoryPath = Path.Combine(CodexHomePath, "history.jsonl");
    private static readonly Regex SessionIdRegex = new(
        @"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})",
        RegexOptions.Compiled);

    public IReadOnlyList<SessionRecord> GetSessions(AppLanguage language = AppLanguage.English)
    {
        if (!Directory.Exists(SessionsRootPath))
        {
            throw new DirectoryNotFoundException($"Codex sessions folder not found: {SessionsRootPath}");
        }

        var titleLookup = LoadThreadTitles();
        var files = new DirectoryInfo(SessionsRootPath)
            .EnumerateFiles("*.jsonl", SearchOption.AllDirectories)
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .ToList();

        var sessions = new List<SessionRecord>(files.Count);

        foreach (var file in files)
        {
            try
            {
                var session = ParseSessionFile(file, titleLookup, language);

                if (session is not null)
                {
                    sessions.Add(session);
                }
            }
            catch (IOException)
            {
                sessions.Add(CreateLockedSessionRecord(file, language));
            }
            catch (UnauthorizedAccessException)
            {
                sessions.Add(CreateLockedSessionRecord(file, language));
            }
        }

        return sessions
            .OrderByDescending(session => session.UpdatedAtUtc)
            .ThenBy(session => session.Title, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public void DeleteSession(SessionRecord session)
    {
        if (!File.Exists(session.FilePath))
        {
            throw new FileNotFoundException("Session file not found.", session.FilePath);
        }

        File.Delete(session.FilePath);
        RemoveJsonlEntries(SessionIndexPath, root => GetString(root, "id") != session.SessionId);
        RemoveJsonlEntries(HistoryPath, root => GetString(root, "session_id") != session.SessionId);
        DeleteEmptyParentDirectories(Path.GetDirectoryName(session.FilePath));
    }

    private static Dictionary<string, string> LoadThreadTitles()
    {
        var titles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(SessionIndexPath))
        {
            return titles;
        }

        foreach (var line in File.ReadLines(SessionIndexPath, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                var id = GetString(root, "id");
                var threadName = GetString(root, "thread_name");

                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(threadName))
                {
                    titles[id] = threadName.Trim();
                }
            }
            catch (JsonException)
            {
            }
        }

        return titles;
    }

    private static SessionRecord? ParseSessionFile(
        FileInfo file,
        IReadOnlyDictionary<string, string> titleLookup,
        AppLanguage language)
    {
        string? sessionId = null;
        string? titleFromIndex = null;
        string firstPrompt = string.Empty;
        string preview = string.Empty;
        string lastMessage = string.Empty;
        string workingDirectory = string.Empty;
        string source = string.Empty;
        string modelProvider = string.Empty;
        string cliVersion = string.Empty;
        DateTimeOffset? startedAt = null;
        var userMessageCount = 0;
        var assistantMessageCount = 0;
        var toolCallCount = 0;
        var transcript = new StringBuilder();

        foreach (var line in File.ReadLines(file.FullName, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                var recordType = GetString(root, "type");
                var lineTimestamp = ParseTimestamp(GetString(root, "timestamp"));

                if (recordType == "session_meta" && TryGetProperty(root, "payload", out var sessionPayload))
                {
                    sessionId = GetString(sessionPayload, "id");
                    titleLookup.TryGetValue(sessionId ?? string.Empty, out titleFromIndex);
                    startedAt = ParseTimestamp(GetString(sessionPayload, "timestamp")) ?? lineTimestamp;
                    workingDirectory = GetString(sessionPayload, "cwd");
                    source = GetString(sessionPayload, "source");
                    modelProvider = GetString(sessionPayload, "model_provider");
                    cliVersion = GetString(sessionPayload, "cli_version");
                    continue;
                }

                if (recordType != "response_item" || !TryGetProperty(root, "payload", out var payload))
                {
                    continue;
                }

                var payloadType = GetString(payload, "type");

                if (payloadType == "function_call" || payloadType == "web_search_call")
                {
                    toolCallCount++;
                    continue;
                }

                if (payloadType != "message")
                {
                    continue;
                }

                var role = GetString(payload, "role");

                if (!string.Equals(role, "user", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var text = ExtractMessageText(payload);
                text = RemoveEnvironmentContext(text);

                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (string.Equals(role, "user", StringComparison.OrdinalIgnoreCase))
                {
                    userMessageCount++;

                    if (string.IsNullOrWhiteSpace(firstPrompt))
                    {
                        firstPrompt = text;
                        preview = TrimPreview(text, 180);
                    }
                }
                else
                {
                    assistantMessageCount++;
                }

                lastMessage = text;
                AppendTranscriptLine(transcript, lineTimestamp, role, text);
            }
            catch (JsonException)
            {
            }
        }

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            sessionId = DeriveSessionId(file.Name);
        }

        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);
        var startedLocal = startedAt?.ToLocalTime();
        var updatedLocal = updatedAt.ToLocalTime();
        var title = ChooseTitle(titleFromIndex, firstPrompt, file.Name);
        var transcriptText = transcript.Length == 0
            ? GetLocalizedText(language, "NoTranscriptFound")
            : transcript.ToString().Trim();
        var totalMessageCount = userMessageCount + assistantMessageCount;
        var unknownText = GetLocalizedText(language, "Unknown");

        var baseSearchBlob = BuildSearchBlob(
            title,
            sessionId,
            workingDirectory,
            preview,
            lastMessage,
            modelProvider,
            source);

        return new SessionRecord
        {
            SessionId = sessionId,
            Title = title,
            Preview = string.IsNullOrWhiteSpace(preview) ? GetLocalizedText(language, "NoPromptPreview") : preview,
            LastMessagePreview = string.IsNullOrWhiteSpace(lastMessage)
                ? GetLocalizedText(language, "NoRecentMessage")
                : TrimPreview(lastMessage, 220),
            StartedAtText = startedLocal?.ToString("dd.MM.yyyy HH:mm:ss") ?? unknownText,
            UpdatedAtText = updatedLocal.ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = FormatDuration(language, startedAt, updatedAt),
            WorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory) ? "-" : workingDirectory,
            Source = string.IsNullOrWhiteSpace(source) ? "-" : source,
            ModelProvider = string.IsNullOrWhiteSpace(modelProvider) ? "-" : modelProvider,
            CliVersion = string.IsNullOrWhiteSpace(cliVersion) ? "-" : cliVersion,
            FilePath = file.FullName,
            RelativePath = Path.GetRelativePath(CodexHomePath, file.FullName),
            TranscriptText = transcriptText,
            UserMessageCount = userMessageCount,
            AssistantMessageCount = assistantMessageCount,
            ToolCallCount = toolCallCount,
            TotalMessageCount = totalMessageCount,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
    }

    private static SessionRecord CreateLockedSessionRecord(FileInfo file, AppLanguage language)
    {
        var sessionId = DeriveSessionId(file.Name);
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);

        var baseSearchBlob = BuildSearchBlob(sessionId, file.FullName, "locked");

        return new SessionRecord
        {
            SessionId = sessionId,
            Title = GetLocalizedText(language, "LockedTitle"),
            Preview = GetLocalizedText(language, "LockedPreview"),
            LastMessagePreview = GetLocalizedText(language, "LockedLastMessage"),
            StartedAtText = GetLocalizedText(language, "Unknown"),
            UpdatedAtText = updatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = GetLocalizedText(language, "Unknown"),
            WorkingDirectory = "-",
            Source = "-",
            ModelProvider = "-",
            CliVersion = "-",
            FilePath = file.FullName,
            RelativePath = Path.GetRelativePath(CodexHomePath, file.FullName),
            TranscriptText = GetLocalizedText(language, "LockedTranscript"),
            UserMessageCount = 0,
            AssistantMessageCount = 0,
            ToolCallCount = 0,
            TotalMessageCount = 0,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
    }

    private static void DeleteEmptyParentDirectories(string? directoryPath)
    {
        while (!string.IsNullOrWhiteSpace(directoryPath) &&
               directoryPath.StartsWith(SessionsRootPath, StringComparison.OrdinalIgnoreCase) &&
               !string.Equals(directoryPath, SessionsRootPath, StringComparison.OrdinalIgnoreCase))
        {
            if (Directory.Exists(directoryPath) &&
                !Directory.EnumerateFileSystemEntries(directoryPath).Any())
            {
                Directory.Delete(directoryPath);
                directoryPath = Path.GetDirectoryName(directoryPath);
                continue;
            }

            break;
        }
    }

    private static string DeriveSessionId(string fileName)
    {
        var fileBaseName = Path.GetFileNameWithoutExtension(fileName);
        var matches = SessionIdRegex.Matches(fileBaseName);

        if (matches.Count > 0)
        {
            return matches[^1].Value;
        }

        return fileBaseName;
    }

    private static void AppendTranscriptLine(
        StringBuilder transcript,
        DateTimeOffset? timestamp,
        string role,
        string text)
    {
        if (transcript.Length > 0)
        {
            transcript.AppendLine();
            transcript.AppendLine();
        }

        var clock = timestamp?.ToLocalTime().ToString("HH:mm:ss") ?? "--:--:--";
        transcript.Append('[').Append(clock).Append("] ").Append(role.ToUpperInvariant()).AppendLine();
        transcript.Append(text.Trim());
    }

    private static string ExtractMessageText(JsonElement payload)
    {
        if (!TryGetProperty(payload, "content", out var content) || content.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();

        foreach (var item in content.EnumerateArray())
        {
            var type = GetString(item, "type");

            if ((type == "input_text" || type == "output_text" || type == "text") &&
                TryGetProperty(item, "text", out var textNode) &&
                textNode.ValueKind == JsonValueKind.String)
            {
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.Append(textNode.GetString());
                continue;
            }

            if (type is "input_image" or "image")
            {
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.Append("[image]");
            }
        }

        return builder.ToString().Trim();
    }

    private static string RemoveEnvironmentContext(string text)
    {
        const string openTag = "<environment_context>";
        const string closeTag = "</environment_context>";
        var start = text.IndexOf(openTag, StringComparison.OrdinalIgnoreCase);
        var end = text.IndexOf(closeTag, StringComparison.OrdinalIgnoreCase);

        if (start < 0 || end <= start)
        {
            return text.Trim();
        }

        var before = text[..start];
        var after = text[(end + closeTag.Length)..];
        return $"{before}\n{after}".Trim();
    }

    private static string ChooseTitle(string? threadName, string firstPrompt, string fileName)
    {
        if (!string.IsNullOrWhiteSpace(threadName))
        {
            return TrimPreview(threadName.Trim(), 90);
        }

        if (!string.IsNullOrWhiteSpace(firstPrompt))
        {
            return TrimPreview(firstPrompt.Trim(), 90);
        }

        return Path.GetFileNameWithoutExtension(fileName);
    }

    private static string BuildSearchBlob(params string?[] values)
    {
        return string.Join(" ", values.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static void RemoveJsonlEntries(string filePath, Func<JsonElement, bool> keepPredicate)
    {
        if (!File.Exists(filePath))
        {
            return;
        }

        var tempPath = $"{filePath}.tmp";

        try
        {
            using var writer = new StreamWriter(tempPath, false, Encoding.UTF8);

            foreach (var line in File.ReadLines(filePath, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var keepLine = true;

                try
                {
                    using var document = JsonDocument.Parse(line);
                    keepLine = keepPredicate(document.RootElement);
                }
                catch (JsonException)
                {
                    keepLine = true;
                }

                if (keepLine)
                {
                    writer.WriteLine(line);
                }
            }

            writer.Flush();
            File.Copy(tempPath, filePath, true);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    private static string FormatDuration(AppLanguage language, DateTimeOffset? startedAt, DateTimeOffset updatedAt)
    {
        if (startedAt is null || updatedAt < startedAt)
        {
            return GetLocalizedText(language, "Unknown");
        }

        var span = updatedAt - startedAt.Value;
        var minuteLabel = GetLocalizedText(language, "DurationMinute");
        var hourLabel = GetLocalizedText(language, "DurationHour");
        var dayLabel = GetLocalizedText(language, "DurationDay");

        if (span.TotalMinutes < 1)
        {
            return GetLocalizedText(language, "DurationLessThanMinute");
        }

        if (span.TotalHours < 1)
        {
            return $"{(int)span.TotalMinutes} {minuteLabel}";
        }

        if (span.TotalDays < 1)
        {
            return $"{(int)span.TotalHours} {hourLabel} {span.Minutes} {minuteLabel}";
        }

        return $"{span.Days} {dayLabel} {span.Hours} {hourLabel}";
    }

    private static string TrimPreview(string text, int maxLength)
    {
        var singleLine = text.Replace("\r", " ").Replace("\n", " ").Trim();

        if (singleLine.Length <= maxLength)
        {
            return singleLine;
        }

        return $"{singleLine[..maxLength].TrimEnd()}...";
    }

    private static DateTimeOffset? ParseTimestamp(string value)
    {
        return DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out var parsed)
            ? parsed
            : null;
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        return TryGetProperty(element, propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? string.Empty
            : string.Empty;
    }

    private static bool TryGetProperty(JsonElement element, string propertyName, out JsonElement property)
    {
        if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty(propertyName, out property))
        {
            return true;
        }

        property = default;
        return false;
    }

    private static string GetLocalizedText(AppLanguage language, string key)
    {
        return (language, key) switch
        {
            (_, "Unknown") => language == AppLanguage.Russian ? "\u041d\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043d\u043e" : "Unknown",
            (_, "NoPromptPreview") => language == AppLanguage.Russian
                ? "\u041f\u0440\u0435\u0434\u043f\u0440\u043e\u0441\u043c\u043e\u0442\u0440 \u0437\u0430\u043f\u0440\u043e\u0441\u0430 \u043d\u0435\u0434\u043e\u0441\u0442\u0443\u043f\u0435\u043d."
                : "No prompt preview available.",
            (_, "NoRecentMessage") => language == AppLanguage.Russian
                ? "\u041d\u0435\u0442 \u043d\u0435\u0434\u0430\u0432\u043d\u0438\u0445 \u0441\u043e\u043e\u0431\u0449\u0435\u043d\u0438\u0439."
                : "No recent message.",
            (_, "NoTranscriptFound") => language == AppLanguage.Russian
                ? "\u0422\u0440\u0430\u043d\u0441\u043a\u0440\u0438\u043f\u0442 \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f/\u0430\u0441\u0441\u0438\u0441\u0442\u0435\u043d\u0442\u0430 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d."
                : "No user/assistant transcript found.",
            (_, "LockedTitle") => language == AppLanguage.Russian
                ? "[\u0437\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u043d\u043d\u044b\u0439 \u0444\u0430\u0439\u043b \u0441\u0435\u0441\u0441\u0438\u0438]"
                : "[locked session file]",
            (_, "LockedPreview") => language == AppLanguage.Russian
                ? "\u0424\u0430\u0439\u043b \u0441\u0435\u0439\u0447\u0430\u0441 \u0438\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u0435\u0442\u0441\u044f \u0434\u0440\u0443\u0433\u0438\u043c \u043f\u0440\u043e\u0446\u0435\u0441\u0441\u043e\u043c."
                : "The file is currently used by another process.",
            (_, "LockedLastMessage") => language == AppLanguage.Russian
                ? "\u041c\u0435\u0442\u0430\u0434\u0430\u043d\u043d\u044b\u0435 \u043f\u043e\u044f\u0432\u044f\u0442\u0441\u044f, \u043a\u043e\u0433\u0434\u0430 \u0444\u0430\u0439\u043b \u0441\u0442\u0430\u043d\u0435\u0442 \u0434\u043e\u0441\u0442\u0443\u043f\u0435\u043d \u0434\u043b\u044f \u0447\u0442\u0435\u043d\u0438\u044f."
                : "Metadata will appear after the file becomes readable.",
            (_, "LockedTranscript") => language == AppLanguage.Russian
                ? "\u042d\u0442\u043e\u0442 \u0444\u0430\u0439\u043b \u0441\u0435\u0441\u0441\u0438\u0438 Codex \u0441\u0435\u0439\u0447\u0430\u0441 \u0437\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u043d \u0434\u0440\u0443\u0433\u0438\u043c \u043f\u0440\u043e\u0446\u0435\u0441\u0441\u043e\u043c."
                : "This Codex session file is locked by another process right now.",
            (_, "DurationLessThanMinute") => language == AppLanguage.Russian ? "< 1 \u043c\u0438\u043d" : "< 1 min",
            (_, "DurationMinute") => language == AppLanguage.Russian ? "\u043c\u0438\u043d" : "min",
            (_, "DurationHour") => language == AppLanguage.Russian ? "\u0447" : "h",
            (_, "DurationDay") => language == AppLanguage.Russian ? "\u0434" : "d",
            _ => key
        };
    }
}
