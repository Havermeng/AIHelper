using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class SessionService
{
    private const long LargeSessionIncrementalThresholdBytes = 25L * 1024L * 1024L;
    private const long TranscriptTailReadBytes = 4L * 1024L * 1024L;
    private const long ExternalSessionMaxParseBytes = 30L * 1024L * 1024L;
    private const int MaxTranscriptCharacters = 240_000;
    private const int MaxJsonlLineCharacters = 4 * 1024 * 1024;
    private const int SessionSnippetSchemaVersion = 2;

    private static readonly string CodexHomePath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex");

    private static readonly string SessionsRootPath = Path.Combine(CodexHomePath, "sessions");
    private static readonly string SessionIndexPath = Path.Combine(CodexHomePath, "session_index.jsonl");
    private static readonly string HistoryPath = Path.Combine(CodexHomePath, "history.jsonl");
    private static readonly string SessionSnippetsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "session-snippets.json");
    private static readonly string ArchiveRootPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "session-archive");
    private static readonly Regex SessionIdRegex = new(
        @"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})",
        RegexOptions.Compiled);
    private readonly Dictionary<SessionCacheKey, SessionCacheEntry> _sessionCache = [];
    private readonly Dictionary<string, SessionSnippetRecord> _sessionSnippets = new(StringComparer.OrdinalIgnoreCase);
    private bool _sessionSnippetsDirty;
    private bool _sessionSnippetsLoaded;
    private ThreadTitleCacheEntry? _threadTitleCache;

    public IReadOnlyList<SessionRecord> GetSessions(AppLanguage language = AppLanguage.English)
    {
        LoadSessionSnippets();
        var threadTitles = LoadThreadTitlesSnapshot();
        var files = DiscoverSessionFiles()
            .OrderByDescending(source => source.File.LastWriteTimeUtc)
            .ToList();

        var sessions = new List<SessionRecord>(files.Count);
        var activeKeys = new HashSet<SessionCacheKey>();

        foreach (var sourceFile in files)
        {
            var file = sourceFile.File;
            var cacheKey = new SessionCacheKey(file.FullName, language);
            activeKeys.Add(cacheKey);

            try
            {
                if (TryGetCachedSession(cacheKey, file, threadTitles, out var cachedSession))
                {
                    sessions.Add(cachedSession);
                    continue;
                }

                if (TryRestorePersistedSession(sourceFile, threadTitles.Titles, language, out var persistedSession))
                {
                    _sessionCache[cacheKey] = new SessionCacheEntry(
                        file.LastWriteTimeUtc.Ticks,
                        file.Length,
                        threadTitles.VersionTicks,
                        threadTitles.VersionLength,
                        persistedSession);
                    sessions.Add(persistedSession);
                    continue;
                }

                var session = ParseDiscoveredSessionFile(sourceFile, threadTitles.Titles, language);

                if (session is not null)
                {
                    UpsertSessionSnippet(session, language);
                    _sessionCache[cacheKey] = new SessionCacheEntry(
                        file.LastWriteTimeUtc.Ticks,
                        file.Length,
                        threadTitles.VersionTicks,
                        threadTitles.VersionLength,
                        session);
                    sessions.Add(session);
                }
            }
            catch (IOException)
            {
                sessions.Add(CreateLockedSessionRecord(sourceFile, language, threadTitles.Titles, TryGetAnyCachedSession(cacheKey)));
            }
            catch (UnauthorizedAccessException)
            {
                sessions.Add(CreateLockedSessionRecord(sourceFile, language, threadTitles.Titles, TryGetAnyCachedSession(cacheKey)));
            }
        }

        CleanupSessionCache(activeKeys);
        SaveSessionSnippetsIfDirty();

        return sessions
            .OrderByDescending(session => session.UpdatedAtUtc)
            .ThenBy(session => session.Title, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public void DeleteSession(SessionRecord session)
    {
        ThrowIfReparsePoint(session.FilePath);
        if (!File.Exists(session.FilePath))
        {
            throw new FileNotFoundException("Session file not found.", session.FilePath);
        }

        File.Delete(session.FilePath);

        if (session.IsCodexSession)
        {
            RemoveJsonlEntries(SessionIndexPath, root => GetString(root, "id") != session.SessionId);
            RemoveJsonlEntries(HistoryPath, root => GetString(root, "session_id") != session.SessionId);
            DeleteEmptyParentDirectories(Path.GetDirectoryName(session.FilePath));
        }
    }

    public string ArchiveSession(SessionRecord session)
    {
        ThrowIfReparsePoint(session.FilePath);
        if (!File.Exists(session.FilePath))
        {
            throw new FileNotFoundException("Session file not found.", session.FilePath);
        }

        var destinationPath = CreateArchiveDestinationPath(session);
        var destinationDirectory = Path.GetDirectoryName(destinationPath);

        if (!string.IsNullOrWhiteSpace(destinationDirectory))
        {
            Directory.CreateDirectory(destinationDirectory);
        }

        File.Move(session.FilePath, destinationPath);

        if (session.IsCodexSession)
        {
            RemoveJsonlEntries(SessionIndexPath, root => GetString(root, "id") != session.SessionId);
            RemoveJsonlEntries(HistoryPath, root => GetString(root, "session_id") != session.SessionId);
            DeleteEmptyParentDirectories(Path.GetDirectoryName(session.FilePath));
        }

        return destinationPath;
    }

    public string GetSessionArchiveRootPath() => ArchiveRootPath;

    public CodexSessionConversation GetConversation(SessionRecord session)
    {
        ThrowIfReparsePoint(session.FilePath);
        if (!File.Exists(session.FilePath))
        {
            throw new FileNotFoundException("Session file not found.", session.FilePath);
        }

        if (!session.IsCodexSession)
        {
            return ParseExternalConversationFile(new FileInfo(session.FilePath), session);
        }

        var threadTitles = LoadThreadTitlesSnapshot();
        var file = new FileInfo(session.FilePath);
        return ParseConversationFile(file, threadTitles.Titles, session);
    }

    public string LoadTranscriptText(SessionRecord session, AppLanguage language = AppLanguage.English)
    {
        if (IsReparsePoint(session.FilePath))
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }

        if (!File.Exists(session.FilePath))
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }

        if (!string.IsNullOrWhiteSpace(session.TranscriptText))
        {
            return session.TranscriptText;
        }

        try
        {
            var transcript = session.IsCodexSession
                ? BuildTranscriptText(session.FilePath, language)
                : BuildExternalTranscriptText(session.FilePath, language);
            session.TranscriptText = transcript;
            UpsertSessionSnippet(session);
            SaveSessionSnippetsIfDirty();
            return transcript;
        }
        catch (IOException)
        {
            LoadSessionSnippets();
            var lockedText = TryGetSessionSnippet(session, out var snippet)
                ? BuildLockedSnippetTranscript(snippet, language)
                : GetLocalizedText(language, "LockedTranscript");
            session.TranscriptText = lockedText;
            return lockedText;
        }
        catch (UnauthorizedAccessException)
        {
            LoadSessionSnippets();
            var lockedText = TryGetSessionSnippet(session, out var snippet)
                ? BuildLockedSnippetTranscript(snippet, language)
                : GetLocalizedText(language, "LockedTranscript");
            session.TranscriptText = lockedText;
            return lockedText;
        }
    }

    private static IReadOnlyList<DiscoveredSessionFile> DiscoverSessionFiles()
    {
        var discovered = new List<DiscoveredSessionFile>();
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        AddJsonlFiles(
            discovered,
            Path.Combine(profile, ".codex", "sessions"),
            "Codex",
            SessionFileKind.Codex);
        AddJsonlFiles(
            discovered,
            Path.Combine(profile, ".qwen", "projects"),
            "Qwen",
            SessionFileKind.Qwen,
            path => path.Contains($"{Path.DirectorySeparatorChar}chats{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase));
        AddJsonlFiles(
            discovered,
            Path.Combine(profile, ".cursor", "projects"),
            "Cursor",
            SessionFileKind.Cursor,
            path => path.Contains("agent-transcripts", StringComparison.OrdinalIgnoreCase));
        AddJsonFiles(
            discovered,
            Path.Combine(profile, ".continue", "sessions"),
            "Continue",
            SessionFileKind.Continue,
            path => !string.Equals(Path.GetFileName(path), "sessions.json", StringComparison.OrdinalIgnoreCase));
        AddJsonlFiles(
            discovered,
            Path.Combine(profile, ".claude", "projects"),
            "Claude",
            SessionFileKind.ClaudeCode);
        AddJsonlFiles(
            discovered,
            Path.Combine(profile, ".gemini", "sessions"),
            "Gemini",
            SessionFileKind.GenericJsonl);

        foreach (var root in new[]
                 {
                     Path.Combine(appData, "OpenCode", "sessions"),
                     Path.Combine(appData, "ai.opencode.desktop", "sessions"),
                     Path.Combine(localAppData, "OpenCode", "sessions"),
                     Path.Combine(localAppData, "ai.opencode.desktop", "sessions")
                 })
        {
            AddJsonFiles(discovered, root, "OpenCode", SessionFileKind.GenericJson);
            AddJsonlFiles(discovered, root, "OpenCode", SessionFileKind.GenericJsonl);
        }

        return discovered
            .GroupBy(item => item.File.FullName, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private static void AddJsonlFiles(
        List<DiscoveredSessionFile> discovered,
        string rootPath,
        string applicationName,
        SessionFileKind kind,
        Func<string, bool>? predicate = null)
    {
        AddFiles(discovered, rootPath, "*.jsonl", applicationName, kind, predicate);
    }

    private static void AddJsonFiles(
        List<DiscoveredSessionFile> discovered,
        string rootPath,
        string applicationName,
        SessionFileKind kind,
        Func<string, bool>? predicate = null)
    {
        AddFiles(discovered, rootPath, "*.json", applicationName, kind, predicate);
    }

    private static void AddFiles(
        List<DiscoveredSessionFile> discovered,
        string rootPath,
        string pattern,
        string applicationName,
        SessionFileKind kind,
        Func<string, bool>? predicate)
    {
        if (!Directory.Exists(rootPath))
        {
            return;
        }

        try
        {
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible = true,
                AttributesToSkip = FileAttributes.ReparsePoint
            };

            foreach (var file in new DirectoryInfo(rootPath).EnumerateFiles(pattern, options))
            {
                if (predicate is not null && !predicate(file.FullName))
                {
                    continue;
                }

                discovered.Add(new DiscoveredSessionFile(file, rootPath, applicationName, kind));
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
        }
    }

    private static SessionRecord? ParseDiscoveredSessionFile(
        DiscoveredSessionFile sourceFile,
        IReadOnlyDictionary<string, string> titleLookup,
        AppLanguage language)
    {
        return sourceFile.Kind switch
        {
            SessionFileKind.Codex => ParseSessionFile(sourceFile.File, titleLookup, language),
            SessionFileKind.Qwen => ParseQwenSessionFile(sourceFile, language),
            SessionFileKind.Continue => ParseContinueSessionFile(sourceFile, language),
            SessionFileKind.ClaudeCode => ParseClaudeCodeSessionFile(sourceFile, language),
            SessionFileKind.Cursor => ParseGenericJsonlSessionFile(sourceFile, language),
            SessionFileKind.GenericJson => ParseGenericJsonSessionFile(sourceFile, language),
            _ => ParseGenericJsonlSessionFile(sourceFile, language)
        };
    }

    private ThreadTitleCacheEntry LoadThreadTitlesSnapshot()
    {
        if (!File.Exists(SessionIndexPath))
        {
            _threadTitleCache = new ThreadTitleCacheEntry(0, 0, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
            return _threadTitleCache;
        }

        var fileInfo = new FileInfo(SessionIndexPath);

        if (_threadTitleCache is not null &&
            _threadTitleCache.VersionTicks == fileInfo.LastWriteTimeUtc.Ticks &&
            _threadTitleCache.VersionLength == fileInfo.Length)
        {
            return _threadTitleCache;
        }

        var titles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in ReadJsonlLinesSafely(SessionIndexPath))
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

        _threadTitleCache = new ThreadTitleCacheEntry(fileInfo.LastWriteTimeUtc.Ticks, fileInfo.Length, titles);
        return _threadTitleCache;
    }

    private bool TryGetCachedSession(
        SessionCacheKey cacheKey,
        FileInfo file,
        ThreadTitleCacheEntry threadTitles,
        out SessionRecord session)
    {
        if (!_sessionCache.TryGetValue(cacheKey, out var cacheEntry) ||
            cacheEntry.ThreadTitleVersionTicks != threadTitles.VersionTicks ||
            cacheEntry.ThreadTitleVersionLength != threadTitles.VersionLength)
        {
            session = null!;
            return false;
        }

        if (cacheEntry.FileLastWriteTicks == file.LastWriteTimeUtc.Ticks &&
            cacheEntry.FileLength == file.Length)
        {
            session = cacheEntry.Session;
            return true;
        }

        if (file.Length >= LargeSessionIncrementalThresholdBytes &&
            file.Length >= cacheEntry.FileLength)
        {
            session = CreateLargeFileCachedSession(cacheEntry.Session, file, cacheKey.Language);
            _sessionCache[cacheKey] = cacheEntry with
            {
                FileLastWriteTicks = file.LastWriteTimeUtc.Ticks,
                FileLength = file.Length,
                Session = session
            };
            return true;
        }

        session = null!;
        return false;
    }

    private SessionRecord? TryGetAnyCachedSession(SessionCacheKey cacheKey)
    {
        return _sessionCache.TryGetValue(cacheKey, out var cacheEntry)
            ? cacheEntry.Session
            : null;
    }

    private bool TryRestorePersistedSession(
        DiscoveredSessionFile sourceFile,
        IReadOnlyDictionary<string, string> titleLookup,
        AppLanguage language,
        out SessionRecord session)
    {
        var file = sourceFile.File;
        var snippetKey = NormalizeSnippetKey(file.FullName);
        if (!_sessionSnippets.TryGetValue(snippetKey, out var snippet) ||
            !IsUsefulSnippet(snippet) ||
            snippet.FileLastWriteTicks != file.LastWriteTimeUtc.Ticks ||
            snippet.FileLength != file.Length ||
            !string.Equals(snippet.Language, language.ToString(), StringComparison.Ordinal))
        {
            session = null!;
            return false;
        }

        if (sourceFile.Kind == SessionFileKind.ClaudeCode &&
            snippet.SchemaVersion < SessionSnippetSchemaVersion)
        {
            session = null!;
            return false;
        }

        var sessionId = !string.IsNullOrWhiteSpace(snippet.SessionId)
            ? snippet.SessionId
            : sourceFile.Kind == SessionFileKind.Codex
                ? DeriveSessionId(file.Name)
                : $"{sourceFile.ApplicationName.ToLowerInvariant()}:{Path.GetFileNameWithoutExtension(file.Name)}";
        var storedTitle = !string.IsNullOrWhiteSpace(snippet.OriginalTitle)
            ? snippet.OriginalTitle
            : snippet.Title;
        var title = sourceFile.Kind == SessionFileKind.Codex &&
                    titleLookup.TryGetValue(sessionId, out var indexedTitle) &&
                    !string.IsNullOrWhiteSpace(indexedTitle)
            ? indexedTitle
            : storedTitle;
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);
        var preview = string.IsNullOrWhiteSpace(snippet.Preview)
            ? GetLocalizedText(language, "NoPromptPreview")
            : snippet.Preview;
        var lastMessage = string.IsNullOrWhiteSpace(snippet.LastMessagePreview)
            ? GetLocalizedText(language, "NoRecentMessage")
            : snippet.LastMessagePreview;
        var workingDirectory = string.IsNullOrWhiteSpace(snippet.WorkingDirectory) ? "-" : snippet.WorkingDirectory;
        var source = string.IsNullOrWhiteSpace(snippet.Source) ? sourceFile.ApplicationName : snippet.Source;
        var modelProvider = string.IsNullOrWhiteSpace(snippet.ModelProvider) ? "-" : snippet.ModelProvider;
        var baseSearchBlob = BuildSearchBlob(
            title,
            sessionId,
            workingDirectory,
            preview,
            lastMessage,
            modelProvider,
            source);

        session = new SessionRecord
        {
            SessionId = sessionId,
            Title = string.IsNullOrWhiteSpace(title) ? file.Name : title,
            Preview = preview,
            LastMessagePreview = lastMessage,
            StartedAtText = string.IsNullOrWhiteSpace(snippet.StartedAtText)
                ? GetLocalizedText(language, "Unknown")
                : snippet.StartedAtText,
            UpdatedAtText = updatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = string.IsNullOrWhiteSpace(snippet.DurationText)
                ? GetLocalizedText(language, "Unknown")
                : snippet.DurationText,
            WorkingDirectory = workingDirectory,
            Source = source,
            ModelProvider = modelProvider,
            CliVersion = string.IsNullOrWhiteSpace(snippet.CliVersion) ? "-" : snippet.CliVersion,
            FilePath = file.FullName,
            RelativePath = string.IsNullOrWhiteSpace(snippet.RelativePath)
                ? Path.GetRelativePath(sourceFile.RootPath, file.FullName)
                : snippet.RelativePath,
            TranscriptText = snippet.TotalMessageCount == 0
                ? GetLocalizedText(language, "NoTranscriptFound")
                : string.Empty,
            UserMessageCount = snippet.UserMessageCount,
            AssistantMessageCount = snippet.AssistantMessageCount,
            ToolCallCount = snippet.ToolCallCount,
            TotalMessageCount = snippet.TotalMessageCount,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
        return true;
    }

    private static SessionRecord CreateLargeFileCachedSession(
        SessionRecord cachedSession,
        FileInfo file,
        AppLanguage language)
    {
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);
        var baseSearchBlob = cachedSession.BaseSearchBlob;

        return new SessionRecord
        {
            SessionId = cachedSession.SessionId,
            Title = cachedSession.Title,
            Preview = cachedSession.Preview,
            LastMessagePreview = cachedSession.LastMessagePreview,
            StartedAtText = cachedSession.StartedAtText,
            UpdatedAtText = updatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = cachedSession.DurationText,
            WorkingDirectory = cachedSession.WorkingDirectory,
            Source = cachedSession.Source,
            ModelProvider = cachedSession.ModelProvider,
            CliVersion = cachedSession.CliVersion,
            FilePath = cachedSession.FilePath,
            RelativePath = cachedSession.RelativePath,
            TranscriptText = string.Empty,
            UserMessageCount = cachedSession.UserMessageCount,
            AssistantMessageCount = cachedSession.AssistantMessageCount,
            ToolCallCount = cachedSession.ToolCallCount,
            TotalMessageCount = cachedSession.TotalMessageCount,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
    }

    private void CleanupSessionCache(HashSet<SessionCacheKey> activeKeys)
    {
        if (_sessionCache.Count == 0)
        {
            return;
        }

        var staleKeys = _sessionCache.Keys
            .Where(key => !activeKeys.Contains(key))
            .ToList();

        foreach (var staleKey in staleKeys)
        {
            _sessionCache.Remove(staleKey);
        }
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
        foreach (var line in ReadJsonlLinesSafely(file.FullName))
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
        var totalMessageCount = userMessageCount + assistantMessageCount;
        var unknownText = GetLocalizedText(language, "Unknown");
        var transcriptText = totalMessageCount == 0
            ? GetLocalizedText(language, "NoTranscriptFound")
            : string.Empty;

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
            Source = "Codex",
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

    private static SessionRecord? ParseQwenSessionFile(DiscoveredSessionFile sourceFile, AppLanguage language)
    {
        var file = sourceFile.File;
        if (file.Length > ExternalSessionMaxParseBytes)
        {
            return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "LargeExternalSession"));
        }

        string sessionId = Path.GetFileNameWithoutExtension(file.Name);
        string firstPrompt = string.Empty;
        string preview = string.Empty;
        string lastMessage = string.Empty;
        string workingDirectory = string.Empty;
        string modelProvider = string.Empty;
        string cliVersion = string.Empty;
        DateTimeOffset? startedAt = null;
        var userMessageCount = 0;
        var assistantMessageCount = 0;
        var toolCallCount = 0;

        foreach (var line in ReadJsonlLinesSafely(file.FullName))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                sessionId = GetString(root, "sessionId") is { Length: > 0 } parsedSessionId
                    ? parsedSessionId
                    : sessionId;
                startedAt ??= ParseTimestamp(GetString(root, "timestamp"));

                var parsedWorkingDirectory = GetString(root, "cwd");
                if (!string.IsNullOrWhiteSpace(parsedWorkingDirectory))
                {
                    workingDirectory = parsedWorkingDirectory;
                }

                var parsedModel = GetString(root, "model");
                if (!string.IsNullOrWhiteSpace(parsedModel))
                {
                    modelProvider = parsedModel;
                }

                var parsedVersion = GetString(root, "version");
                if (!string.IsNullOrWhiteSpace(parsedVersion))
                {
                    cliVersion = parsedVersion;
                }

                var type = GetString(root, "type");
                if (type == "tool_result")
                {
                    continue;
                }

                if (type == "system")
                {
                    if (TryGetProperty(root, "systemPayload", out var systemPayload) &&
                        TryGetProperty(systemPayload, "uiEvent", out var uiEvent) &&
                        !string.IsNullOrWhiteSpace(GetString(uiEvent, "function_name")))
                    {
                        toolCallCount++;
                    }

                    continue;
                }

                if (!TryGetProperty(root, "message", out var message))
                {
                    continue;
                }

                var role = NormalizeRole(GetString(message, "role"));
                if (string.IsNullOrWhiteSpace(role))
                {
                    role = NormalizeRole(type);
                }

                var text = ExtractPartsText(message);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (role == "user")
                {
                    userMessageCount++;
                    if (string.IsNullOrWhiteSpace(firstPrompt))
                    {
                        firstPrompt = text;
                        preview = TrimPreview(text, 180);
                    }
                }
                else if (role == "assistant")
                {
                    assistantMessageCount++;
                }
                else
                {
                    continue;
                }

                lastMessage = text;
            }
            catch (JsonException)
            {
            }
        }

        return CreateExternalSessionRecord(
            sourceFile,
            language,
            $"qwen:{sessionId}",
            ChooseTitle(null, firstPrompt, file.Name),
            preview,
            lastMessage,
            startedAt,
            workingDirectory,
            modelProvider,
            cliVersion,
            userMessageCount,
            assistantMessageCount,
            toolCallCount);
    }

    private static SessionRecord? ParseContinueSessionFile(DiscoveredSessionFile sourceFile, AppLanguage language)
    {
        var file = sourceFile.File;
        if (file.Length > ExternalSessionMaxParseBytes)
        {
            return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "LargeExternalSession"));
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(file.FullName, Encoding.UTF8));
            var root = document.RootElement;
            var sessionId = GetString(root, "sessionId");
            var title = GetString(root, "title");
            var workingDirectory = NormalizeFileUriPath(GetString(root, "workspaceDirectory"));
            var modelProvider = GetString(root, "chatModelTitle");
            var userMessageCount = 0;
            var assistantMessageCount = 0;
            var firstPrompt = string.Empty;
            var preview = string.Empty;
            var lastMessage = string.Empty;

            if (TryGetProperty(root, "history", out var history) && history.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in history.EnumerateArray())
                {
                    if (!TryGetProperty(item, "message", out var message))
                    {
                        continue;
                    }

                    var role = NormalizeRole(GetString(message, "role"));
                    var text = ExtractFlexibleContentText(message);
                    if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (role == "user")
                    {
                        userMessageCount++;
                        if (string.IsNullOrWhiteSpace(firstPrompt))
                        {
                            firstPrompt = text;
                            preview = TrimPreview(text, 180);
                        }
                    }
                    else if (role == "assistant")
                    {
                        assistantMessageCount++;
                    }

                    lastMessage = text;
                }
            }

            if (string.IsNullOrWhiteSpace(sessionId))
            {
                sessionId = Path.GetFileNameWithoutExtension(file.Name);
            }

            return CreateExternalSessionRecord(
                sourceFile,
                language,
                $"continue:{sessionId}",
                ChooseTitle(title, firstPrompt, file.Name),
                preview,
                lastMessage,
                startedAt: null,
                workingDirectory,
                modelProvider,
                cliVersion: "-",
                userMessageCount,
                assistantMessageCount,
                toolCallCount: 0);
        }
        catch (JsonException)
        {
            return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "NoTranscriptFound"));
        }
    }

    private static SessionRecord? ParseGenericJsonSessionFile(DiscoveredSessionFile sourceFile, AppLanguage language)
    {
        return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "ExternalSessionFile"));
    }

    private static SessionRecord? ParseClaudeCodeSessionFile(DiscoveredSessionFile sourceFile, AppLanguage language)
    {
        var file = sourceFile.File;
        if (file.Length > ExternalSessionMaxParseBytes)
        {
            return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "LargeExternalSession"));
        }

        var sessionId = Path.GetFileNameWithoutExtension(file.Name);
        string customTitle = string.Empty;
        string aiTitle = string.Empty;
        string summaryTitle = string.Empty;
        string firstPrompt = string.Empty;
        string preview = string.Empty;
        string lastMessage = string.Empty;
        string workingDirectory = InferProjectPathFromSourceFile(sourceFile);
        string modelProvider = string.Empty;
        string cliVersion = string.Empty;
        DateTimeOffset? startedAt = null;
        var userMessageCount = 0;
        var assistantMessageCount = 0;
        var toolCallCount = 0;
        var sawMainlineEntries = false;

        foreach (var line in ReadJsonlLinesSafely(file.FullName))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;

                if (root.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var entryType = GetString(root, "type");

                if (string.Equals(entryType, "summary", StringComparison.OrdinalIgnoreCase))
                {
                    var summary = GetString(root, "summary");
                    if (!string.IsNullOrWhiteSpace(summary))
                    {
                        summaryTitle = summary;
                    }

                    continue;
                }

                if (string.Equals(entryType, "custom-title", StringComparison.OrdinalIgnoreCase))
                {
                    var parsedCustomTitle = GetString(root, "customTitle");
                    if (!string.IsNullOrWhiteSpace(parsedCustomTitle))
                    {
                        customTitle = parsedCustomTitle;
                    }

                    continue;
                }

                if (string.Equals(entryType, "ai-title", StringComparison.OrdinalIgnoreCase))
                {
                    var parsedAiTitle = GetString(root, "aiTitle");
                    if (!string.IsNullOrWhiteSpace(parsedAiTitle))
                    {
                        aiTitle = parsedAiTitle;
                    }

                    continue;
                }

                var isUserEntry = string.Equals(entryType, "user", StringComparison.OrdinalIgnoreCase);
                var isAssistantEntry = string.Equals(entryType, "assistant", StringComparison.OrdinalIgnoreCase);

                if (!isUserEntry && !isAssistantEntry)
                {
                    continue;
                }

                if (TryGetProperty(root, "isSidechain", out var sidechain) &&
                    sidechain.ValueKind == JsonValueKind.True)
                {
                    continue;
                }

                sawMainlineEntries = true;
                startedAt ??= ParseTimestamp(GetString(root, "timestamp"));

                var parsedSessionId = GetString(root, "sessionId");
                if (!string.IsNullOrWhiteSpace(parsedSessionId))
                {
                    sessionId = parsedSessionId;
                }

                var parsedWorkingDirectory = GetString(root, "cwd");
                if (!string.IsNullOrWhiteSpace(parsedWorkingDirectory))
                {
                    workingDirectory = parsedWorkingDirectory;
                }

                var parsedVersion = GetString(root, "version");
                if (!string.IsNullOrWhiteSpace(parsedVersion))
                {
                    cliVersion = parsedVersion;
                }

                if (!TryGetProperty(root, "message", out var message))
                {
                    continue;
                }

                var text = ExtractClaudeMessageText(message, ref toolCallCount);

                if (isAssistantEntry)
                {
                    var parsedModel = GetString(message, "model");
                    if (!string.IsNullOrWhiteSpace(parsedModel))
                    {
                        modelProvider = parsedModel;
                    }

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        assistantMessageCount++;
                        lastMessage = text;
                    }

                    continue;
                }

                if (TryGetProperty(root, "isMeta", out var isMeta) &&
                    isMeta.ValueKind == JsonValueKind.True)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(text) || IsClaudeServiceUserText(text))
                {
                    continue;
                }

                userMessageCount++;
                if (string.IsNullOrWhiteSpace(firstPrompt))
                {
                    firstPrompt = text;
                    preview = TrimPreview(text, 180);
                }

                lastMessage = text;
            }
            catch (JsonException)
            {
            }
        }

        if (!sawMainlineEntries)
        {
            return null;
        }

        var bestTitle = new[] { customTitle, aiTitle, summaryTitle }
            .FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate));

        return CreateExternalSessionRecord(
            sourceFile,
            language,
            sessionId,
            ChooseTitle(bestTitle, firstPrompt, file.Name),
            preview,
            lastMessage,
            startedAt,
            workingDirectory,
            modelProvider,
            cliVersion,
            userMessageCount,
            assistantMessageCount,
            toolCallCount);
    }

    private static string ExtractClaudeMessageText(JsonElement message, ref int toolCallCount)
    {
        if (!TryGetProperty(message, "content", out var content))
        {
            return string.Empty;
        }

        if (content.ValueKind == JsonValueKind.String)
        {
            return content.GetString() ?? string.Empty;
        }

        if (content.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();

        foreach (var block in content.EnumerateArray())
        {
            if (block.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var blockType = GetString(block, "type");

            if (string.Equals(blockType, "tool_use", StringComparison.OrdinalIgnoreCase))
            {
                toolCallCount++;
                continue;
            }

            if (!string.Equals(blockType, "text", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var text = GetString(block, "text");

            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (builder.Length > 0)
            {
                builder.AppendLine();
            }

            builder.Append(text.Trim());
        }

        return builder.ToString();
    }

    private static bool IsClaudeServiceUserText(string text)
    {
        var trimmed = text.TrimStart();

        return trimmed.StartsWith("<command-", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("<local-command", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("<system-reminder", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("<task-notification", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("Caveat: The messages below were generated", StringComparison.Ordinal);
    }

    private static SessionRecord? ParseGenericJsonlSessionFile(DiscoveredSessionFile sourceFile, AppLanguage language)
    {
        var file = sourceFile.File;
        if (file.Length > ExternalSessionMaxParseBytes)
        {
            return CreateExternalFileRecord(sourceFile, language, GetLocalizedText(language, "LargeExternalSession"));
        }

        var sessionId = Path.GetFileNameWithoutExtension(file.Name);
        string firstPrompt = string.Empty;
        string preview = string.Empty;
        string lastMessage = string.Empty;
        string workingDirectory = InferProjectPathFromSourceFile(sourceFile);
        string modelProvider = string.Empty;
        string cliVersion = string.Empty;
        DateTimeOffset? startedAt = null;
        var userMessageCount = 0;
        var assistantMessageCount = 0;
        var toolCallCount = 0;

        foreach (var line in ReadJsonlLinesSafely(file.FullName))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                startedAt ??= ParseTimestamp(GetString(root, "timestamp"));

                var parsedSessionId = GetString(root, "sessionId");
                if (!string.IsNullOrWhiteSpace(parsedSessionId))
                {
                    sessionId = parsedSessionId;
                }

                var parsedWorkingDirectory = GetString(root, "cwd");
                if (!string.IsNullOrWhiteSpace(parsedWorkingDirectory))
                {
                    workingDirectory = parsedWorkingDirectory;
                }

                var parsedModel = GetString(root, "model");
                if (!string.IsNullOrWhiteSpace(parsedModel))
                {
                    modelProvider = parsedModel;
                }

                var parsedVersion = GetString(root, "version");
                if (!string.IsNullOrWhiteSpace(parsedVersion))
                {
                    cliVersion = parsedVersion;
                }

                var role = NormalizeRole(GetString(root, "role"));
                var text = ExtractFlexibleContentText(root);
                if (string.IsNullOrWhiteSpace(text) && TryGetProperty(root, "message", out var message))
                {
                    role = string.IsNullOrWhiteSpace(role)
                        ? NormalizeRole(GetString(message, "role"))
                        : role;
                    text = ExtractFlexibleContentText(message);
                }

                if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(text))
                {
                    if (ContainsToolCall(root))
                    {
                        toolCallCount++;
                    }

                    continue;
                }

                if (role == "user")
                {
                    userMessageCount++;
                    if (string.IsNullOrWhiteSpace(firstPrompt))
                    {
                        firstPrompt = text;
                        preview = TrimPreview(text, 180);
                    }
                }
                else if (role == "assistant")
                {
                    assistantMessageCount++;
                }
                else
                {
                    continue;
                }

                lastMessage = text;
            }
            catch (JsonException)
            {
            }
        }

        return CreateExternalSessionRecord(
            sourceFile,
            language,
            $"{sourceFile.ApplicationName.ToLowerInvariant()}:{sessionId}",
            ChooseTitle(null, firstPrompt, file.Name),
            preview,
            lastMessage,
            startedAt,
            workingDirectory,
            modelProvider,
            cliVersion,
            userMessageCount,
            assistantMessageCount,
            toolCallCount);
    }

    private static SessionRecord CreateExternalSessionRecord(
        DiscoveredSessionFile sourceFile,
        AppLanguage language,
        string sessionId,
        string title,
        string preview,
        string lastMessage,
        DateTimeOffset? startedAt,
        string workingDirectory,
        string modelProvider,
        string cliVersion,
        int userMessageCount,
        int assistantMessageCount,
        int toolCallCount)
    {
        var file = sourceFile.File;
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);
        var updatedLocal = updatedAt.ToLocalTime();
        var startedLocal = startedAt?.ToLocalTime();
        var unknownText = GetLocalizedText(language, "Unknown");
        var totalMessageCount = userMessageCount + assistantMessageCount;
        var normalizedPreview = string.IsNullOrWhiteSpace(preview)
            ? GetLocalizedText(language, "NoPromptPreview")
            : preview;
        var normalizedLastMessage = string.IsNullOrWhiteSpace(lastMessage)
            ? GetLocalizedText(language, "NoRecentMessage")
            : TrimPreview(lastMessage, 220);
        var normalizedWorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory)
            ? "-"
            : workingDirectory;
        var normalizedModelProvider = string.IsNullOrWhiteSpace(modelProvider)
            ? "-"
            : modelProvider;
        var normalizedCliVersion = string.IsNullOrWhiteSpace(cliVersion)
            ? "-"
            : cliVersion;
        var baseSearchBlob = BuildSearchBlob(
            title,
            sessionId,
            normalizedWorkingDirectory,
            normalizedPreview,
            normalizedLastMessage,
            normalizedModelProvider,
            sourceFile.ApplicationName,
            file.FullName);

        return new SessionRecord
        {
            SessionId = sessionId,
            Title = title,
            Preview = normalizedPreview,
            LastMessagePreview = normalizedLastMessage,
            StartedAtText = startedLocal?.ToString("dd.MM.yyyy HH:mm:ss") ?? unknownText,
            UpdatedAtText = updatedLocal.ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = FormatDuration(language, startedAt, updatedAt),
            WorkingDirectory = normalizedWorkingDirectory,
            Source = sourceFile.ApplicationName,
            ModelProvider = normalizedModelProvider,
            CliVersion = normalizedCliVersion,
            FilePath = file.FullName,
            RelativePath = Path.GetRelativePath(sourceFile.RootPath, file.FullName),
            TranscriptText = totalMessageCount == 0 ? GetLocalizedText(language, "NoTranscriptFound") : string.Empty,
            UserMessageCount = userMessageCount,
            AssistantMessageCount = assistantMessageCount,
            ToolCallCount = toolCallCount,
            TotalMessageCount = totalMessageCount,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
    }

    private static SessionRecord CreateExternalFileRecord(
        DiscoveredSessionFile sourceFile,
        AppLanguage language,
        string preview)
    {
        return CreateExternalSessionRecord(
            sourceFile,
            language,
            $"{sourceFile.ApplicationName.ToLowerInvariant()}:{Path.GetFileNameWithoutExtension(sourceFile.File.Name)}",
            Path.GetFileNameWithoutExtension(sourceFile.File.Name),
            preview,
            preview,
            startedAt: null,
            workingDirectory: InferProjectPathFromSourceFile(sourceFile),
            modelProvider: "-",
            cliVersion: "-",
            userMessageCount: 0,
            assistantMessageCount: 0,
            toolCallCount: 0);
    }

    private static string BuildExternalTranscriptText(string filePath, AppLanguage language)
    {
        var file = new FileInfo(filePath);
        if (!file.Exists)
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }

        if (file.Length > ExternalSessionMaxParseBytes)
        {
            return GetLocalizedText(language, "LargeExternalSession");
        }

        if (string.Equals(file.Extension, ".json", StringComparison.OrdinalIgnoreCase))
        {
            return BuildExternalJsonTranscriptText(file, language);
        }

        var transcript = new StringBuilder();
        var wasTrimmed = false;

        foreach (var line in ReadJsonlLinesSafely(file.FullName))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                var role = NormalizeRole(GetString(root, "role"));
                var text = ExtractFlexibleContentText(root);

                if (string.IsNullOrWhiteSpace(text) && TryGetProperty(root, "message", out var message))
                {
                    role = string.IsNullOrWhiteSpace(role)
                        ? NormalizeRole(GetString(message, "role"))
                        : role;
                    text = ExtractFlexibleContentText(message);
                }

                if (string.IsNullOrWhiteSpace(role))
                {
                    role = NormalizeRole(GetString(root, "type"));
                }

                if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                AppendTranscriptLine(transcript, ParseTimestamp(GetString(root, "timestamp")), role, text);
                wasTrimmed |= TrimTranscriptBuilder(transcript);
            }
            catch (JsonException)
            {
            }
        }

        if (transcript.Length == 0)
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }

        var result = transcript.ToString().Trim();
        return wasTrimmed
            ? $"{GetLocalizedText(language, "TranscriptTrimmedNotice")}{Environment.NewLine}{Environment.NewLine}{result}"
            : result;
    }

    private static string BuildExternalJsonTranscriptText(FileInfo file, AppLanguage language)
    {
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(file.FullName, Encoding.UTF8));
            var root = document.RootElement;
            var transcript = new StringBuilder();
            var wasTrimmed = false;

            if (TryGetProperty(root, "history", out var history) && history.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in history.EnumerateArray())
                {
                    if (!TryGetProperty(item, "message", out var message))
                    {
                        continue;
                    }

                    var role = NormalizeRole(GetString(message, "role"));
                    var text = ExtractFlexibleContentText(message);
                    if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    AppendTranscriptLine(transcript, null, role, text);
                    wasTrimmed |= TrimTranscriptBuilder(transcript);
                }
            }

            if (transcript.Length == 0)
            {
                return GetLocalizedText(language, "NoTranscriptFound");
            }

            var result = transcript.ToString().Trim();
            return wasTrimmed
                ? $"{GetLocalizedText(language, "TranscriptTrimmedNotice")}{Environment.NewLine}{Environment.NewLine}{result}"
                : result;
        }
        catch (JsonException)
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }
    }

    private static CodexSessionConversation ParseExternalConversationFile(FileInfo file, SessionRecord sourceSession)
    {
        var messages = new List<CodexSessionMessage>();

        if (file.Exists && file.Length <= ExternalSessionMaxParseBytes)
        {
            if (string.Equals(file.Extension, ".json", StringComparison.OrdinalIgnoreCase))
            {
                AppendExternalJsonMessages(file, messages);
            }
            else
            {
                AppendExternalJsonlMessages(file, messages);
            }
        }

        return new CodexSessionConversation
        {
            SessionId = sourceSession.SessionId,
            Title = sourceSession.DisplayTitle,
            WorkingDirectory = string.IsNullOrWhiteSpace(sourceSession.WorkingDirectory) || sourceSession.WorkingDirectory == "-"
                ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                : sourceSession.WorkingDirectory,
            ModelProvider = sourceSession.ModelProvider == "-" ? string.Empty : sourceSession.ModelProvider,
            StartedAtUtc = null,
            UpdatedAtUtc = new DateTimeOffset(file.LastWriteTimeUtc),
            Messages = messages
        };
    }

    private static void AppendExternalJsonMessages(FileInfo file, List<CodexSessionMessage> messages)
    {
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(file.FullName, Encoding.UTF8));
            if (!TryGetProperty(document.RootElement, "history", out var history) ||
                history.ValueKind != JsonValueKind.Array)
            {
                return;
            }

            foreach (var item in history.EnumerateArray())
            {
                if (!TryGetProperty(item, "message", out var message))
                {
                    continue;
                }

                AddExternalMessage(messages, NormalizeRole(GetString(message, "role")), ExtractFlexibleContentText(message), null);
            }
        }
        catch (JsonException)
        {
        }
    }

    private static void AppendExternalJsonlMessages(FileInfo file, List<CodexSessionMessage> messages)
    {
        foreach (var line in ReadJsonlLinesSafely(file.FullName))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                var role = NormalizeRole(GetString(root, "role"));
                var text = ExtractFlexibleContentText(root);

                if (string.IsNullOrWhiteSpace(text) && TryGetProperty(root, "message", out var message))
                {
                    role = string.IsNullOrWhiteSpace(role)
                        ? NormalizeRole(GetString(message, "role"))
                        : role;
                    text = ExtractFlexibleContentText(message);
                }

                if (string.IsNullOrWhiteSpace(role))
                {
                    role = NormalizeRole(GetString(root, "type"));
                }

                AddExternalMessage(messages, role, text, ParseTimestamp(GetString(root, "timestamp")));
            }
            catch (JsonException)
            {
            }
        }
    }

    private static void AddExternalMessage(
        List<CodexSessionMessage> messages,
        string role,
        string text,
        DateTimeOffset? timestamp)
    {
        if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        messages.Add(
            new CodexSessionMessage
            {
                Role = role,
                Text = text,
                Timestamp = timestamp
            });
    }

    private static string BuildTranscriptText(string filePath, AppLanguage language)
    {
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.Length > TranscriptTailReadBytes)
        {
            return BuildTranscriptTextFromTail(fileInfo, language);
        }

        return BuildTranscriptTextFromLines(ReadJsonlLinesSafely(filePath), language, isTailOnly: false);
    }

    private static string BuildTranscriptTextFromTail(FileInfo file, AppLanguage language)
    {
        using var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var bytesToRead = (int)Math.Min(TranscriptTailReadBytes, stream.Length);
        var offset = Math.Max(0, stream.Length - bytesToRead);
        var buffer = new byte[bytesToRead];

        stream.Seek(offset, SeekOrigin.Begin);
        var bytesRead = stream.Read(buffer, 0, buffer.Length);
        var text = Encoding.UTF8.GetString(buffer, 0, bytesRead);
        var lines = text.Split('\n');

        if (offset > 0 && lines.Length > 0)
        {
            lines = lines.Skip(1).ToArray();
        }

        return BuildTranscriptTextFromLines(lines, language, isTailOnly: offset > 0);
    }

    private static string BuildTranscriptTextFromLines(
        IEnumerable<string> lines,
        AppLanguage language,
        bool isTailOnly)
    {
        var transcript = new StringBuilder();
        var wasTrimmed = false;

        foreach (var line in lines)
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

                if (recordType != "response_item" || !TryGetProperty(root, "payload", out var payload))
                {
                    continue;
                }

                if (!string.Equals(GetString(payload, "type"), "message", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var role = GetString(payload, "role");

                if (!string.Equals(role, "user", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var text = RemoveEnvironmentContext(ExtractMessageText(payload));

                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                AppendTranscriptLine(
                    transcript,
                    ParseTimestamp(GetString(root, "timestamp")),
                    role,
                    text);
                wasTrimmed |= TrimTranscriptBuilder(transcript);
            }
            catch (JsonException)
            {
            }
        }

        if (transcript.Length == 0)
        {
            return GetLocalizedText(language, "NoTranscriptFound");
        }

        var result = transcript.ToString().Trim();
        if (isTailOnly || wasTrimmed)
        {
            var notice = language == AppLanguage.Russian
                ? "[Показана только последняя часть большого transcript-файла, чтобы AIHelper не зависал.]"
                : "[Only the latest part of this large transcript file is shown to keep AIHelper responsive.]";
            result = $"{notice}{Environment.NewLine}{Environment.NewLine}{result}";
        }

        return result;
    }

    private static bool TrimTranscriptBuilder(StringBuilder transcript)
    {
        if (transcript.Length <= MaxTranscriptCharacters)
        {
            return false;
        }

        var removeLength = transcript.Length - MaxTranscriptCharacters;
        transcript.Remove(0, removeLength);
        return true;
    }

    private SessionRecord CreateLockedSessionRecord(
        DiscoveredSessionFile sourceFile,
        AppLanguage language,
        IReadOnlyDictionary<string, string> titleLookup,
        SessionRecord? cachedSession)
    {
        var file = sourceFile.File;
        var sessionId = sourceFile.Kind == SessionFileKind.Codex
            ? DeriveSessionId(file.Name)
            : $"{sourceFile.ApplicationName.ToLowerInvariant()}:{Path.GetFileNameWithoutExtension(file.Name)}";
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);
        var snippet = CreateSnippetFromSession(cachedSession);
        if (snippet is null && TryGetSessionSnippet(file.FullName, sessionId, out var storedSnippet))
        {
            snippet = storedSnippet;
        }

        snippet ??= TryReadCodexHistorySnippet(sessionId, titleLookup);
        if (snippet is not null)
        {
            PersistLockedSnippetContext(sourceFile, sessionId, snippet, updatedAt.UtcDateTime);
        }

        var titleFromIndex = titleLookup.TryGetValue(sessionId, out var indexedTitle)
            ? indexedTitle
            : string.Empty;
        var title = snippet is not null && !string.IsNullOrWhiteSpace(snippet.Title)
            ? snippet.Title
            : !string.IsNullOrWhiteSpace(titleFromIndex)
                ? titleFromIndex
                : GetLocalizedText(language, "LockedTitle");
        var preview = snippet is not null && !string.IsNullOrWhiteSpace(snippet.Preview)
            ? BuildLockedSnippetPreview(snippet.Preview, language)
            : GetLocalizedText(language, "LockedPreview");
        var lastMessage = snippet is not null && !string.IsNullOrWhiteSpace(snippet.LastMessagePreview)
            ? snippet.LastMessagePreview
            : snippet is not null && !string.IsNullOrWhiteSpace(snippet.Preview)
                ? snippet.Preview
                : GetLocalizedText(language, "LockedLastMessage");
        var transcript = snippet is not null
            ? BuildLockedSnippetTranscript(snippet, language)
            : GetLocalizedText(language, "LockedTranscript");

        var baseSearchBlob = BuildSearchBlob(sessionId, file.FullName, "locked", title, preview, lastMessage);

        return new SessionRecord
        {
            SessionId = sessionId,
            Title = title,
            Preview = preview,
            LastMessagePreview = lastMessage,
            StartedAtText = snippet is not null && !string.IsNullOrWhiteSpace(snippet.StartedAtText)
                ? snippet.StartedAtText
                : GetLocalizedText(language, "Unknown"),
            UpdatedAtText = updatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"),
            DurationText = snippet is not null && !string.IsNullOrWhiteSpace(snippet.DurationText)
                ? snippet.DurationText
                : GetLocalizedText(language, "Unknown"),
            WorkingDirectory = snippet is not null && !string.IsNullOrWhiteSpace(snippet.WorkingDirectory)
                ? snippet.WorkingDirectory
                : "-",
            Source = sourceFile.ApplicationName,
            ModelProvider = snippet is not null && !string.IsNullOrWhiteSpace(snippet.ModelProvider)
                ? snippet.ModelProvider
                : "-",
            CliVersion = snippet is not null && !string.IsNullOrWhiteSpace(snippet.CliVersion)
                ? snippet.CliVersion
                : "-",
            FilePath = file.FullName,
            RelativePath = Path.GetRelativePath(sourceFile.RootPath, file.FullName),
            TranscriptText = transcript,
            UserMessageCount = snippet?.UserMessageCount ?? 0,
            AssistantMessageCount = snippet?.AssistantMessageCount ?? 0,
            ToolCallCount = snippet?.ToolCallCount ?? 0,
            TotalMessageCount = snippet?.TotalMessageCount ?? 0,
            UpdatedAtUtc = updatedAt.UtcDateTime,
            BaseSearchBlob = baseSearchBlob,
            SearchBlob = baseSearchBlob
        };
    }

    private void LoadSessionSnippets()
    {
        if (_sessionSnippetsLoaded)
        {
            return;
        }

        _sessionSnippetsLoaded = true;

        if (!File.Exists(SessionSnippetsPath))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(SessionSnippetsPath, Encoding.UTF8);
            var snippets = JsonSerializer.Deserialize<List<SessionSnippetRecord>>(json) ?? [];

            foreach (var snippet in snippets)
            {
                if (string.IsNullOrWhiteSpace(snippet.FilePath))
                {
                    continue;
                }

                _sessionSnippets[NormalizeSnippetKey(snippet.FilePath)] = snippet;
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
        {
            _sessionSnippets.Clear();
        }
    }

    private void SaveSessionSnippetsIfDirty()
    {
        if (!_sessionSnippetsDirty)
        {
            return;
        }

        try
        {
            var directoryPath = Path.GetDirectoryName(SessionSnippetsPath);
            if (!string.IsNullOrWhiteSpace(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            var snippets = _sessionSnippets.Values
                .OrderByDescending(snippet => snippet.UpdatedAtUtc)
                .Take(1000)
                .ToList();
            var json = JsonSerializer.Serialize(
                snippets,
                new JsonSerializerOptions
                {
                    WriteIndented = false
                });

            File.WriteAllText(SessionSnippetsPath, json, Encoding.UTF8);
            _sessionSnippetsDirty = false;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
        }
    }

    private void UpsertSessionSnippet(SessionRecord session, AppLanguage? language = null)
    {
        if (!ShouldPersistSessionSnippet(session))
        {
            return;
        }

        var snippet = CreateSnippetFromSession(session);
        if (snippet is null)
        {
            return;
        }

        var key = NormalizeSnippetKey(session.FilePath);
        _sessionSnippets.TryGetValue(key, out var existingSnippet);

        try
        {
            var file = new FileInfo(session.FilePath);
            snippet.FileLastWriteTicks = file.LastWriteTimeUtc.Ticks;
            snippet.FileLength = file.Length;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            snippet.FileLastWriteTicks = existingSnippet?.FileLastWriteTicks ?? 0;
            snippet.FileLength = existingSnippet?.FileLength ?? 0;
        }

        snippet.Language = language?.ToString() ?? existingSnippet?.Language ?? string.Empty;
        if (existingSnippet is not null &&
            existingSnippet.UpdatedAtUtc == snippet.UpdatedAtUtc &&
            existingSnippet.TotalMessageCount == snippet.TotalMessageCount &&
            existingSnippet.FileLastWriteTicks == snippet.FileLastWriteTicks &&
            existingSnippet.FileLength == snippet.FileLength &&
            string.Equals(existingSnippet.Language, snippet.Language, StringComparison.Ordinal) &&
            string.Equals(existingSnippet.Preview, snippet.Preview, StringComparison.Ordinal) &&
            string.Equals(existingSnippet.LastMessagePreview, snippet.LastMessagePreview, StringComparison.Ordinal) &&
            string.Equals(existingSnippet.TranscriptSnippet, snippet.TranscriptSnippet, StringComparison.Ordinal))
        {
            return;
        }

        _sessionSnippets[key] = snippet;
        _sessionSnippetsDirty = true;
    }

    private bool TryGetSessionSnippet(SessionRecord session, out SessionSnippetRecord snippet)
    {
        return TryGetSessionSnippet(session.FilePath, session.SessionId, out snippet);
    }

    private bool TryGetSessionSnippet(string filePath, string sessionId, out SessionSnippetRecord snippet)
    {
        if (!string.IsNullOrWhiteSpace(filePath) &&
            _sessionSnippets.TryGetValue(NormalizeSnippetKey(filePath), out snippet!) &&
            IsUsefulSnippet(snippet))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            foreach (var candidate in _sessionSnippets.Values)
            {
                if (string.Equals(candidate.SessionId, sessionId, StringComparison.OrdinalIgnoreCase) &&
                    IsUsefulSnippet(candidate))
                {
                    snippet = candidate;
                    return true;
                }
            }
        }

        snippet = null!;
        return false;
    }

    private static SessionSnippetRecord? CreateSnippetFromSession(SessionRecord? session)
    {
        if (session is null || !ShouldPersistSessionSnippet(session))
        {
            return null;
        }

        return new SessionSnippetRecord
        {
            SchemaVersion = SessionSnippetSchemaVersion,
            SessionId = session.SessionId,
            Title = session.DisplayTitle,
            OriginalTitle = session.Title,
            Preview = TrimPreview(session.Preview, 600),
            LastMessagePreview = TrimPreview(session.LastMessagePreview, 800),
            StartedAtText = session.StartedAtText,
            DurationText = session.DurationText,
            WorkingDirectory = session.WorkingDirectory,
            Source = session.Source,
            ModelProvider = session.ModelProvider,
            CliVersion = session.CliVersion,
            FilePath = session.FilePath,
            RelativePath = session.RelativePath,
            TranscriptSnippet = BuildTranscriptSnippet(session),
            UserMessageCount = session.UserMessageCount,
            AssistantMessageCount = session.AssistantMessageCount,
            ToolCallCount = session.ToolCallCount,
            TotalMessageCount = session.TotalMessageCount,
            UpdatedAtUtc = session.UpdatedAtUtc
        };
    }

    private static bool ShouldPersistSessionSnippet(SessionRecord session)
    {
        return !string.IsNullOrWhiteSpace(session.FilePath) &&
               !string.Equals(session.Title, "[locked session file]", StringComparison.OrdinalIgnoreCase) &&
               !session.Title.Contains("заблокирован", StringComparison.OrdinalIgnoreCase) &&
               (session.TotalMessageCount > 0 ||
                !string.IsNullOrWhiteSpace(session.Preview) ||
                !string.IsNullOrWhiteSpace(session.LastMessagePreview));
    }

    private static string BuildTranscriptSnippet(SessionRecord session)
    {
        if (!string.IsNullOrWhiteSpace(session.TranscriptText) &&
            !session.TranscriptText.Contains("locked", StringComparison.OrdinalIgnoreCase) &&
            !session.TranscriptText.Contains("заблокирован", StringComparison.OrdinalIgnoreCase))
        {
            return TrimTail(session.TranscriptText, 4000);
        }

        var builder = new StringBuilder();

        if (!string.IsNullOrWhiteSpace(session.Preview))
        {
            builder.AppendLine("Preview:");
            builder.AppendLine(session.Preview.Trim());
            builder.AppendLine();
        }

        if (!string.IsNullOrWhiteSpace(session.LastMessagePreview) &&
            !string.Equals(session.LastMessagePreview, session.Preview, StringComparison.Ordinal))
        {
            builder.AppendLine("Last message:");
            builder.AppendLine(session.LastMessagePreview.Trim());
        }

        return TrimTail(builder.ToString().Trim(), 4000);
    }

    private static bool IsUsefulSnippet(SessionSnippetRecord snippet)
    {
        return !string.IsNullOrWhiteSpace(snippet.Preview) ||
               !string.IsNullOrWhiteSpace(snippet.LastMessagePreview) ||
               !string.IsNullOrWhiteSpace(snippet.TranscriptSnippet) ||
               snippet.TotalMessageCount > 0;
    }

    private void PersistLockedSnippetContext(
        DiscoveredSessionFile sourceFile,
        string sessionId,
        SessionSnippetRecord snippet,
        DateTime updatedAtUtc)
    {
        if (!IsUsefulSnippet(snippet))
        {
            return;
        }

        snippet.SessionId = string.IsNullOrWhiteSpace(snippet.SessionId) ? sessionId : snippet.SessionId;
        snippet.FilePath = sourceFile.File.FullName;
        snippet.RelativePath = Path.GetRelativePath(sourceFile.RootPath, sourceFile.File.FullName);
        snippet.Source = string.IsNullOrWhiteSpace(snippet.Source) ? sourceFile.ApplicationName : snippet.Source;
        snippet.UpdatedAtUtc = snippet.UpdatedAtUtc == default ? updatedAtUtc : snippet.UpdatedAtUtc;

        var key = NormalizeSnippetKey(sourceFile.File.FullName);
        if (_sessionSnippets.TryGetValue(key, out var existingSnippet) &&
            existingSnippet.UpdatedAtUtc == snippet.UpdatedAtUtc &&
            existingSnippet.TotalMessageCount == snippet.TotalMessageCount &&
            string.Equals(existingSnippet.Preview, snippet.Preview, StringComparison.Ordinal) &&
            string.Equals(existingSnippet.LastMessagePreview, snippet.LastMessagePreview, StringComparison.Ordinal) &&
            string.Equals(existingSnippet.TranscriptSnippet, snippet.TranscriptSnippet, StringComparison.Ordinal))
        {
            return;
        }

        _sessionSnippets[key] = snippet;
        _sessionSnippetsDirty = true;
    }

    private static SessionSnippetRecord? TryReadCodexHistorySnippet(
        string sessionId,
        IReadOnlyDictionary<string, string> titleLookup)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !File.Exists(HistoryPath))
        {
            return null;
        }

        string firstText = string.Empty;
        string lastText = string.Empty;
        var count = 0;

        try
        {
            foreach (var line in ReadJsonlLinesSafely(HistoryPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                try
                {
                    using var document = JsonDocument.Parse(line);
                    var root = document.RootElement;

                    if (!string.Equals(GetString(root, "session_id"), sessionId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var text = GetString(root, "text");
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    count++;
                    if (string.IsNullOrWhiteSpace(firstText))
                    {
                        firstText = text;
                    }

                    lastText = text;
                }
                catch (JsonException)
                {
                }
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(firstText) && string.IsNullOrWhiteSpace(lastText))
        {
            return null;
        }

        titleLookup.TryGetValue(sessionId, out var title);
        var preview = string.IsNullOrWhiteSpace(firstText) ? lastText : firstText;
        var lastMessage = string.IsNullOrWhiteSpace(lastText) ? preview : lastText;

        return new SessionSnippetRecord
        {
            SessionId = sessionId,
            Title = string.IsNullOrWhiteSpace(title) ? TrimPreview(preview, 90) : title,
            Preview = TrimPreview(preview, 600),
            LastMessagePreview = TrimPreview(lastMessage, 800),
            Source = "Codex history",
            TranscriptSnippet = lastMessage,
            UserMessageCount = count,
            TotalMessageCount = count
        };
    }

    private static string BuildLockedSnippetPreview(string preview, AppLanguage language)
    {
        var prefix = language == AppLanguage.Russian
            ? "\u0424\u0430\u0439\u043b \u0437\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u043d. \u041f\u043e\u043a\u0430\u0437\u0430\u043d \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d\u043d\u044b\u0439 \u0444\u0440\u0430\u0433\u043c\u0435\u043d\u0442: "
            : "The file is locked. Showing saved context: ";

        return TrimPreview($"{prefix}{preview}", 220);
    }

    private static string BuildLockedSnippetTranscript(SessionSnippetRecord snippet, AppLanguage language)
    {
        var builder = new StringBuilder();
        builder.AppendLine(language == AppLanguage.Russian
            ? "\u0424\u0430\u0439\u043b \u0441\u0435\u0441\u0441\u0438\u0438 \u0441\u0435\u0439\u0447\u0430\u0441 \u0437\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u043d, \u043d\u043e AIHelper \u043d\u0430\u0448\u0451\u043b \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d\u043d\u044b\u0439 \u0444\u0440\u0430\u0433\u043c\u0435\u043d\u0442."
            : "The session file is locked right now, but AIHelper found saved context.");
        builder.AppendLine();

        if (!string.IsNullOrWhiteSpace(snippet.Title))
        {
            builder.AppendLine($"Title: {snippet.Title}");
        }

        if (!string.IsNullOrWhiteSpace(snippet.Preview))
        {
            builder.AppendLine();
            builder.AppendLine("Preview:");
            builder.AppendLine(snippet.Preview);
        }

        if (!string.IsNullOrWhiteSpace(snippet.LastMessagePreview) &&
            !string.Equals(snippet.LastMessagePreview, snippet.Preview, StringComparison.Ordinal))
        {
            builder.AppendLine();
            builder.AppendLine("Last message:");
            builder.AppendLine(snippet.LastMessagePreview);
        }

        if (!string.IsNullOrWhiteSpace(snippet.TranscriptSnippet) &&
            !string.Equals(snippet.TranscriptSnippet, snippet.Preview, StringComparison.Ordinal) &&
            !string.Equals(snippet.TranscriptSnippet, snippet.LastMessagePreview, StringComparison.Ordinal))
        {
            builder.AppendLine();
            builder.AppendLine("Saved transcript tail:");
            builder.AppendLine(snippet.TranscriptSnippet);
        }

        return builder.ToString().Trim();
    }

    private static string NormalizeSnippetKey(string filePath)
    {
        try
        {
            return Path.GetFullPath(filePath);
        }
        catch
        {
            return filePath;
        }
    }

    private static string TrimTail(string text, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length <= maxLength)
        {
            return text.Trim();
        }

        return text[^maxLength..].Trim();
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

    private static string CreateArchiveDestinationPath(SessionRecord session)
    {
        var sourceDirectoryName = SanitizeArchiveSegment(session.Source);
        var relativePath = string.IsNullOrWhiteSpace(session.RelativePath)
            ? Path.GetFileName(session.FilePath)
            : session.RelativePath;
        var safeRelativePath = SanitizeArchiveRelativePath(relativePath);
        var destinationPath = Path.Combine(ArchiveRootPath, sourceDirectoryName, safeRelativePath);
        var fullArchiveRoot = Path.GetFullPath(ArchiveRootPath);
        var fullDestination = Path.GetFullPath(destinationPath);

        if (!fullDestination.StartsWith(fullArchiveRoot, StringComparison.OrdinalIgnoreCase))
        {
            fullDestination = Path.Combine(
                fullArchiveRoot,
                sourceDirectoryName,
                $"{SanitizeArchiveSegment(session.SessionId)}{Path.GetExtension(session.FilePath)}");
        }

        return CreateUniqueFilePath(fullDestination);
    }

    private static string SanitizeArchiveRelativePath(string relativePath)
    {
        var parts = relativePath
            .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries)
            .Where(part => part != "." && part != "..")
            .Select(SanitizeArchiveSegment)
            .ToArray();

        return parts.Length == 0
            ? "session.jsonl"
            : Path.Combine(parts);
    }

    private static string SanitizeArchiveSegment(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var safe = new string(value
            .Select(character => invalidCharacters.Contains(character) ? '_' : character)
            .ToArray())
            .Trim();

        return string.IsNullOrWhiteSpace(safe) ? "unknown" : safe;
    }

    private static string CreateUniqueFilePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);

        for (var index = 1; index < 10_000; index++)
        {
            var candidate = Path.Combine(directory, $"{name}-{index}{extension}");

            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(directory, $"{name}-{DateTime.UtcNow:yyyyMMddHHmmssfff}{extension}");
    }

    private static CodexSessionConversation ParseConversationFile(
        FileInfo file,
        IReadOnlyDictionary<string, string> titleLookup,
        SessionRecord sourceSession)
    {
        var messages = new List<CodexSessionMessage>();
        string? sessionId = null;
        string? titleFromIndex = null;
        string firstPrompt = string.Empty;
        string workingDirectory = string.IsNullOrWhiteSpace(sourceSession.WorkingDirectory) || sourceSession.WorkingDirectory == "-"
            ? string.Empty
            : sourceSession.WorkingDirectory;
        string modelProvider = string.IsNullOrWhiteSpace(sourceSession.ModelProvider) || sourceSession.ModelProvider == "-"
            ? string.Empty
            : sourceSession.ModelProvider;
        DateTimeOffset? startedAt = null;

        foreach (var line in ReadJsonlLinesSafely(file.FullName))
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

                    var parsedWorkingDirectory = GetString(sessionPayload, "cwd");
                    if (!string.IsNullOrWhiteSpace(parsedWorkingDirectory))
                    {
                        workingDirectory = parsedWorkingDirectory;
                    }

                    var parsedModelProvider = GetString(sessionPayload, "model_provider");
                    if (!string.IsNullOrWhiteSpace(parsedModelProvider))
                    {
                        modelProvider = parsedModelProvider;
                    }

                    continue;
                }

                if (recordType != "response_item" || !TryGetProperty(root, "payload", out var payload))
                {
                    continue;
                }

                if (!string.Equals(GetString(payload, "type"), "message", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var role = GetString(payload, "role");

                if (!string.Equals(role, "user", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var text = RemoveEnvironmentContext(ExtractMessageText(payload));

                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(firstPrompt) &&
                    string.Equals(role, "user", StringComparison.OrdinalIgnoreCase))
                {
                    firstPrompt = text;
                }

                messages.Add(
                    new CodexSessionMessage
                    {
                        Role = role,
                        Text = text,
                        Timestamp = lineTimestamp
                    });
            }
            catch (JsonException)
            {
            }
        }

        var conversationTitle = !string.IsNullOrWhiteSpace(sourceSession.DisplayTitle)
            ? sourceSession.DisplayTitle
            : ChooseTitle(titleFromIndex, firstPrompt, file.Name);
        var updatedAt = new DateTimeOffset(file.LastWriteTimeUtc);

        return new CodexSessionConversation
        {
            SessionId = string.IsNullOrWhiteSpace(sessionId) ? sourceSession.SessionId : sessionId,
            Title = conversationTitle,
            WorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory)
                ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                : workingDirectory,
            ModelProvider = modelProvider,
            StartedAtUtc = startedAt,
            UpdatedAtUtc = updatedAt,
            Messages = messages
        };
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

    private static string ExtractFlexibleContentText(JsonElement element)
    {
        var directText = GetString(element, "text");
        if (!string.IsNullOrWhiteSpace(directText))
        {
            return CleanExternalText(directText);
        }

        var contentText = GetString(element, "content");
        if (!string.IsNullOrWhiteSpace(contentText))
        {
            return CleanExternalText(contentText);
        }

        if (TryGetProperty(element, "content", out var content))
        {
            var extracted = ExtractTextArray(content);
            if (!string.IsNullOrWhiteSpace(extracted))
            {
                return CleanExternalText(extracted);
            }
        }

        if (TryGetProperty(element, "parts", out var parts))
        {
            var extracted = ExtractTextArray(parts);
            if (!string.IsNullOrWhiteSpace(extracted))
            {
                return CleanExternalText(extracted);
            }
        }

        return string.Empty;
    }

    private static string ExtractPartsText(JsonElement message)
    {
        if (!TryGetProperty(message, "parts", out var parts))
        {
            return ExtractFlexibleContentText(message);
        }

        return CleanExternalText(ExtractTextArray(parts));
    }

    private static string ExtractTextArray(JsonElement arrayElement)
    {
        if (arrayElement.ValueKind == JsonValueKind.String)
        {
            return arrayElement.GetString() ?? string.Empty;
        }

        if (arrayElement.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();

        foreach (var item in arrayElement.EnumerateArray())
        {
            var text = GetString(item, "text");
            if (string.IsNullOrWhiteSpace(text) &&
                TryGetProperty(item, "functionCall", out var functionCall))
            {
                text = $"[tool call] {GetString(functionCall, "name")}";
            }
            else if (string.IsNullOrWhiteSpace(text) &&
                     TryGetProperty(item, "functionResponse", out var functionResponse))
            {
                text = $"[tool result] {GetString(functionResponse, "name")}";
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (builder.Length > 0)
            {
                builder.AppendLine();
            }

            builder.Append(text);
        }

        return builder.ToString();
    }

    private static string CleanExternalText(string text)
    {
        return text
            .Replace("<user_query>", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("</user_query>", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private static bool ContainsToolCall(JsonElement root)
    {
        return GetString(root, "type").Contains("tool", StringComparison.OrdinalIgnoreCase) ||
               (TryGetProperty(root, "message", out var message) &&
                TryGetProperty(message, "parts", out var parts) &&
                parts.ValueKind == JsonValueKind.Array &&
                parts.EnumerateArray().Any(part => TryGetProperty(part, "functionCall", out _)));
    }

    private static string NormalizeRole(string value)
    {
        if (string.Equals(value, "model", StringComparison.OrdinalIgnoreCase))
        {
            return "assistant";
        }

        if (string.Equals(value, "assistant", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "user", StringComparison.OrdinalIgnoreCase))
        {
            return value.ToLowerInvariant();
        }

        return string.Empty;
    }

    private static string NormalizeFileUriPath(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) || !uri.IsFile)
        {
            return value;
        }

        return Uri.UnescapeDataString(uri.LocalPath);
    }

    private static string InferProjectPathFromSourceFile(DiscoveredSessionFile sourceFile)
    {
        var path = sourceFile.File.FullName;
        var projectsMarker = $"{Path.DirectorySeparatorChar}projects{Path.DirectorySeparatorChar}";
        var markerIndex = path.IndexOf(projectsMarker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return string.Empty;
        }

        var afterMarker = path[(markerIndex + projectsMarker.Length)..];
        var firstSeparator = afterMarker.IndexOf(Path.DirectorySeparatorChar);
        if (firstSeparator <= 0)
        {
            return string.Empty;
        }

        var encodedProject = afterMarker[..firstSeparator];
        if (encodedProject.StartsWith("d--", StringComparison.OrdinalIgnoreCase))
        {
            return $"d:\\{encodedProject[3..].Replace('-', ' ')}";
        }

        if (encodedProject.StartsWith("c-", StringComparison.OrdinalIgnoreCase))
        {
            return $"C:\\{encodedProject[2..].Replace('-', '\\')}";
        }

        if (encodedProject.StartsWith("d-", StringComparison.OrdinalIgnoreCase))
        {
            return $"D:\\{encodedProject[2..].Replace('-', '\\')}";
        }

        return encodedProject.Replace('-', ' ');
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

        ThrowIfReparsePoint(filePath);
        var tempPath = $"{filePath}.tmp";

        try
        {
            using var writer = new StreamWriter(tempPath, false, Encoding.UTF8);

            foreach (var line in ReadJsonlLinesSafely(filePath))
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

    private static IEnumerable<string> ReadJsonlLinesSafely(string filePath)
    {
        if (IsReparsePoint(filePath))
        {
            yield break;
        }

        using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            bufferSize: 64 * 1024,
            FileOptions.SequentialScan);
        using var reader = new StreamReader(
            stream,
            Encoding.UTF8,
            detectEncodingFromByteOrderMarks: true,
            bufferSize: 64 * 1024,
            leaveOpen: false);

        var buffer = new char[8192];
        var line = new StringBuilder(Math.Min(MaxJsonlLineCharacters, 64 * 1024));
        var oversized = false;
        int read;

        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            for (var index = 0; index < read; index++)
            {
                var character = buffer[index];
                if (character == '\n')
                {
                    if (!oversized && line.Length > 0)
                    {
                        yield return line.ToString();
                    }

                    line.Clear();
                    oversized = false;
                    continue;
                }

                if (character == '\r' || oversized)
                {
                    continue;
                }

                if (line.Length >= MaxJsonlLineCharacters)
                {
                    line.Clear();
                    oversized = true;
                    continue;
                }

                line.Append(character);
            }
        }

        if (!oversized && line.Length > 0)
        {
            yield return line.ToString();
        }
    }

    private static bool IsReparsePoint(string filePath)
    {
        try
        {
            return (File.GetAttributes(filePath) & FileAttributes.ReparsePoint) != 0;
        }
        catch
        {
            return false;
        }
    }

    private static void ThrowIfReparsePoint(string filePath)
    {
        if (IsReparsePoint(filePath))
        {
            throw new UnauthorizedAccessException("AIHelper refuses to modify or import a session through a symbolic link or reparse point.");
        }
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
            (_, "LargeExternalSession") => language == AppLanguage.Russian
                ? "\u0424\u0430\u0439\u043b \u0441\u0435\u0441\u0441\u0438\u0438 \u043e\u0447\u0435\u043d\u044c \u0431\u043e\u043b\u044c\u0448\u043e\u0439; AIHelper \u043f\u043e\u043a\u0430\u0437\u044b\u0432\u0430\u0435\u0442 \u0435\u0433\u043e \u043a\u0430\u043a \u0432\u043d\u0435\u0448\u043d\u044e\u044e \u0441\u0435\u0441\u0441\u0438\u044e, \u0447\u0442\u043e\u0431\u044b \u043d\u0435 \u0437\u0430\u0432\u0438\u0441\u0430\u0442\u044c."
                : "This session file is very large; AIHelper shows it as an external session to stay responsive.",
            (_, "ExternalSessionFile") => language == AppLanguage.Russian
                ? "\u0412\u043d\u0435\u0448\u043d\u0438\u0439 \u0444\u0430\u0439\u043b \u0441\u0435\u0441\u0441\u0438\u0438. \u0415\u0433\u043e \u043c\u043e\u0436\u043d\u043e \u043e\u0442\u043a\u0440\u044b\u0442\u044c \u0438\u043b\u0438 \u043f\u043e\u0434\u043f\u0438\u0441\u0430\u0442\u044c \u0432 AIHelper."
                : "External session file. You can open it or label it in AIHelper.",
            (_, "TranscriptTrimmedNotice") => language == AppLanguage.Russian
                ? "[\u0427\u0430\u0441\u0442\u044c transcript \u043e\u0431\u0440\u0435\u0437\u0430\u043d\u0430, \u0447\u0442\u043e\u0431\u044b AIHelper \u043d\u0435 \u0437\u0430\u0432\u0438\u0441\u0430\u043b.]"
                : "[Part of the transcript was trimmed to keep AIHelper responsive.]",
            (_, "DurationLessThanMinute") => language == AppLanguage.Russian ? "< 1 \u043c\u0438\u043d" : "< 1 min",
            (_, "DurationMinute") => language == AppLanguage.Russian ? "\u043c\u0438\u043d" : "min",
            (_, "DurationHour") => language == AppLanguage.Russian ? "\u0447" : "h",
            (_, "DurationDay") => language == AppLanguage.Russian ? "\u0434" : "d",
            _ => key
        };
    }

    private enum SessionFileKind
    {
        Codex,
        Qwen,
        Continue,
        Cursor,
        ClaudeCode,
        GenericJson,
        GenericJsonl
    }

    private readonly record struct DiscoveredSessionFile(
        FileInfo File,
        string RootPath,
        string ApplicationName,
        SessionFileKind Kind);

    private readonly record struct SessionCacheKey(string FilePath, AppLanguage Language);

    private readonly record struct SessionCacheEntry(
        long FileLastWriteTicks,
        long FileLength,
        long ThreadTitleVersionTicks,
        long ThreadTitleVersionLength,
        SessionRecord Session);

    private sealed class SessionSnippetRecord
    {
        public int SchemaVersion { get; set; }

        public string SessionId { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;

        public string OriginalTitle { get; set; } = string.Empty;

        public string Preview { get; set; } = string.Empty;

        public string LastMessagePreview { get; set; } = string.Empty;

        public string StartedAtText { get; set; } = string.Empty;

        public string DurationText { get; set; } = string.Empty;

        public string WorkingDirectory { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public string ModelProvider { get; set; } = string.Empty;

        public string CliVersion { get; set; } = string.Empty;

        public string FilePath { get; set; } = string.Empty;

        public string RelativePath { get; set; } = string.Empty;

        public string TranscriptSnippet { get; set; } = string.Empty;

        public int UserMessageCount { get; set; }

        public int AssistantMessageCount { get; set; }

        public int ToolCallCount { get; set; }

        public int TotalMessageCount { get; set; }

        public DateTime UpdatedAtUtc { get; set; }

        public long FileLastWriteTicks { get; set; }

        public long FileLength { get; set; }

        public string Language { get; set; } = string.Empty;
    }

    private sealed record ThreadTitleCacheEntry(
        long VersionTicks,
        long VersionLength,
        Dictionary<string, string> Titles);
}
