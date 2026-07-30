using System.IO;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class SessionFeedExportService
{
    private readonly string _feedPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "sessions-feed.json");

    public string FeedPath => _feedPath;

    public void SaveSessions(IEnumerable<SessionRecord> sessions)
    {
        var directoryPath = Path.GetDirectoryName(_feedPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var feed = new SessionFeedDocument
        {
            GeneratedAtUtc = DateTime.UtcNow,
            Sessions = sessions
                .OrderByDescending(session => session.UpdatedAtUtc)
                .Select(session => new SessionFeedItem
                {
                    SessionId = session.SessionId,
                    Title = session.DisplayTitle,
                    OriginalTitle = session.Title,
                    Preview = session.Preview,
                    LastMessagePreview = session.LastMessagePreview,
                    Source = session.Source,
                    ModelProvider = session.ModelProvider,
                    CliVersion = session.CliVersion,
                    WorkingDirectory = session.WorkingDirectory,
                    FilePath = session.FilePath,
                    RelativePath = session.RelativePath,
                    UpdatedAtText = session.UpdatedAtText,
                    UpdatedAtUtc = session.UpdatedAtUtc,
                    StartedAtText = session.StartedAtText,
                    DurationText = session.DurationText,
                    IsFavorite = session.IsFavorite,
                    IsHidden = session.IsHidden,
                    CustomName = session.Note,
                    UserMessageCount = session.UserMessageCount,
                    AssistantMessageCount = session.AssistantMessageCount,
                    ToolCallCount = session.ToolCallCount,
                    TotalMessageCount = session.TotalMessageCount
                })
                .ToList()
        };

        feed.SessionCount = feed.Sessions.Count;

        var json = JsonSerializer.Serialize(
            feed,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_feedPath, json);
    }

    private sealed class SessionFeedDocument
    {
        public DateTime GeneratedAtUtc { get; set; }

        public int SessionCount { get; set; }

        public List<SessionFeedItem> Sessions { get; set; } = [];
    }

    private sealed class SessionFeedItem
    {
        public string SessionId { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;

        public string OriginalTitle { get; set; } = string.Empty;

        public string Preview { get; set; } = string.Empty;

        public string LastMessagePreview { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public string ModelProvider { get; set; } = string.Empty;

        public string CliVersion { get; set; } = string.Empty;

        public string WorkingDirectory { get; set; } = string.Empty;

        public string FilePath { get; set; } = string.Empty;

        public string RelativePath { get; set; } = string.Empty;

        public string UpdatedAtText { get; set; } = string.Empty;

        public DateTime UpdatedAtUtc { get; set; }

        public string StartedAtText { get; set; } = string.Empty;

        public string DurationText { get; set; } = string.Empty;

        public bool IsFavorite { get; set; }

        public bool IsHidden { get; set; }

        public string CustomName { get; set; } = string.Empty;

        public int UserMessageCount { get; set; }

        public int AssistantMessageCount { get; set; }

        public int ToolCallCount { get; set; }

        public int TotalMessageCount { get; set; }
    }
}
