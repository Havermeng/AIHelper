using System.Collections.Generic;

namespace AIHelper.VisualStudioExtension;

public sealed class AIHelperSessionFeed
{
    public string GeneratedAtUtc { get; set; } = string.Empty;

    public int SessionCount { get; set; }

    public List<AIHelperSessionItem> Sessions { get; set; } = new();
}

public sealed class AIHelperSessionItem
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

    public string UpdatedAtUtc { get; set; } = string.Empty;

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
