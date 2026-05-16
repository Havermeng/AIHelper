namespace LaptopSessionViewer.Models;

public sealed class CodexSessionConversation
{
    public required string SessionId { get; init; }

    public required string Title { get; init; }

    public required string WorkingDirectory { get; init; }

    public required string ModelProvider { get; init; }

    public required DateTimeOffset? StartedAtUtc { get; init; }

    public required DateTimeOffset UpdatedAtUtc { get; init; }

    public required IReadOnlyList<CodexSessionMessage> Messages { get; init; }
}

public sealed class CodexSessionMessage
{
    public required string Role { get; init; }

    public required string Text { get; init; }

    public required DateTimeOffset? Timestamp { get; init; }
}
