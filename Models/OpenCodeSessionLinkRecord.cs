namespace LaptopSessionViewer.Models;

public sealed class OpenCodeSessionLinkRecord
{
    public required string CodexSessionId { get; init; }

    public required string OpenCodeSessionId { get; init; }

    public required string OpenCodeTitle { get; init; }

    public required string WorkingDirectory { get; init; }

    public string? HandoffPath { get; init; }

    public required DateTime LinkedAtUtc { get; init; }

    public required DateTime CodexUpdatedAtUtc { get; init; }

    public required int CodexMessageCount { get; init; }
}
