using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace LaptopSessionViewer.Models;

public sealed class SessionRecord : INotifyPropertyChanged
{
    private bool _isFavorite;
    private string _note = string.Empty;
    private string _searchBlob = string.Empty;

    public required string SessionId { get; init; }
    public required string Title { get; init; }
    public required string Preview { get; init; }
    public required string LastMessagePreview { get; init; }
    public required string StartedAtText { get; init; }
    public required string UpdatedAtText { get; init; }
    public required string DurationText { get; init; }
    public required string WorkingDirectory { get; init; }
    public required string Source { get; init; }
    public required string ModelProvider { get; init; }
    public required string CliVersion { get; init; }
    public required string FilePath { get; init; }
    public required string RelativePath { get; init; }
    public required string TranscriptText { get; init; }
    public required int UserMessageCount { get; init; }
    public required int AssistantMessageCount { get; init; }
    public required int ToolCallCount { get; init; }
    public required int TotalMessageCount { get; init; }
    public required DateTime UpdatedAtUtc { get; init; }
    public required string BaseSearchBlob { get; init; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public bool IsFavorite
    {
        get => _isFavorite;
        set
        {
            if (SetField(ref _isFavorite, value))
            {
                OnPropertyChanged(nameof(FavoriteBadgeText));
            }
        }
    }

    public string Note
    {
        get => _note;
        set
        {
            if (SetField(ref _note, value))
            {
                OnPropertyChanged(nameof(HasNote));
                OnPropertyChanged(nameof(NotePreview));
                OnPropertyChanged(nameof(DisplayTitle));
                OnPropertyChanged(nameof(SecondaryTitle));
                OnPropertyChanged(nameof(HasSecondaryTitle));
            }
        }
    }

    public string SearchBlob
    {
        get => _searchBlob;
        set => SetField(ref _searchBlob, value);
    }

    public string MessageCountText => $"{TotalMessageCount} msgs";

    public string ToolCallCountText => $"{ToolCallCount} tools";

    public string ShortSessionId => SessionId.Length <= 12 ? SessionId : SessionId[..12];

    public string DisplayTitle => HasNote ? TrimPreview(Note, 90) : Title;

    public string SecondaryTitle => HasNote ? TrimPreview(Title, 90) : Preview;

    public bool HasSecondaryTitle => !string.IsNullOrWhiteSpace(SecondaryTitle);

    public string FavoriteBadgeText => IsFavorite ? "FAV" : string.Empty;

    public bool HasNote => !string.IsNullOrWhiteSpace(Note);

    public string NotePreview => HasNote ? TrimPreview(Note, 90) : string.Empty;

    private static string TrimPreview(string text, int maxLength)
    {
        var singleLine = text.Replace("\r", " ").Replace("\n", " ").Trim();

        if (singleLine.Length <= maxLength)
        {
            return singleLine;
        }

        return $"{singleLine[..maxLength].TrimEnd()}...";
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
