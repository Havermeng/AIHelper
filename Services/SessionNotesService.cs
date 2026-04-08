using System.IO;
using System.Text.Json;

namespace LaptopSessionViewer.Services;

public sealed class SessionNotesService
{
    private readonly string _notesPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".codex",
            "session_viewer_notes.json");

    public Dictionary<string, string> LoadNotes()
    {
        if (!File.Exists(_notesPath))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        try
        {
            var json = File.ReadAllText(_notesPath);
            var notes = JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            return notes
                .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && !string.IsNullOrWhiteSpace(pair.Value))
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception exception) when (exception is JsonException or IOException)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public void SaveNotes(IReadOnlyDictionary<string, string> notes)
    {
        var directoryPath = Path.GetDirectoryName(_notesPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var normalizedNotes = notes
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && !string.IsNullOrWhiteSpace(pair.Value))
            .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(pair => pair.Key, pair => pair.Value.Trim(), StringComparer.OrdinalIgnoreCase);

        var json = JsonSerializer.Serialize(
            normalizedNotes,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_notesPath, json);
    }
}
