using System.IO;
using System.Text.Json;

namespace LaptopSessionViewer.Services;

public sealed class SessionVisibilityService
{
    private readonly string _hiddenSessionsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "hidden-sessions.json");

    public HashSet<string> LoadHiddenSessions()
    {
        if (!File.Exists(_hiddenSessionsPath))
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        try
        {
            var json = File.ReadAllText(_hiddenSessionsPath);
            var ids = JsonSerializer.Deserialize<List<string>>(json) ?? [];
            return ids
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public void SaveHiddenSessions(IEnumerable<string> sessionIds)
    {
        var directoryPath = Path.GetDirectoryName(_hiddenSessionsPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var normalizedIds = sessionIds
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var json = JsonSerializer.Serialize(
            normalizedIds,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_hiddenSessionsPath, json);
    }
}
