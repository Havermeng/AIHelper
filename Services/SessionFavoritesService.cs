using System.IO;
using System.Text.Json;

namespace LaptopSessionViewer.Services;

public sealed class SessionFavoritesService
{
    private readonly string _favoritesPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".codex",
            "session_viewer_favorites.json");

    public HashSet<string> LoadFavorites()
    {
        if (!File.Exists(_favoritesPath))
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        try
        {
            var json = File.ReadAllText(_favoritesPath);
            var ids = JsonSerializer.Deserialize<List<string>>(json) ?? [];
            return ids
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }
        catch (JsonException)
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public void SaveFavorites(IEnumerable<string> sessionIds)
    {
        var directoryPath = Path.GetDirectoryName(_favoritesPath);

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

        File.WriteAllText(_favoritesPath, json);
    }
}
