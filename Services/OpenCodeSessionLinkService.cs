using System.IO;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class OpenCodeSessionLinkService
{
    private readonly string _linksPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".codex",
            "session_viewer_opencode_links.json");

    public Dictionary<string, OpenCodeSessionLinkRecord> LoadLinks()
    {
        if (!File.Exists(_linksPath))
        {
            return new Dictionary<string, OpenCodeSessionLinkRecord>(StringComparer.OrdinalIgnoreCase);
        }

        try
        {
            var json = File.ReadAllText(_linksPath);
            var links = JsonSerializer.Deserialize<Dictionary<string, OpenCodeSessionLinkRecord>>(json)
                ?? new Dictionary<string, OpenCodeSessionLinkRecord>(StringComparer.OrdinalIgnoreCase);

            return links
                .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && pair.Value is not null)
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception exception) when (exception is JsonException or IOException)
        {
            return new Dictionary<string, OpenCodeSessionLinkRecord>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public void SaveLinks(IReadOnlyDictionary<string, OpenCodeSessionLinkRecord> links)
    {
        var directoryPath = Path.GetDirectoryName(_linksPath);

        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        var normalizedLinks = links
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && pair.Value is not null)
            .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

        var json = JsonSerializer.Serialize(
            normalizedLinks,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_linksPath, json);
    }
}
