using System.IO;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class AiExtensionCatalogService
{
    private readonly string _extensionsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "extensions.json");

    public IReadOnlyList<AiExtensionItem> LoadExtensions(LocalizationService? strings = null)
    {
        var presets = CreateDefaultPresets(strings);
        var storedItems = LoadStoredExtensions();
        var storedById = storedItems
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var presetIds = presets.Select(item => item.Id).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var preset in presets)
        {
            if (!storedById.TryGetValue(preset.Id, out var stored))
            {
                continue;
            }

            preset.IsInstalled = stored.IsInstalled;
            preset.IsEnabled = stored.IsEnabled;
        }

        var customItems = storedItems
            .Where(item => !presetIds.Contains(item.Id) && !string.IsNullOrWhiteSpace(item.Name))
            .Select(item =>
            {
                item.IsPreset = false;
                item.IsInstalled = item.IsInstalled || item.IsEnabled;
                return item;
            });

        return presets
            .Concat(customItems)
            .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First().Clone())
            .OrderBy(item => item.Kind)
            .ThenBy(item => item.IsPreset ? 0 : 1)
            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public void SaveExtensions(IEnumerable<AiExtensionItem> extensions)
    {
        var directory = Path.GetDirectoryName(_extensionsPath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var items = extensions
            .Select(item => item.Clone())
            .ToList();

        var json = JsonSerializer.Serialize(
            items,
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(_extensionsPath, json);
    }

    public string GetStoragePath() => _extensionsPath;

    private IReadOnlyList<AiExtensionItem> LoadStoredExtensions()
    {
        if (!File.Exists(_extensionsPath))
        {
            return [];
        }

        try
        {
            var json = File.ReadAllText(_extensionsPath);
            var items = JsonSerializer.Deserialize<List<AiExtensionItem>>(json) ?? [];

            return items
                .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                .Select(item =>
                {
                    item.Id = string.IsNullOrWhiteSpace(item.Id)
                        ? Guid.NewGuid().ToString("N")
                        : item.Id;
                    return item;
                })
                .ToList();
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            return [];
        }
    }

    private static IReadOnlyList<AiExtensionItem> CreateDefaultPresets(LocalizationService? strings)
    {
        return
        [
            new AiExtensionItem
            {
                Id = "preset-plugin-skill-installer",
                Name = "Skill Installer",
                Kind = AiExtensionKind.Plugin,
                Description = T(strings, "ExtensionPresetSkillInstallerDescription", "Installs and manages Codex skills from curated sources or GitHub repositories."),
                CommandOrUri = "codex skills install <skill-or-github-url>",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-plugin-hugging-face",
                Name = "Hugging Face",
                Kind = AiExtensionKind.Plugin,
                Description = T(strings, "ExtensionPresetHuggingFaceDescription", "Plugin for model, dataset, paper and Space workflows on Hugging Face."),
                CommandOrUri = "plugin:hugging-face",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-filesystem",
                Name = "Filesystem MCP",
                Kind = AiExtensionKind.Mcp,
                Description = T(strings, "ExtensionPresetFilesystemDescription", "Lets an assistant work with allowed local files through an MCP server."),
                CommandOrUri = "mcp:filesystem",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-github",
                Name = "GitHub MCP",
                Kind = AiExtensionKind.Mcp,
                Description = T(strings, "ExtensionPresetGithubDescription", "Connects assistant workflows to GitHub repositories, issues and pull requests."),
                CommandOrUri = "mcp:github",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-playwright",
                Name = "Playwright MCP",
                Kind = AiExtensionKind.Mcp,
                Description = T(strings, "ExtensionPresetPlaywrightDescription", "Provides browser automation for testing, screenshots and web interaction."),
                CommandOrUri = "mcp:playwright",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-canva",
                Name = "Canva MCP",
                Kind = AiExtensionKind.Mcp,
                Description = T(strings, "ExtensionPresetCanvaDescription", "Connects assistant workflows to Canva designs and assets when configured."),
                CommandOrUri = "mcp:canva",
                IsPreset = true
            }
        ];
    }

    private static string T(LocalizationService? strings, string key, string fallback)
    {
        return strings is null ? fallback : strings[key];
    }
}
