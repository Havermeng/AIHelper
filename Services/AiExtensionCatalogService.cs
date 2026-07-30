using System.IO;
using System.Text.Json;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class AiExtensionCatalogService
{
    private static readonly HashSet<string> ObsoletePresetIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "preset-plugin-skill-installer",
        "preset-opencode-filesystem-mcp",
        "preset-lmstudio-local-mcp"
    };

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
            .Where(item =>
                !presetIds.Contains(item.Id) &&
                !ObsoletePresetIds.Contains(item.Id) &&
                !string.IsNullOrWhiteSpace(item.Name))
            .Select(item =>
            {
                if (!item.IsDetected)
                {
                    item.IsPreset = false;
                }

                item.TargetApp = NormalizeTargetApp(item);
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
                Id = "preset-plugin-hugging-face",
                Name = "Hugging Face",
                Kind = AiExtensionKind.Plugin,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetHuggingFaceDescription", "Plugin for model, dataset, paper and Space workflows on Hugging Face."),
                CommandOrUri = "codex plugin add hugging-face@openai-curated",
                RequestedAccess = T(strings, "ExtensionAccessHuggingFace", "Network access and the Hugging Face account or token only when a workflow needs it."),
                PackageVersion = "openai-curated",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-filesystem",
                Name = "Filesystem MCP",
                Kind = AiExtensionKind.Mcp,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetFilesystemDescription", "Lets an assistant work with allowed local files through an MCP server."),
                CommandOrUri = "codex mcp add aihelper-filesystem -- npx -y @modelcontextprotocol/server-filesystem@2026.7.10 \"%USERPROFILE%\\AIHelper Workspaces\"",
                RequestedAccess = T(strings, "ExtensionAccessFilesystem", "Read and write access is restricted to the AIHelper Workspaces folder."),
                PackageVersion = "2026.7.10",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-github",
                Name = "GitHub",
                Kind = AiExtensionKind.Plugin,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetGithubDescription", "Connects assistant workflows to GitHub repositories, issues and pull requests."),
                CommandOrUri = "codex plugin add github@openai-curated",
                RequestedAccess = T(strings, "ExtensionAccessGithub", "GitHub repositories, issues, and pull requests allowed by the account used during sign-in."),
                PackageVersion = "openai-curated",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-playwright",
                Name = "Playwright MCP",
                Kind = AiExtensionKind.Mcp,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetPlaywrightDescription", "Provides browser automation for testing, screenshots and web interaction."),
                CommandOrUri = "codex mcp add aihelper-playwright -- npx -y @playwright/mcp@0.0.78",
                RequestedAccess = T(strings, "ExtensionAccessPlaywright", "Can open websites and control an isolated browser. It does not receive unrestricted Windows desktop access."),
                PackageVersion = "0.0.78",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-mcp-canva",
                Name = "Canva",
                Kind = AiExtensionKind.Plugin,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetCanvaDescription", "Connects assistant workflows to Canva designs and assets when configured."),
                CommandOrUri = "codex plugin add canva@openai-curated",
                RequestedAccess = T(strings, "ExtensionAccessCanva", "Canva designs and assets allowed by the account used during sign-in."),
                PackageVersion = "openai-curated",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-skill-creator",
                Name = "Skill Creator",
                Kind = AiExtensionKind.Skill,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetSkillCreatorDescription", "Skill for creating your own skills. Installs into the shared skills folder used by both Codex and Claude Code."),
                CommandOrUri = "github.com/anthropics/skills → skill-creator",
                RequestedAccess = T(strings, "ExtensionAccessSharedSkill", "Files are copied from the official anthropics/skills GitHub repository into the local shared skills folder."),
                PackageVersion = "anthropics/skills",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-skill-pdf",
                Name = "PDF Skill",
                Kind = AiExtensionKind.Skill,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetSkillPdfDescription", "Skill for working with PDF files. Installs into the shared skills folder used by both Codex and Claude Code."),
                CommandOrUri = "github.com/anthropics/skills → pdf",
                RequestedAccess = T(strings, "ExtensionAccessSharedSkill", "Files are copied from the official anthropics/skills GitHub repository into the local shared skills folder."),
                PackageVersion = "anthropics/skills",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-skill-docx",
                Name = "Word (DOCX) Skill",
                Kind = AiExtensionKind.Skill,
                TargetApp = "Codex",
                Description = T(strings, "ExtensionPresetSkillDocxDescription", "Skill for working with Word documents. Installs into the shared skills folder used by both Codex and Claude Code."),
                CommandOrUri = "github.com/anthropics/skills → docx",
                RequestedAccess = T(strings, "ExtensionAccessSharedSkill", "Files are copied from the official anthropics/skills GitHub repository into the local shared skills folder."),
                PackageVersion = "anthropics/skills",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-opencode-session-bridge",
                Name = "AIHelper Session Bridge",
                Kind = AiExtensionKind.Plugin,
                TargetApp = "OpenCode",
                Description = T(strings, "ExtensionPresetOpenCodeBridgeDescription", "Prepares AIHelper session context files so OpenCode can continue compatible sessions."),
                CommandOrUri = "Built into AIHelper",
                RequestedAccess = T(strings, "ExtensionAccessSessionBridge", "Writes a local handoff file containing the selected session context. It does not upload it."),
                PackageVersion = "AIHelper",
                IsPreset = true
            },
            new AiExtensionItem
            {
                Id = "preset-lmstudio-openai-provider",
                Name = "LM Studio Local Provider",
                Kind = AiExtensionKind.Plugin,
                TargetApp = "LmStudio",
                Description = T(strings, "ExtensionPresetLmStudioProviderDescription", "OpenAI-compatible local provider endpoint for models served by LM Studio."),
                CommandOrUri = "http://127.0.0.1:1234/v1",
                RequestedAccess = T(strings, "ExtensionAccessLmStudio", "Connects only to the LM Studio server on this PC. Cloud models selected inside LM Studio may still use the internet."),
                PackageVersion = "Local endpoint",
                IsPreset = true
            }
        ];
    }

    private static string NormalizeTargetApp(AiExtensionItem item)
    {
        if (string.Equals(item.TargetApp, "OpenCode", StringComparison.OrdinalIgnoreCase))
        {
            return "OpenCode";
        }

        if (string.Equals(item.TargetApp, "LM Studio", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.TargetApp, "LmStudio", StringComparison.OrdinalIgnoreCase))
        {
            return "LmStudio";
        }

        if (item.Id.Contains("opencode", StringComparison.OrdinalIgnoreCase) ||
            item.Name.Contains("OpenCode", StringComparison.OrdinalIgnoreCase) ||
            item.CommandOrUri.Contains("opencode", StringComparison.OrdinalIgnoreCase))
        {
            return "OpenCode";
        }

        if (item.Id.Contains("lm-studio", StringComparison.OrdinalIgnoreCase) ||
            item.Id.Contains("lmstudio", StringComparison.OrdinalIgnoreCase) ||
            item.Name.Contains("LM Studio", StringComparison.OrdinalIgnoreCase) ||
            item.CommandOrUri.Contains("lmstudio", StringComparison.OrdinalIgnoreCase) ||
            item.CommandOrUri.Contains("localhost:1234", StringComparison.OrdinalIgnoreCase))
        {
            return "LmStudio";
        }

        return "Codex";
    }

    private static string T(LocalizationService? strings, string key, string fallback)
    {
        return strings is null ? fallback : strings[key];
    }
}
