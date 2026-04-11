using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class CodexEnvironmentService
{
    private const string OllamaWingetId = "Ollama.Ollama";
    private const string LmStudioWingetId = "ElementLabs.LMStudio";
    private const string ComfyUiDesktopWingetId = "Comfy.ComfyUI-Desktop";
    private const string PinokioWingetId = "pinokiocomputer.pinokio";
    private static readonly HashSet<string> ManagedCreativeToolPackages =
    [
        ComfyUiDesktopWingetId,
        PinokioWingetId
    ];
    private static readonly IReadOnlyList<string> KnownCodexModels =
    [
        "gpt-5.4",
        "gpt-5.4-mini",
        "gpt-5.3-codex",
        "gpt-5.3-codex-spark",
        "gpt-5.2"
    ];

    public string CodexCommandPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm", "codex.cmd");

    public string SessionsFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "sessions");

    public string CodexHomeFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex");

    public string ConfigFilePath =>
        Path.Combine(CodexHomeFolder, "config.toml");

    public string OpenClawHomeFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".openclaw");

    public string OpenClawConfigFilePath =>
        Path.Combine(OpenClawHomeFolder, "openclaw.json");

    public CodexEnvironmentSnapshot GetEnvironmentSnapshot()
    {
        var winget = RunCommand("cmd.exe", "/c winget --version");
        var node = RunCommand("cmd.exe", "/c node --version");
        var npm = RunCommand("cmd.exe", "/c npm --version");
        var git = RunCommand("cmd.exe", "/c git --version");

        var codexVersion = File.Exists(CodexCommandPath)
            ? RunCommand("cmd.exe", $"/c \"\"{CodexCommandPath}\" --version\"")
            : CommandResult.Missing;

        var loginStatus = File.Exists(CodexCommandPath)
            ? RunCommand("cmd.exe", $"/c \"\"{CodexCommandPath}\" login status\"")
            : CommandResult.Missing;
        var ollamaCommandPath = ResolveExecutableOnPath("ollama.exe");
        var ollamaPath = ResolveOllamaExecutablePath();
        var ollamaAppPath = ResolveOllamaAppExecutablePath();
        var lmStudioPath = ResolveLmStudioExecutablePath();
        var openClawPath = ResolveOpenClawExecutablePath();
        var openClawConfig = ReadOpenClawConfig(OpenClawConfigFilePath);
        var openClawNodeStatus = !string.IsNullOrWhiteSpace(openClawPath)
            ? RunCommand("cmd.exe", $"/c \"\"{openClawPath}\" node status\"", timeoutMilliseconds: 7000)
            : CommandResult.Missing;
        var openClawBrowserStatus = !string.IsNullOrWhiteSpace(openClawPath)
            ? RunCommand("cmd.exe", $"/c \"\"{openClawPath}\" browser status\"", timeoutMilliseconds: 7000)
            : CommandResult.Missing;
        var comfyUiPackage = GetInstalledWingetPackageInfo(ComfyUiDesktopWingetId);
        var pinokioPackage = GetInstalledWingetPackageInfo(PinokioWingetId);
        var ollamaVersion = !string.IsNullOrWhiteSpace(ollamaPath)
            ? RunCommand("cmd.exe", $"/c \"\"{ollamaPath}\" --version\"")
            : CommandResult.Missing;
        var openClawVersion = !string.IsNullOrWhiteSpace(openClawPath)
            ? RunCommand("cmd.exe", $"/c \"\"{openClawPath}\" --version\"")
            : CommandResult.Missing;
        var installedOllamaModels = GetInstalledOllamaModels(ollamaPath);
        var ollamaServerRunning = !string.IsNullOrWhiteSpace(ollamaPath) && IsLocalTcpEndpointReachable(11434);
        var ollamaTrayRunning = IsAnyProcessRunning("ollama");
        var ollamaModelsSummary = BuildOllamaModelsSummary(installedOllamaModels);

        return new CodexEnvironmentSnapshot
        {
            WingetAvailable = winget.Success,
            WingetVersion = winget.Output,
            NodeAvailable = node.Success,
            NodeVersion = node.Output,
            NpmAvailable = npm.Success,
            NpmVersion = npm.Output,
            GitAvailable = git.Success,
            GitVersion = git.Output,
            CodexAvailable = codexVersion.Success,
            CodexVersion = codexVersion.Output,
            LoggedIn =
                loginStatus.Success &&
                loginStatus.Output.Contains("logged in", StringComparison.OrdinalIgnoreCase),
            LoginStatus = loginStatus.Output,
            SessionsFolderExists = Directory.Exists(SessionsFolder),
            SessionsFolderPath = SessionsFolder,
            OllamaAvailable = !string.IsNullOrWhiteSpace(ollamaPath),
            OllamaCommandVisible = !string.IsNullOrWhiteSpace(ollamaCommandPath),
            OllamaExecutablePath = ollamaPath ?? string.Empty,
            OllamaAppAvailable = !string.IsNullOrWhiteSpace(ollamaAppPath),
            OllamaAppPath = ollamaAppPath ?? string.Empty,
            OllamaServerRunning = ollamaServerRunning,
            OllamaTrayRunning = ollamaTrayRunning,
            OllamaModelCount = installedOllamaModels.Count,
            OllamaModelsSummary = ollamaModelsSummary,
            OllamaDetail = !string.IsNullOrWhiteSpace(ollamaPath)
                ? BuildOllamaDetail(
                    ollamaVersion,
                    ollamaPath,
                    !string.IsNullOrWhiteSpace(ollamaCommandPath),
                    !string.IsNullOrWhiteSpace(ollamaAppPath),
                    ollamaServerRunning,
                    ollamaTrayRunning,
                    installedOllamaModels.Count)
                : "Ollama is not installed yet. Install it above, then refresh this page.",
            LmStudioAvailable = !string.IsNullOrWhiteSpace(lmStudioPath),
            LmStudioDetail = !string.IsNullOrWhiteSpace(lmStudioPath)
                ? $"{GetProductVersion(lmStudioPath)}{Environment.NewLine}{lmStudioPath}"
                : "LM Studio is not installed in the standard local paths.",
            ComfyUiDesktopAvailable = comfyUiPackage.IsInstalled,
            ComfyUiDesktopDetail = comfyUiPackage.Detail,
            PinokioAvailable = pinokioPackage.IsInstalled,
            PinokioDetail = pinokioPackage.Detail,
            OpenClawAvailable = !string.IsNullOrWhiteSpace(openClawPath),
            OpenClawDetail = !string.IsNullOrWhiteSpace(openClawPath)
                ? BuildOpenClawDetail(openClawVersion, openClawPath, openClawConfig)
                : "OpenClaw is not installed or not on PATH.",
            OpenClawConfigExists = openClawConfig.Exists,
            OpenClawConfigPath = openClawConfig.Path,
            OpenClawPrimaryModel = openClawConfig.PrimaryModel,
            OpenClawToolProfile = openClawConfig.ToolProfile,
            OpenClawWebSearchEnabled = openClawConfig.WebSearchEnabled,
            OpenClawTelegramConfigured = openClawConfig.TelegramConfigured,
            OpenClawNodeInstalled = DetermineOpenClawNodeInstalled(openClawNodeStatus),
            OpenClawNodeDetail = BuildOpenClawNodeDetail(openClawNodeStatus),
            OpenClawBrowserCliAvailable = !string.IsNullOrWhiteSpace(openClawPath),
            OpenClawBrowserReady = DetermineOpenClawBrowserReady(openClawBrowserStatus),
            OpenClawBrowserDetail = BuildOpenClawBrowserDetail(openClawBrowserStatus),
            InstalledOllamaModels = installedOllamaModels
        };
    }

    public CodexConfigInfo GetCodexConfigInfo()
    {
        if (!File.Exists(ConfigFilePath))
        {
            return new CodexConfigInfo
            {
                DefaultModel = string.Empty,
                AvailableModels = KnownCodexModels,
                Profiles = []
            };
        }

        var lines = File.ReadAllLines(ConfigFilePath);
        var currentSection = string.Empty;
        var defaultModel = string.Empty;
        var profiles = new List<string>();

        var profileSectionPattern =
            new Regex("^\\[profiles(?:\\.(?<plain>[^\\]]+)|\\.'(?<single>[^']+)'|\\.\"(?<double>[^\"]+)\")\\]$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        var modelPattern =
            new Regex("^model\\s*=\\s*[\"'](?<value>[^\"']+)[\"']$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        foreach (var rawLine in lines)
        {
            var trimmed = rawLine.Trim();

            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith('#'))
            {
                continue;
            }

            if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
            {
                currentSection = trimmed;
                var profileMatch = profileSectionPattern.Match(trimmed);

                if (profileMatch.Success)
                {
                    var profileName =
                        profileMatch.Groups["plain"].Value.Trim('\'', '"') +
                        profileMatch.Groups["single"].Value +
                        profileMatch.Groups["double"].Value;

                    if (!string.IsNullOrWhiteSpace(profileName))
                    {
                        profiles.Add(profileName);
                    }
                }

                continue;
            }

            if (string.IsNullOrWhiteSpace(currentSection) && string.IsNullOrWhiteSpace(defaultModel))
            {
                var modelMatch = modelPattern.Match(trimmed);

                if (modelMatch.Success)
                {
                    defaultModel = modelMatch.Groups["value"].Value.Trim();
                }
            }
        }

        return new CodexConfigInfo
        {
            DefaultModel = defaultModel,
            AvailableModels = BuildAvailableModels(defaultModel),
            Profiles = profiles
                .Where(profile => !string.IsNullOrWhiteSpace(profile))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(profile => profile, StringComparer.OrdinalIgnoreCase)
                .ToList()
        };
    }

    private static IReadOnlyList<string> BuildAvailableModels(string defaultModel)
    {
        var models = new List<string>();

        if (!string.IsNullOrWhiteSpace(defaultModel))
        {
            models.Add(defaultModel.Trim());
        }

        models.AddRange(KnownCodexModels);

        return models
            .Where(model => !string.IsNullOrWhiteSpace(model))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public string BuildInteractiveCommandPreview(NewSessionLaunchOptions options)
    {
        var args = BuildInteractiveArguments(options);
        return string.IsNullOrWhiteSpace(args) ? "codex" : $"codex {args}";
    }

    public void LaunchInteractiveSession(NewSessionLaunchOptions options)
    {
        var workingDirectory = Directory.Exists(options.WorkingDirectory)
            ? options.WorkingDirectory
            : Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        var arguments = BuildInteractiveArguments(options);
        var commandLine = new StringBuilder();
        commandLine.Append("/k cd /d ");
        commandLine.Append(QuoteForCommandLine(workingDirectory));
        commandLine.Append(" && call ");
        commandLine.Append(QuoteForCommandLine(CodexCommandPath));

        if (!string.IsNullOrWhiteSpace(arguments))
        {
            commandLine.Append(' ');
            commandLine.Append(arguments);
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = commandLine.ToString(),
                WorkingDirectory = workingDirectory,
                UseShellExecute = true
            });
    }

    public void LaunchCodexInstallRepairTerminal()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "aihelper-install-codex.ps1");
        File.WriteAllText(scriptPath, BuildInstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    public void LaunchPrerequisitesInstallTerminal()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "aihelper-install-prerequisites.ps1");
        File.WriteAllText(scriptPath, BuildPrerequisitesInstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    public void LaunchNodeInstallTerminal()
    {
        LaunchWingetInstallTerminal(
            "nodejs-lts",
            "OpenJS.NodeJS.LTS",
            "Node.js LTS");
    }

    public void LaunchGitInstallTerminal()
    {
        LaunchWingetInstallTerminal(
            "git",
            "Git.Git",
            "Git");
    }

    public void LaunchWingetRepairTerminal()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "aihelper-repair-winget.ps1");
        File.WriteAllText(scriptPath, BuildWingetRepairScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    public bool IsOllamaInstalled()
    {
        return !string.IsNullOrWhiteSpace(ResolveOllamaExecutablePath());
    }

    public void LaunchOllamaInstallTerminal()
    {
        var scriptPath = CreateTempScriptPath("aihelper-install-ollama");
        File.WriteAllText(scriptPath, BuildOllamaInstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    public void LaunchLmStudioInstallTerminal()
    {
        LaunchWingetInstallTerminal(
            "lmstudio",
            LmStudioWingetId,
            "LM Studio");
    }

    public void LaunchOllamaUninstallTerminal()
    {
        LaunchWingetUninstallTerminal(
            "ollama",
            OllamaWingetId,
            "Ollama");
    }

    public void LaunchOllamaApp()
    {
        var ollamaAppPath = ResolveOllamaAppExecutablePath();

        if (string.IsNullOrWhiteSpace(ollamaAppPath))
        {
            throw new FileNotFoundException("Ollama app was not found.");
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = ollamaAppPath,
                WorkingDirectory = Path.GetDirectoryName(ollamaAppPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    public void LaunchOllamaServeTerminal()
    {
        var ollamaPath = ResolveOllamaExecutablePath();

        if (string.IsNullOrWhiteSpace(ollamaPath))
        {
            throw new FileNotFoundException("ollama.exe was not found.");
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k call {QuoteForCommandLine(ollamaPath)} serve",
                WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    public void LaunchOllamaStopTerminal()
    {
        var scriptPath = CreateTempScriptPath("aihelper-stop-ollama");
        File.WriteAllText(scriptPath, BuildOllamaStopScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoLogo -NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true
            });
    }

    public void LaunchLmStudioUninstallTerminal()
    {
        LaunchWingetUninstallTerminal(
            "lmstudio",
            LmStudioWingetId,
            "LM Studio");
    }

    public void LaunchCreativeToolInstallTerminal(string packageId, string label)
    {
        EnsureManagedCreativeToolPackage(packageId);
        LaunchWingetInstallTerminal(
            BuildPackageScriptSuffix(packageId),
            packageId,
            label);
    }

    public void LaunchCreativeToolUninstallTerminal(string packageId, string label)
    {
        EnsureManagedCreativeToolPackage(packageId);
        LaunchWingetUninstallTerminal(
            BuildPackageScriptSuffix(packageId),
            packageId,
            label);
    }

    public void LaunchOpenClawInstallTerminal()
    {
        var scriptPath = CreateTempScriptPath("aihelper-install-openclaw");
        File.WriteAllText(scriptPath, BuildOpenClawInstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoLogo -NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true
            });
    }

    public void LaunchOpenClawUninstallTerminal()
    {
        var scriptPath = CreateTempScriptPath("aihelper-uninstall-openclaw");
        File.WriteAllText(scriptPath, BuildOpenClawUninstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoLogo -NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true
            });
    }

    public void LaunchOpenClawStatusTerminal()
    {
        LaunchOpenClawCommandTerminal(
            "aihelper-openclaw-status",
            "OpenClaw status",
            ["status", "--all"]);
    }

    public void LaunchOpenClawNodeInstallTerminal()
    {
        LaunchOpenClawCommandTerminal(
            "aihelper-openclaw-node-install",
            "OpenClaw node install",
            ["node", "install"]);
    }

    public void LaunchOpenClawNodeStatusTerminal()
    {
        LaunchOpenClawCommandTerminal(
            "aihelper-openclaw-node-status",
            "OpenClaw node status",
            ["node", "status"]);
    }

    public void LaunchOpenClawBrowserStatusTerminal()
    {
        LaunchOpenClawCommandTerminal(
            "aihelper-openclaw-browser-status",
            "OpenClaw browser status",
            ["browser", "status"]);
    }

    public OpenClawConfigApplyResult ApplyOpenClawQuickStartMode()
    {
        return ApplyOpenClawMode(
            "Quick local start",
            ResolvePreferredOllamaModel(["qwen2.5:3b"], "qwen2.5:3b"),
            "minimal");
    }

    public OpenClawConfigApplyResult ApplyOpenClawAdvancedMode()
    {
        return ApplyOpenClawMode(
            "Local advanced agent",
            ResolvePreferredOllamaModel(["qwen3.5:latest", "qwen3.5"], "qwen3.5:latest"),
            "coding");
    }

    public OpenClawConfigApplyResult PrepareOpenClawAlmostFullMode()
    {
        return ApplyOpenClawMode(
            "Almost full PC assistant",
            ResolvePreferredOllamaModel(["qwen3.5:latest", "qwen3.5"], "qwen3.5:latest"),
            "coding");
    }

    public void LaunchOllamaModelInstallTerminal(string modelTag)
    {
        var ollamaPath = ResolveOllamaExecutablePath();

        if (string.IsNullOrWhiteSpace(ollamaPath))
        {
            throw new FileNotFoundException("ollama.exe was not found.");
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k call {QuoteForCommandLine(ollamaPath)} pull {modelTag}",
                WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    public void LaunchOllamaModelRemoveTerminal(string modelTag)
    {
        var ollamaPath = ResolveOllamaExecutablePath();

        if (string.IsNullOrWhiteSpace(ollamaPath))
        {
            throw new FileNotFoundException("ollama.exe was not found.");
        }

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k call {QuoteForCommandLine(ollamaPath)} rm {modelTag}",
                WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    public void LaunchCodexLoginTerminal()
    {
        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k call {QuoteForCommandLine(CodexCommandPath)} login",
                WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    public static void OpenFolder(string path)
    {
        Process.Start(
            new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
    }

    private string BuildInteractiveArguments(NewSessionLaunchOptions options)
    {
        var arguments = new List<string>
        {
            "-C",
            QuoteForCommandLine(options.WorkingDirectory)
        };

        if (!string.IsNullOrWhiteSpace(options.Model))
        {
            arguments.Add("-m");
            arguments.Add(QuoteForCommandLine(options.Model.Trim()));
        }

        if (!string.IsNullOrWhiteSpace(options.Profile))
        {
            arguments.Add("-p");
            arguments.Add(QuoteForCommandLine(options.Profile.Trim()));
        }

        if (options.UseFullAuto)
        {
            arguments.Add("--full-auto");
        }
        else
        {
            if (!string.IsNullOrWhiteSpace(options.SandboxMode))
            {
                arguments.Add("-s");
                arguments.Add(options.SandboxMode.Trim());
            }

            if (!string.IsNullOrWhiteSpace(options.ApprovalPolicy))
            {
                arguments.Add("-a");
                arguments.Add(options.ApprovalPolicy.Trim());
            }
        }

        if (options.UseSearch)
        {
            arguments.Add("--search");
        }

        if (options.UseOss)
        {
            arguments.Add("--oss");
        }

        if (options.UseOss && !string.IsNullOrWhiteSpace(options.LocalProvider))
        {
            arguments.Add("--local-provider");
            arguments.Add(options.LocalProvider.Trim());
        }

        if (!string.IsNullOrWhiteSpace(options.Prompt))
        {
            arguments.Add(QuoteForCommandLine(options.Prompt.Trim()));
        }

        return string.Join(" ", arguments);
    }

    private void LaunchWingetInstallTerminal(string scriptSuffix, string packageId, string label)
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"aihelper-install-{scriptSuffix}.ps1");
        File.WriteAllText(scriptPath, BuildWingetInstallScript(packageId, label), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    private void LaunchWingetUninstallTerminal(string scriptSuffix, string packageId, string label)
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"aihelper-uninstall-{scriptSuffix}.ps1");
        File.WriteAllText(scriptPath, BuildWingetUninstallScript(packageId, label), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    private void LaunchOpenClawCommandTerminal(string scriptPrefix, string title, IReadOnlyList<string> arguments)
    {
        var openClawPath = ResolveOpenClawExecutablePath();

        if (string.IsNullOrWhiteSpace(openClawPath))
        {
            throw new FileNotFoundException("openclaw was not found.");
        }

        var scriptPath = CreateTempScriptPath(scriptPrefix);
        File.WriteAllText(scriptPath, BuildOpenClawCommandScript(openClawPath, title, arguments), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoLogo -NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                UseShellExecute = true
            });
    }

    private OpenClawConfigApplyResult ApplyOpenClawMode(string modeName, string primaryModelTag, string toolsProfile)
    {
        Directory.CreateDirectory(OpenClawHomeFolder);

        JsonObject root;
        string backupPath = string.Empty;

        if (File.Exists(OpenClawConfigFilePath))
        {
            root = LoadOpenClawConfigRoot();
            backupPath = BackupOpenClawConfig();
        }
        else
        {
            root = new JsonObject();
        }

        var agents = EnsureObject(root, "agents");
        var defaults = EnsureObject(agents, "defaults");
        defaults["contextTokens"] = 32768;

        var model = EnsureObject(defaults, "model");
        model["primary"] = BuildOllamaPrimaryModel(primaryModelTag);
        model["contextWindow"] = 32768;
        model["maxTokens"] = 4096;

        var tools = EnsureObject(root, "tools");
        tools["profile"] = toolsProfile;
        tools["alsoAllow"] = new JsonArray("web_search");

        var web = EnsureObject(tools, "web");
        var search = EnsureObject(web, "search");
        search["provider"] = "ollama";

        var aiHelper = EnsureObject(root, "aihelper");
        aiHelper["lastAppliedOpenClawMode"] = modeName;
        aiHelper["lastUpdatedAtUtc"] = DateTime.UtcNow.ToString("O");

        SaveOpenClawConfigRoot(root);

        return new OpenClawConfigApplyResult
        {
            ModeName = modeName,
            ConfigPath = OpenClawConfigFilePath,
            BackupPath = backupPath,
            PrimaryModel = BuildOllamaPrimaryModel(primaryModelTag),
            ToolsProfile = toolsProfile
        };
    }

    private static void EnsureManagedCreativeToolPackage(string packageId)
    {
        if (!ManagedCreativeToolPackages.Contains(packageId))
        {
            throw new InvalidOperationException($"Unsupported creative AI package: {packageId}");
        }
    }

    private static string BuildPackageScriptSuffix(string packageId)
    {
        var normalized = Regex.Replace(packageId, "[^a-zA-Z0-9]+", "-")
            .Trim('-')
            .ToLowerInvariant();

        return string.IsNullOrWhiteSpace(normalized) ? "package" : normalized;
    }

    private static string CreateTempScriptPath(string scriptPrefix)
    {
        return Path.Combine(Path.GetTempPath(), $"{scriptPrefix}-{Guid.NewGuid():N}.ps1");
    }

    private static string BuildInstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host "== $Text ==" -ForegroundColor Cyan
}

function Ensure-WingetPackage([string]$Id, [string]$Label) {
    if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
        Write-Host "winget.exe not found. Skipping $Label." -ForegroundColor Yellow
        return
    }

    Write-Step "Installing or updating $Label"
    winget install --id $Id -e --accept-package-agreements --accept-source-agreements
}

function Resolve-NpmCmd {
    $command = Get-Command npm.cmd -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallback = Join-Path ${env:ProgramFiles} 'nodejs\npm.cmd'
    if (Test-Path $fallback) {
        return $fallback
    }

    return $null
}

Write-Step 'Installing prerequisites'
Ensure-WingetPackage 'OpenJS.NodeJS.LTS' 'Node.js LTS'
Ensure-WingetPackage 'Git.Git' 'Git'

$npmCmd = Resolve-NpmCmd
if (-not $npmCmd) {
    throw 'npm.cmd not found after prerequisite installation.'
}

Write-Step 'Installing or updating OpenAI Codex CLI'
& $npmCmd install -g @openai/codex

Write-Step 'Finished'
Write-Host 'Codex stack install/update completed.' -ForegroundColor Green
Write-Host 'Next step: run `codex login` if the CLI is not authenticated yet.' -ForegroundColor Green
""";
    }

    private static string BuildPrerequisitesInstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host "== $Text ==" -ForegroundColor Cyan
}

function Ensure-WingetPackage([string]$Id, [string]$Label) {
    if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
        throw 'winget.exe is not available on this system.'
    }

    Write-Step "Installing or updating $Label"
    winget install --id $Id -e --accept-package-agreements --accept-source-agreements
}

Ensure-WingetPackage 'OpenJS.NodeJS.LTS' 'Node.js LTS'
Ensure-WingetPackage 'Git.Git' 'Git'

Write-Host ''
Write-Host 'Node.js and Git install/update command finished.' -ForegroundColor Green
Write-Host 'npm is installed together with Node.js LTS.' -ForegroundColor Green
Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green
""";
    }

    private static string BuildOllamaInstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
    throw 'winget.exe is not available on this system.'
}

Write-Host ''
Write-Host '== Installing or updating Ollama ==' -ForegroundColor Cyan
winget install --id Ollama.Ollama -e --accept-package-agreements --accept-source-agreements

Write-Host ''
Write-Host 'Ollama install/update command finished.' -ForegroundColor Green
Write-Host 'What to do next:' -ForegroundColor Cyan
Write-Host '1. Return to AIHelper and refresh the Local AI status.' -ForegroundColor Yellow
Write-Host '2. If old terminals still cannot see ollama, open a new terminal window.' -ForegroundColor Yellow
Write-Host '3. A large main window is not required. Use AIHelper to start the local server if needed.' -ForegroundColor Yellow
Write-Host '4. Download your first local model before connecting Ollama to OpenClaw.' -ForegroundColor Yellow
""";
    }

    private static string BuildOpenClawCommandScript(
        string openClawPath,
        string title,
        IReadOnlyList<string> arguments)
    {
        var escapedPath = EscapeForSingleQuotedPowerShell(openClawPath);
        var escapedTitle = EscapeForSingleQuotedPowerShell(title);
        var argumentLiteral = string.Join(
            ", ",
            arguments.Select(argument => $"'{EscapeForSingleQuotedPowerShell(argument)}'"));

        return $$"""
$ErrorActionPreference = 'Stop'

$openClaw = '{{escapedPath}}'
$arguments = @({{argumentLiteral}})

Write-Host ''
Write-Host '== {{escapedTitle}} ==' -ForegroundColor Cyan
& $openClaw @arguments

Write-Host ''
Write-Host 'The terminal remains open for inspection.' -ForegroundColor Green
""";
    }

    private static string BuildOllamaStopScript()
    {
        return """
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '== Stopping Ollama processes ==' -ForegroundColor Yellow

$targets = Get-Process | Where-Object { $_.ProcessName -match 'ollama' }

if (-not $targets) {
    Write-Host 'No Ollama processes are running right now.' -ForegroundColor Yellow
} else {
    $targets | Stop-Process -Force
    Write-Host 'Ollama processes were stopped.' -ForegroundColor Green
}

Write-Host ''
Write-Host 'Return to AIHelper and refresh the Local AI status.' -ForegroundColor Green
""";
    }

    private static string BuildWingetInstallScript(string packageId, string label)
    {
        return
            "$ErrorActionPreference = 'Stop'" + Environment.NewLine +
            Environment.NewLine +
            "if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {" + Environment.NewLine +
            "    throw 'winget.exe is not available on this system.'" + Environment.NewLine +
            "}" + Environment.NewLine +
            Environment.NewLine +
            "Write-Host ''" + Environment.NewLine +
            $"Write-Host '== Installing or updating {label} ==' -ForegroundColor Cyan" + Environment.NewLine +
            $"winget install --id {packageId} -e --accept-package-agreements --accept-source-agreements" + Environment.NewLine +
            Environment.NewLine +
            "Write-Host ''" + Environment.NewLine +
            $"Write-Host '{label} install/update command finished.' -ForegroundColor Green" + Environment.NewLine +
            "Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green" + Environment.NewLine;
    }

    private static string BuildWingetUninstallScript(string packageId, string label)
    {
        return
            "$ErrorActionPreference = 'Stop'" + Environment.NewLine +
            Environment.NewLine +
            "if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {" + Environment.NewLine +
            "    throw 'winget.exe is not available on this system.'" + Environment.NewLine +
            "}" + Environment.NewLine +
            Environment.NewLine +
            "Write-Host ''" + Environment.NewLine +
            $"Write-Host '== Removing {label} ==' -ForegroundColor Yellow" + Environment.NewLine +
            $"winget uninstall --id {packageId} -e --accept-source-agreements" + Environment.NewLine +
            Environment.NewLine +
            "Write-Host ''" + Environment.NewLine +
            $"Write-Host '{label} uninstall command finished.' -ForegroundColor Green" + Environment.NewLine +
            "Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green" + Environment.NewLine;
    }

    private static string BuildWingetRepairScript()
    {
        return """
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '== Restoring WinGet / App Installer ==' -ForegroundColor Cyan

Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe

if (Get-Command winget.exe -ErrorAction SilentlyContinue) {
    Write-Host ''
    Write-Host 'WinGet is available now.' -ForegroundColor Green
} else {
    Write-Host ''
    Write-Host 'WinGet is still unavailable. Opening the official Microsoft Learn page.' -ForegroundColor Yellow
    Start-Process 'https://learn.microsoft.com/en-us/windows/package-manager/winget/'
}

Write-Host ''
Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green
""";
    }

    private static string BuildOpenClawInstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host "== $Text ==" -ForegroundColor Cyan
}

function Resolve-NodeCommand {
    $command = Get-Command node.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $command = Get-Command node -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallback = Join-Path ${env:ProgramFiles} 'nodejs\node.exe'
    if (Test-Path $fallback) {
        return $fallback
    }

    return $null
}

function Resolve-NpmCommand {
    $command = Get-Command npm.cmd -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $command = Get-Command npm.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallback = Join-Path ${env:ProgramFiles} 'nodejs\npm.cmd'
    if (Test-Path $fallback) {
        return $fallback
    }

    return $null
}

function Resolve-GitCommand {
    $command = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $command = Get-Command git -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallback = Join-Path ${env:ProgramFiles} 'Git\cmd\git.exe'
    if (Test-Path $fallback) {
        return $fallback
    }

    return $null
}

function Resolve-OpenClawCommand {
    foreach ($candidate in @('openclaw.cmd', 'openclaw')) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }

    foreach ($fallback in @(
        (Join-Path $env:APPDATA 'npm\openclaw.cmd'),
        (Join-Path $env:USERPROFILE '.local\bin\openclaw.cmd')
    )) {
        if (Test-Path $fallback) {
            return $fallback
        }
    }

    return $null
}

function Add-UserPathEntryIfMissing([string]$PathEntry) {
    if ([string]::IsNullOrWhiteSpace($PathEntry) -or -not (Test-Path $PathEntry)) {
        return
    }

    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $entries = @($userPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($entries | Where-Object { $_ -ieq $PathEntry }) {
        return
    }

    $newUserPath = if ([string]::IsNullOrWhiteSpace($userPath)) {
        $PathEntry
    } else {
        "$userPath;$PathEntry"
    }

    [Environment]::SetEnvironmentVariable('Path', $newUserPath, 'User')
    $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + $newUserPath
    Write-Host "[OK] Added $PathEntry to user PATH" -ForegroundColor Green
}

Write-Step 'Checking Node.js and npm'
$nodeCmd = Resolve-NodeCommand
$npmCmd = Resolve-NpmCommand

if (-not $nodeCmd -or -not $npmCmd) {
    throw 'Node.js and npm are required before installing OpenClaw. Install the base components in AIHelper first.'
}

$nodeVersion = (& $nodeCmd --version 2>$null).Trim()
if ([string]::IsNullOrWhiteSpace($nodeVersion)) {
    throw 'Node.js was found, but its version could not be read.'
}

$majorString = ($nodeVersion -replace '^v', '').Split('.')[0]
$majorVersion = 0
if (-not [int]::TryParse($majorString, [ref]$majorVersion) -or $majorVersion -lt 22) {
    throw "OpenClaw requires Node.js 22+. Current version: $nodeVersion"
}

Write-Host "[OK] Node.js $nodeVersion found" -ForegroundColor Green
Write-Host "[OK] npm command: $npmCmd" -ForegroundColor Green

Write-Step 'Checking Git'
$gitCmd = Resolve-GitCommand
if (-not $gitCmd) {
    throw 'Git is required before installing OpenClaw. Install Git in AIHelper first, then try again.'
}

Write-Host "[OK] git command: $gitCmd" -ForegroundColor Green

Write-Step 'Installing OpenClaw from npm'
Write-Host 'npm install can stay quiet for a while during package download and unpacking.' -ForegroundColor Yellow
Write-Host 'Wait for the next step header or an explicit error before closing this terminal.' -ForegroundColor Yellow

$prevScriptShell = $env:NPM_CONFIG_SCRIPT_SHELL
$prevLogLevel = $env:NPM_CONFIG_LOGLEVEL
$prevUpdateNotifier = $env:NPM_CONFIG_UPDATE_NOTIFIER
$prevFund = $env:NPM_CONFIG_FUND
$prevAudit = $env:NPM_CONFIG_AUDIT
$prevNodeLlamaSkipDownload = $env:NODE_LLAMA_CPP_SKIP_DOWNLOAD

try {
    $env:NPM_CONFIG_SCRIPT_SHELL = 'cmd.exe'
    $env:NPM_CONFIG_LOGLEVEL = 'info'
    $env:NPM_CONFIG_UPDATE_NOTIFIER = 'false'
    $env:NPM_CONFIG_FUND = 'false'
    $env:NPM_CONFIG_AUDIT = 'false'
    $env:NODE_LLAMA_CPP_SKIP_DOWNLOAD = '1'
    & $npmCmd install -g openclaw@latest --loglevel info

    if ($LASTEXITCODE -ne 0) {
        throw "npm install finished with exit code $LASTEXITCODE."
    }
} finally {
    $env:NPM_CONFIG_SCRIPT_SHELL = $prevScriptShell
    $env:NPM_CONFIG_LOGLEVEL = $prevLogLevel
    $env:NPM_CONFIG_UPDATE_NOTIFIER = $prevUpdateNotifier
    $env:NPM_CONFIG_FUND = $prevFund
    $env:NPM_CONFIG_AUDIT = $prevAudit
    $env:NODE_LLAMA_CPP_SKIP_DOWNLOAD = $prevNodeLlamaSkipDownload
}

$npmPrefix = (& $npmCmd config get prefix 2>$null).Trim()
if (-not [string]::IsNullOrWhiteSpace($npmPrefix)) {
    Add-UserPathEntryIfMissing $npmPrefix
}

$openClawCmd = Resolve-OpenClawCommand
if (-not $openClawCmd) {
    throw 'OpenClaw was installed, but the openclaw command is still not visible. Restart AIHelper or open a new terminal and try again.'
}

Write-Host "[OK] OpenClaw command: $openClawCmd" -ForegroundColor Green

Write-Step 'Starting OpenClaw onboard'
Write-Host 'Interactive OpenClaw setup starts now.' -ForegroundColor Cyan
Write-Host 'If OpenClaw asks additional questions, complete them in this same terminal.' -ForegroundColor Yellow
& $openClawCmd onboard

if ($LASTEXITCODE -ne 0) {
    throw "OpenClaw onboard finished with exit code $LASTEXITCODE."
}

Write-Host ''
Write-Host 'OpenClaw install/update command finished.' -ForegroundColor Green
Write-Host 'Return to AIHelper. The status block should refresh automatically.' -ForegroundColor Green
""";
    }

    private static string BuildOpenClawUninstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

function Resolve-OpenClawCommand {
    $command = Get-Command openclaw.cmd -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $command = Get-Command openclaw -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallbacks = @(
        (Join-Path $env:USERPROFILE '.local\bin\openclaw.cmd'),
        (Join-Path $env:APPDATA 'npm\openclaw.cmd')
    )

    foreach ($fallback in $fallbacks) {
        if (Test-Path $fallback) {
            return $fallback
        }
    }

    return $null
}

function Resolve-NpmCmd {
    $command = Get-Command npm.cmd -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallback = Join-Path ${env:ProgramFiles} 'nodejs\npm.cmd'
    if (Test-Path $fallback) {
        return $fallback
    }

    return $null
}

Write-Host ''
Write-Host '== Removing OpenClaw ==' -ForegroundColor Yellow

$openClaw = Resolve-OpenClawCommand
if ($openClaw) {
    & $openClaw uninstall --all --yes --non-interactive
} else {
    Write-Host 'OpenClaw command not found. Trying npx cleanup.' -ForegroundColor Yellow
    npx -y openclaw uninstall --all --yes --non-interactive
}

$npmCmd = Resolve-NpmCmd
if ($npmCmd) {
    & $npmCmd rm -g openclaw
} else {
    Write-Host 'npm.cmd not found, skipping global package removal.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'OpenClaw uninstall command finished.' -ForegroundColor Green
Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green
""";
    }

    private static IReadOnlyDictionary<string, string> GetInstalledOllamaModels(string? ollamaPath)
    {
        if (string.IsNullOrWhiteSpace(ollamaPath))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var modelsResult = RunCommand("cmd.exe", $"/c \"\"{ollamaPath}\" ls\"");

        if (!modelsResult.Success || string.IsNullOrWhiteSpace(modelsResult.Output))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var installedModels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lines = modelsResult.Output.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var line in lines)
        {
            if (line.StartsWith("NAME", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var columns = Regex.Split(line, "\\s{2,}")
                .Where(column => !string.IsNullOrWhiteSpace(column))
                .ToArray();

            if (columns.Length < 3)
            {
                continue;
            }

            installedModels[columns[0].Trim()] = columns[2].Trim();
        }

        return installedModels;
    }

    private static WingetPackageInfo GetInstalledWingetPackageInfo(string packageId)
    {
        var listResult = RunCommand(
            "cmd.exe",
            $"/c winget list --id {packageId} -e --accept-source-agreements --disable-interactivity");

        if (!listResult.Success ||
            listResult.Output.Contains("No installed package found", StringComparison.OrdinalIgnoreCase))
        {
            return WingetPackageInfo.Missing;
        }

        var outputLines = listResult.Output
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(line => !line.StartsWith("Name", StringComparison.OrdinalIgnoreCase))
            .Where(line => !Regex.IsMatch(line, "^-+$"))
            .ToList();

        var packageLine = outputLines
            .FirstOrDefault(line => line.Contains(packageId, StringComparison.OrdinalIgnoreCase));

        if (string.IsNullOrWhiteSpace(packageLine))
        {
            return new WingetPackageInfo(true, packageId);
        }

        var columns = Regex.Split(packageLine.Trim(), "\\s{2,}")
            .Where(column => !string.IsNullOrWhiteSpace(column))
            .ToArray();
        var packageIdIndex = Array.FindIndex(
            columns,
            column => string.Equals(column, packageId, StringComparison.OrdinalIgnoreCase));
        var version = packageIdIndex >= 0 && packageIdIndex + 1 < columns.Length
            ? columns[packageIdIndex + 1].Trim()
            : string.Empty;
        var detail = string.IsNullOrWhiteSpace(version)
            ? packageId
            : $"{version}{Environment.NewLine}{packageId}";

        return new WingetPackageInfo(true, detail);
    }

    private static CommandResult RunCommand(string fileName, string arguments, int timeoutMilliseconds = 15000)
    {
        try
        {
            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            process.Start();
            var standardOutput = process.StandardOutput.ReadToEnd();
            var standardError = process.StandardError.ReadToEnd();
            process.WaitForExit(timeoutMilliseconds);

            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                return new CommandResult(false, "Timed out.");
            }

            var output = string.IsNullOrWhiteSpace(standardOutput)
                ? standardError.Trim()
                : standardOutput.Trim();

            if (!string.IsNullOrWhiteSpace(standardError) && process.ExitCode != 0)
            {
                output = string.IsNullOrWhiteSpace(output)
                    ? standardError.Trim()
                    : $"{output}{Environment.NewLine}{standardError.Trim()}";
            }

            return new CommandResult(process.ExitCode == 0, string.IsNullOrWhiteSpace(output) ? "-" : output);
        }
        catch (Exception exception)
        {
            return new CommandResult(false, exception.Message);
        }
    }

    private static string GetProductVersion(string executablePath)
    {
        try
        {
            var fileVersion = FileVersionInfo.GetVersionInfo(executablePath);
            return string.IsNullOrWhiteSpace(fileVersion.ProductVersion)
                ? executablePath
                : fileVersion.ProductVersion!;
        }
        catch
        {
            return executablePath;
        }
    }

    private static string? ResolveOllamaExecutablePath()
    {
        var commandPath = ResolveExecutableOnPath("ollama.exe");

        if (!string.IsNullOrWhiteSpace(commandPath))
        {
            return commandPath;
        }

        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "ollama.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ollama", "ollama.exe"));
    }

    private static string? ResolveOllamaAppExecutablePath()
    {
        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "Ollama.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ollama", "Ollama.exe"));
    }

    private static string? ResolveLmStudioExecutablePath()
    {
        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "LM Studio", "LM Studio.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LM Studio", "LM Studio.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LM Studio", "LM Studio.exe"));
    }

    private static string? ResolveOpenClawExecutablePath()
    {
        var commandPath = ResolveExecutableOnPath("openclaw.cmd", "openclaw");

        if (!string.IsNullOrWhiteSpace(commandPath))
        {
            return commandPath;
        }

        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "openclaw.cmd"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm", "openclaw.cmd"));
    }

    private JsonObject LoadOpenClawConfigRoot()
    {
        if (!File.Exists(OpenClawConfigFilePath))
        {
            return new JsonObject();
        }

        var parsed = JsonNode.Parse(File.ReadAllText(OpenClawConfigFilePath)) as JsonObject;
        return parsed ?? new JsonObject();
    }

    private void SaveOpenClawConfigRoot(JsonObject root)
    {
        var json = root.ToJsonString(
            new JsonSerializerOptions
            {
                WriteIndented = true
            });

        File.WriteAllText(OpenClawConfigFilePath, json, Encoding.UTF8);
    }

    private string BackupOpenClawConfig()
    {
        var backupDirectory = Path.Combine(OpenClawHomeFolder, "aihelper-backups");
        Directory.CreateDirectory(backupDirectory);

        var backupPath = Path.Combine(
            backupDirectory,
            $"openclaw-{DateTime.Now:yyyyMMdd-HHmmss}.json");

        File.Copy(OpenClawConfigFilePath, backupPath, overwrite: true);
        return backupPath;
    }

    private string ResolvePreferredOllamaModel(IReadOnlyList<string> preferredTags, string fallbackTag)
    {
        var installedModels = GetInstalledOllamaModels(ResolveOllamaExecutablePath());

        foreach (var preferredTag in preferredTags)
        {
            var exactMatch = installedModels.Keys.FirstOrDefault(model =>
                model.Equals(preferredTag, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(exactMatch))
            {
                return exactMatch;
            }

            var partialMatch = installedModels.Keys.FirstOrDefault(model =>
                model.Contains(preferredTag, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(partialMatch))
            {
                return partialMatch;
            }
        }

        return fallbackTag;
    }

    private static string BuildOllamaPrimaryModel(string modelTag)
    {
        return modelTag.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase)
            ? modelTag
            : $"ollama/{modelTag}";
    }

    private static JsonObject EnsureObject(JsonObject root, string propertyName)
    {
        if (root[propertyName] is JsonObject objectNode)
        {
            return objectNode;
        }

        objectNode = new JsonObject();
        root[propertyName] = objectNode;
        return objectNode;
    }

    private static OpenClawConfigDetails ReadOpenClawConfig(string path)
    {
        if (!File.Exists(path))
        {
            return new OpenClawConfigDetails
            {
                Exists = false,
                Path = path
            };
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(path));
            var root = document.RootElement;

            var primaryModel = GetJsonString(root, "agents", "defaults", "model", "primary");
            var toolProfile = GetJsonString(root, "tools", "profile");
            var webSearchProvider = GetJsonString(root, "tools", "web", "search", "provider");
            var alsoAllow = GetJsonArrayValues(root, "tools", "alsoAllow");
            var telegramAllowFrom = GetJsonArrayValues(root, "channels", "telegram", "allowFrom");
            var telegramDmPolicy = GetJsonString(root, "channels", "telegram", "dmPolicy");
            var telegramBotToken = GetJsonString(root, "channels", "telegram", "botToken");

            return new OpenClawConfigDetails
            {
                Exists = true,
                Path = path,
                PrimaryModel = primaryModel,
                ToolProfile = toolProfile,
                WebSearchEnabled =
                    alsoAllow.Contains("web_search", StringComparer.OrdinalIgnoreCase) ||
                    !string.IsNullOrWhiteSpace(webSearchProvider),
                TelegramConfigured =
                    telegramAllowFrom.Count > 0 ||
                    !string.IsNullOrWhiteSpace(telegramBotToken) ||
                    telegramDmPolicy.Equals("allowlist", StringComparison.OrdinalIgnoreCase)
            };
        }
        catch (Exception exception)
        {
            return new OpenClawConfigDetails
            {
                Exists = true,
                Path = path,
                ParseError = exception.Message
            };
        }
    }

    private static string GetJsonString(JsonElement root, params string[] path)
    {
        if (!TryGetJsonValue(root, out var value, path))
        {
            return string.Empty;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString()?.Trim() ?? string.Empty,
            JsonValueKind.Number => value.ToString(),
            JsonValueKind.True => bool.TrueString,
            JsonValueKind.False => bool.FalseString,
            _ => string.Empty
        };
    }

    private static IReadOnlyList<string> GetJsonArrayValues(JsonElement root, params string[] path)
    {
        if (!TryGetJsonValue(root, out var value, path) || value.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var result = new List<string>();

        foreach (var item in value.EnumerateArray())
        {
            switch (item.ValueKind)
            {
                case JsonValueKind.String:
                    var stringValue = item.GetString()?.Trim();

                    if (!string.IsNullOrWhiteSpace(stringValue))
                    {
                        result.Add(stringValue);
                    }

                    break;
                case JsonValueKind.Number:
                    result.Add(item.ToString());
                    break;
            }
        }

        return result;
    }

    private static bool TryGetJsonValue(JsonElement root, out JsonElement value, params string[] path)
    {
        value = root;

        foreach (var segment in path)
        {
            if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty(segment, out value))
            {
                value = default;
                return false;
            }
        }

        return true;
    }

    private static bool DetermineOpenClawNodeInstalled(CommandResult nodeStatus)
    {
        if (nodeStatus == CommandResult.Missing)
        {
            return false;
        }

        return !ContainsAny(
            nodeStatus.Output,
            "scheduled task (missing)",
            "service missing",
            "not installed",
            "cannot find the file specified",
            "timed out");
    }

    private static bool DetermineOpenClawBrowserReady(CommandResult browserStatus)
    {
        if (browserStatus == CommandResult.Missing)
        {
            return false;
        }

        return browserStatus.Success &&
               !ContainsAny(
                   browserStatus.Output,
                   "pairing required",
                   "not installed",
                   "timed out",
                   "missing",
                   "error");
    }

    private static string BuildOpenClawDetail(
        CommandResult openClawVersion,
        string openClawPath,
        OpenClawConfigDetails config)
    {
        var lines = new List<string>();

        if (openClawVersion.Success &&
            !string.IsNullOrWhiteSpace(openClawVersion.Output) &&
            !string.Equals(openClawVersion.Output, "-", StringComparison.Ordinal))
        {
            lines.Add(openClawVersion.Output);
        }

        lines.Add(openClawPath);
        lines.Add(config.Exists
            ? $"Config: {config.Path}"
            : $"Config not created yet: {config.Path}");

        if (!string.IsNullOrWhiteSpace(config.ParseError))
        {
            lines.Add($"Config parse warning: {config.ParseError}");
            return string.Join(Environment.NewLine, lines);
        }

        if (!string.IsNullOrWhiteSpace(config.PrimaryModel))
        {
            lines.Add($"Primary model: {config.PrimaryModel}");
        }

        if (!string.IsNullOrWhiteSpace(config.ToolProfile))
        {
            lines.Add($"Tools profile: {config.ToolProfile}");
        }

        lines.Add(config.WebSearchEnabled ? "Web search is configured." : "Web search is not configured.");
        lines.Add(config.TelegramConfigured ? "Telegram channel is configured." : "Telegram channel is not configured.");

        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildOpenClawNodeDetail(CommandResult nodeStatus)
    {
        if (nodeStatus == CommandResult.Missing)
        {
            return "OpenClaw is not installed yet.";
        }

        return string.IsNullOrWhiteSpace(nodeStatus.Output)
            ? "No node status output."
            : nodeStatus.Output;
    }

    private static string BuildOpenClawBrowserDetail(CommandResult browserStatus)
    {
        if (browserStatus == CommandResult.Missing)
        {
            return "OpenClaw browser tools are unavailable because OpenClaw is not installed yet.";
        }

        return string.IsNullOrWhiteSpace(browserStatus.Output)
            ? "No browser status output."
            : browserStatus.Output;
    }

    private static string BuildOllamaDetail(
        CommandResult ollamaVersion,
        string ollamaPath,
        bool commandVisible,
        bool appAvailable,
        bool serverRunning,
        bool trayRunning,
        int modelCount)
    {
        var lines = new List<string>();

        if (ollamaVersion.Success &&
            !string.IsNullOrWhiteSpace(ollamaVersion.Output) &&
            !string.Equals(ollamaVersion.Output, "-", StringComparison.Ordinal))
        {
            lines.Add(ollamaVersion.Output);
        }

        lines.Add(ollamaPath);
        lines.Add(commandVisible
            ? "The ollama command is visible in new terminals."
            : "Installed, but older terminals may need to be reopened before PATH updates are visible.");

        if (serverRunning)
        {
            lines.Add("Local server is listening on 127.0.0.1:11434.");
        }
        else if (trayRunning || appAvailable)
        {
            lines.Add("Ollama app is present, but the local server is not listening yet.");
        }
        else
        {
            lines.Add("Local server is not running yet.");
        }

        lines.Add(modelCount == 0
            ? "No local models downloaded yet."
            : $"{modelCount} local model(s) downloaded.");

        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildOllamaModelsSummary(IReadOnlyDictionary<string, string> installedModels)
    {
        if (installedModels.Count == 0)
        {
            return string.Empty;
        }

        var names = installedModels.Keys
            .Take(2)
            .ToList();

        return installedModels.Count <= 2
            ? string.Join(", ", names)
            : $"{string.Join(", ", names)} +{installedModels.Count - 2}";
    }

    private static bool IsLocalTcpEndpointReachable(int port)
    {
        try
        {
            using var client = new TcpClient();
            var connectTask = client.ConnectAsync("127.0.0.1", port);
            return connectTask.Wait(750) && client.Connected;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsAnyProcessRunning(params string[] nameFragments)
    {
        try
        {
            foreach (var process in Process.GetProcesses())
            {
                using (process)
                {
                    if (nameFragments.Any(fragment =>
                            !string.IsNullOrWhiteSpace(fragment) &&
                            process.ProcessName.Contains(fragment, StringComparison.OrdinalIgnoreCase)))
                    {
                        return true;
                    }
                }
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static bool ContainsAny(string text, params string[] fragments)
    {
        return fragments.Any(fragment =>
            !string.IsNullOrWhiteSpace(fragment) &&
            text.Contains(fragment, StringComparison.OrdinalIgnoreCase));
    }

    private static string EscapeForSingleQuotedPowerShell(string value)
    {
        return value.Replace("'", "''");
    }

    private static string? ResolveExecutableOnPath(params string[] commandNames)
    {
        foreach (var commandName in commandNames)
        {
            var whereResult = RunCommand("cmd.exe", $"/c where {commandName}");

            if (!whereResult.Success)
            {
                continue;
            }

            var discoveredPath = whereResult.Output
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault(File.Exists);

            if (!string.IsNullOrWhiteSpace(discoveredPath))
            {
                return discoveredPath;
            }
        }

        return null;
    }

    private sealed class OpenClawConfigDetails
    {
        public required bool Exists { get; init; }
        public required string Path { get; init; }
        public string PrimaryModel { get; init; } = string.Empty;
        public string ToolProfile { get; init; } = string.Empty;
        public bool WebSearchEnabled { get; init; }
        public bool TelegramConfigured { get; init; }
        public string ParseError { get; init; } = string.Empty;
    }

    private static string? FindExistingPath(params string[] paths)
    {
        foreach (var path in paths)
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    private static string QuoteForCommandLine(string value)
    {
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private readonly record struct CommandResult(bool Success, string Output)
    {
        public static CommandResult Missing => new(false, "Not found");
    }

    private readonly record struct WingetPackageInfo(bool IsInstalled, string Detail)
    {
        public static WingetPackageInfo Missing => new(false, "Not found");
    }
}


