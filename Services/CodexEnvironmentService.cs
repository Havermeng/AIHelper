using System.Diagnostics;
using System.IO;
using System.Text;
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
        var ollamaPath = ResolveOllamaExecutablePath();
        var lmStudioPath = ResolveLmStudioExecutablePath();
        var openClawPath = ResolveOpenClawExecutablePath();
        var comfyUiPackage = GetInstalledWingetPackageInfo(ComfyUiDesktopWingetId);
        var pinokioPackage = GetInstalledWingetPackageInfo(PinokioWingetId);
        var ollamaVersion = !string.IsNullOrWhiteSpace(ollamaPath)
            ? RunCommand("cmd.exe", $"/c \"\"{ollamaPath}\" --version\"")
            : CommandResult.Missing;
        var openClawVersion = !string.IsNullOrWhiteSpace(openClawPath)
            ? RunCommand("cmd.exe", $"/c \"\"{openClawPath}\" --version\"")
            : CommandResult.Missing;
        var installedOllamaModels = GetInstalledOllamaModels(ollamaPath);

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
            OllamaDetail = !string.IsNullOrWhiteSpace(ollamaPath)
                ? $"{ollamaVersion.Output}{Environment.NewLine}{ollamaPath}"
                : "Ollama is not installed or not on PATH.",
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
                ? $"{openClawVersion.Output}{Environment.NewLine}{openClawPath}"
                : "OpenClaw is not installed or not on PATH.",
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

    public bool IsOllamaInstalled()
    {
        return !string.IsNullOrWhiteSpace(ResolveOllamaExecutablePath());
    }

    public void LaunchOllamaInstallTerminal()
    {
        LaunchWingetInstallTerminal(
            "ollama",
            OllamaWingetId,
            "Ollama");
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
        var scriptPath = Path.Combine(Path.GetTempPath(), "aihelper-install-openclaw.ps1");
        File.WriteAllText(scriptPath, BuildOpenClawInstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
    }

    public void LaunchOpenClawUninstallTerminal()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "aihelper-uninstall-openclaw.ps1");
        File.WriteAllText(scriptPath, BuildOpenClawUninstallScript(), Encoding.UTF8);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -ExecutionPolicy Bypass -File {QuoteForCommandLine(scriptPath)}",
                UseShellExecute = true,
                Verb = "runas"
            });
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

    private static string BuildOpenClawInstallScript()
    {
        return """
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '== Installing or updating OpenClaw ==' -ForegroundColor Cyan

& ([scriptblock]::Create((iwr -useb https://openclaw.ai/install.ps1))) -NoOnboard

Write-Host ''
Write-Host 'OpenClaw install/update command finished.' -ForegroundColor Green
Write-Host 'Return to AIHelper and refresh the environment status.' -ForegroundColor Green
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

    private static CommandResult RunCommand(string fileName, string arguments)
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
            process.WaitForExit(15000);

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
        var whereResult = RunCommand("cmd.exe", "/c where ollama.exe");

        if (whereResult.Success)
        {
            var discoveredPath = whereResult.Output
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault(File.Exists);

            if (!string.IsNullOrWhiteSpace(discoveredPath))
            {
                return discoveredPath;
            }
        }

        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "ollama.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ollama", "ollama.exe"));
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
        var whereCmdResult = RunCommand("cmd.exe", "/c where openclaw.cmd");

        if (whereCmdResult.Success)
        {
            var discoveredCmdPath = whereCmdResult.Output
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault(File.Exists);

            if (!string.IsNullOrWhiteSpace(discoveredCmdPath))
            {
                return discoveredCmdPath;
            }
        }

        var whereExeResult = RunCommand("cmd.exe", "/c where openclaw");

        if (whereExeResult.Success)
        {
            var discoveredExePath = whereExeResult.Output
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault(File.Exists);

            if (!string.IsNullOrWhiteSpace(discoveredExePath))
            {
                return discoveredExePath;
            }
        }

        return FindExistingPath(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "openclaw.cmd"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm", "openclaw.cmd"));
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
