using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class CodexEnvironmentService
{
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
            SessionsFolderPath = SessionsFolder
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

    private static string QuoteForCommandLine(string value)
    {
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private readonly record struct CommandResult(bool Success, string Output)
    {
        public static CommandResult Missing => new(false, "Not found");
    }
}
