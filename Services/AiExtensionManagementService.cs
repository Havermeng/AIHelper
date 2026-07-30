using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class AiExtensionManagementService
{
    private const int CommandTimeoutMilliseconds = 45000;
    private readonly CodexEnvironmentService _environmentService;
    private readonly AppLogService _logService;
    private readonly HttpClient _localHttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(3)
    };

    private static readonly IReadOnlyDictionary<string, ProvisioningSpec> Specs =
        new Dictionary<string, ProvisioningSpec>(StringComparer.OrdinalIgnoreCase)
        {
            ["preset-plugin-hugging-face"] = new(
                "plugin",
                "hugging-face@openai-curated",
                string.Empty,
                string.Empty,
                []),
            ["preset-mcp-github"] = new(
                "plugin",
                "github@openai-curated",
                string.Empty,
                string.Empty,
                []),
            ["preset-mcp-canva"] = new(
                "plugin",
                "canva@openai-curated",
                string.Empty,
                string.Empty,
                []),
            ["preset-mcp-playwright"] = new(
                "mcp",
                string.Empty,
                "aihelper-playwright",
                "0.0.78",
                ["-y", "@playwright/mcp@0.0.78"],
                InstallIntoClaude: true),
            ["preset-mcp-filesystem"] = new(
                "mcp",
                string.Empty,
                "aihelper-filesystem",
                "2026.7.10",
                ["-y", "@modelcontextprotocol/server-filesystem@2026.7.10", "{workspace}"],
                InstallIntoClaude: true),
            ["preset-skill-creator"] = new(
                "skill",
                "skill-creator",
                string.Empty,
                "anthropics/skills",
                ["https://github.com/anthropics/skills.git"]),
            ["preset-skill-pdf"] = new(
                "skill",
                "pdf",
                string.Empty,
                "anthropics/skills",
                ["https://github.com/anthropics/skills.git"]),
            ["preset-skill-docx"] = new(
                "skill",
                "docx",
                string.Empty,
                "anthropics/skills",
                ["https://github.com/anthropics/skills.git"]),
            ["preset-opencode-session-bridge"] = new(
                "builtin",
                string.Empty,
                string.Empty,
                "AIHelper",
                []),
            ["preset-lmstudio-openai-provider"] = new(
                "endpoint",
                string.Empty,
                string.Empty,
                "http://127.0.0.1:1234/v1",
                [])
        };

    public AiExtensionManagementService(
        CodexEnvironmentService environmentService,
        AppLogService logService)
    {
        _environmentService = environmentService;
        _logService = logService;
    }

    public async Task RefreshAsync(
        IEnumerable<AiExtensionItem> items,
        CancellationToken cancellationToken = default)
    {
        var managedItems = items
            .Where(item => Specs.ContainsKey(item.Id))
            .ToList();

        if (managedItems.Count == 0)
        {
            return;
        }

        var codexAvailable = File.Exists(_environmentService.CodexCommandPath);
        var claudeAvailable = File.Exists(_environmentService.ClaudeCommandPath);
        var needsClaudeMcpList = managedItems.Any(item => Specs[item.Id].InstallIntoClaude);
        var pluginListTask = codexAvailable
            ? RunCodexAsync(["plugin", "list"], cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Codex CLI was not found."));
        var mcpListTask = codexAvailable
            ? RunCodexAsync(["mcp", "list"], cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Codex CLI was not found."));
        var claudeMcpListTask = claudeAvailable && needsClaudeMcpList
            ? RunClaudeAsync(["mcp", "list"], cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Claude Code CLI was not found."));
        var lmStudioTask = ProbeLmStudioEndpointAsync(cancellationToken);

        await Task.WhenAll(pluginListTask, mcpListTask, claudeMcpListTask, lmStudioTask);
        var pluginList = await pluginListTask;
        var mcpList = await mcpListTask;
        var claudeMcpList = await claudeMcpListTask;
        var lmStudioReady = await lmStudioTask;
        _logService.Info(
            nameof(AiExtensionManagementService),
            $"Managed extension verification completed. Codex={_environmentService.CodexCommandPath}; PluginsOk={pluginList.Success}; PluginsLength={pluginList.Output.Length}; McpOk={mcpList.Success}; McpLength={mcpList.Output.Length}; Claude={_environmentService.ClaudeCommandPath}; ClaudeMcpOk={claudeMcpList.Success}; LmStudio={lmStudioReady}.");

        foreach (var item in managedItems)
        {
            var spec = Specs[item.Id];
            item.ManagementKind = spec.Kind;
            item.PackageVersion = spec.Version;
            item.CanProvision = spec.Kind is "plugin" or "mcp" or "endpoint" or "skill";
            item.CanUninstall = spec.Kind is "plugin" or "mcp" or "skill";
            item.HasVerificationError = false;

            switch (spec.Kind)
            {
                case "plugin":
                ApplyPluginStatus(item, spec, pluginList, codexAvailable);
                break;
                case "mcp":
                    ApplyMcpStatus(item, spec, mcpList, codexAvailable, claudeMcpList, claudeAvailable);
                    break;
                case "skill":
                    ApplySkillStatus(item, spec);
                    break;
                case "builtin":
                    ApplyBuiltInStatus(item);
                    break;
                case "endpoint":
                    ApplyEndpointStatus(item, lmStudioReady);
                    break;
            }
        }
    }

    public async Task<AiExtensionOperationResult> InstallAsync(
        AiExtensionItem item,
        CancellationToken cancellationToken = default)
    {
        if (!Specs.TryGetValue(item.Id, out var spec))
        {
            return AiExtensionOperationResult.Fail("This entry has no trusted automatic installer.");
        }

        item.IsBusy = true;
        item.HasVerificationError = false;

        try
        {
            CommandExecutionResult operation;

            switch (spec.Kind)
            {
                case "plugin":
                    operation = await RunCodexAsync(
                        ["plugin", "add", spec.Selector, "--json"],
                        cancellationToken);
                    break;
                case "mcp":
                    var preflight = await VerifyNpmPackageAsync(spec, cancellationToken);
                    if (!preflight.Success)
                    {
                        return FailAndMark(item, preflight.Output);
                    }

                    var mcpCommandTail = spec.CommandArguments
                        .Skip(1)
                        .Select(value => string.Equals(value, "{workspace}", StringComparison.Ordinal)
                            ? EnsureWorkspaceDirectory()
                            : value)
                        .ToList();
                    var arguments = new List<string> { "mcp", "add", spec.McpName, "--", "npx", "-y" };
                    arguments.AddRange(mcpCommandTail);
                    operation = await RunCodexAsync(arguments, cancellationToken);

                    if (operation.Success &&
                        spec.InstallIntoClaude &&
                        File.Exists(_environmentService.ClaudeCommandPath))
                    {
                        var claudeArguments = new List<string>
                        {
                            "mcp", "add", "-s", "user", spec.McpName, "--", "npx", "-y"
                        };
                        claudeArguments.AddRange(mcpCommandTail);
                        var claudeOperation = await RunClaudeAsync(claudeArguments, cancellationToken);

                        if (!claudeOperation.Success &&
                            !claudeOperation.Output.Contains("already exists", StringComparison.OrdinalIgnoreCase))
                        {
                            return FailAndMark(
                                item,
                                $"Installed for Codex, but Claude Code setup failed: {claudeOperation.Output}");
                        }
                    }

                    break;
                case "skill":
                    operation = await InstallSharedSkillAsync(spec, cancellationToken);
                    break;
                case "endpoint":
                    var endpointReady = await ProbeLmStudioEndpointAsync(cancellationToken);
                    if (!endpointReady)
                    {
                        return FailAndMark(
                            item,
                            "LM Studio is installed or configured, but its local server did not answer on 127.0.0.1:1234.");
                    }

                    item.IsInstalled = true;
                    item.IsEnabled = true;
                    item.IsVerified = true;
                    item.VerificationDetail = "The local /v1/models endpoint answered successfully.";
                    return AiExtensionOperationResult.Ok(item.VerificationDetail);
                default:
                    return FailAndMark(item, "This entry is built in and does not need installation.");
            }

            if (!operation.Success)
            {
                return FailAndMark(item, operation.Output);
            }

            await RefreshAsync([item], cancellationToken);
            if (!item.IsVerified)
            {
                return FailAndMark(
                    item,
                    string.IsNullOrWhiteSpace(item.VerificationDetail)
                        ? "The command finished, but AIHelper could not verify the resulting configuration."
                        : item.VerificationDetail);
            }

            return AiExtensionOperationResult.Ok(item.VerificationDetail);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(AiExtensionManagementService), $"Failed to install {item.Id}.", exception);
            return FailAndMark(item, exception.Message);
        }
        finally
        {
            item.IsBusy = false;
        }
    }

    public async Task<AiExtensionOperationResult> RemoveAsync(
        AiExtensionItem item,
        CancellationToken cancellationToken = default)
    {
        if (!Specs.TryGetValue(item.Id, out var spec) || spec.Kind is not ("plugin" or "mcp" or "skill"))
        {
            return AiExtensionOperationResult.Fail("This entry has no trusted automatic removal path.");
        }

        item.IsBusy = true;
        item.HasVerificationError = false;

        try
        {
            var result = spec.Kind switch
            {
                "plugin" => await RunCodexAsync(["plugin", "remove", spec.Selector, "--json"], cancellationToken),
                "skill" => RemoveSharedSkill(spec),
                _ => await RunCodexAsync(["mcp", "remove", spec.McpName], cancellationToken)
            };

            if (!result.Success)
            {
                return FailAndMark(item, result.Output);
            }

            if (spec.Kind == "mcp" &&
                spec.InstallIntoClaude &&
                File.Exists(_environmentService.ClaudeCommandPath))
            {
                var claudeResult = await RunClaudeAsync(
                    ["mcp", "remove", "-s", "user", spec.McpName],
                    cancellationToken);

                if (!claudeResult.Success &&
                    !claudeResult.Output.Contains("not found", StringComparison.OrdinalIgnoreCase) &&
                    !claudeResult.Output.Contains("No MCP", StringComparison.OrdinalIgnoreCase))
                {
                    return FailAndMark(
                        item,
                        $"Removed from Codex, but Claude Code removal failed: {claudeResult.Output}");
                }
            }

            await RefreshAsync([item], cancellationToken);
            if (item.IsInstalled)
            {
                return FailAndMark(item, "Removal finished, but the integration is still present in Codex.");
            }

            item.IsVerified = false;
            item.IsEnabled = false;
            item.VerificationDetail = "Removed and no longer reported by Codex.";
            return AiExtensionOperationResult.Ok(item.VerificationDetail);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(AiExtensionManagementService), $"Failed to remove {item.Id}.", exception);
            return FailAndMark(item, exception.Message);
        }
        finally
        {
            item.IsBusy = false;
        }
    }

    private static void ApplyPluginStatus(
        AiExtensionItem item,
        ProvisioningSpec spec,
        CommandExecutionResult pluginList,
        bool codexAvailable)
    {
        var selectorPattern = Regex.Escape(spec.Selector);
        var installed = pluginList.Success &&
                        Regex.IsMatch(
                            pluginList.Output,
                            $@"(?im)^[^\r\n]*{selectorPattern}\s+installed(?:,|\s)",
                            RegexOptions.CultureInvariant);
        var enabled = installed &&
                      Regex.IsMatch(
                          pluginList.Output,
                          $@"(?im)^[^\r\n]*{selectorPattern}[^\r\n]*\benabled\b",
                          RegexOptions.CultureInvariant) &&
                      !Regex.IsMatch(
                          pluginList.Output,
                          $@"(?im)^[^\r\n]*{selectorPattern}[^\r\n]*\bdisabled\b",
                          RegexOptions.CultureInvariant);
        item.IsInstalled = installed;
        item.IsEnabled = enabled;
        item.IsVerified = installed && enabled;
        item.HasVerificationError = codexAvailable && !pluginList.Success;
        item.VerificationDetail = !codexAvailable
            ? "Install or repair Codex CLI first."
            : item.IsVerified
                ? "Codex reports this marketplace plugin as installed and enabled."
                : pluginList.Success
                    ? "Codex does not report this marketplace plugin as installed."
                    : pluginList.Output;
    }

    private static void ApplyMcpStatus(
        AiExtensionItem item,
        ProvisioningSpec spec,
        CommandExecutionResult mcpList,
        bool codexAvailable,
        CommandExecutionResult claudeMcpList,
        bool claudeAvailable)
    {
        var matchingLine = mcpList.Output
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(line => line.StartsWith(spec.McpName, StringComparison.OrdinalIgnoreCase));
        var installed = !string.IsNullOrWhiteSpace(matchingLine);
        var enabled = installed &&
                      matchingLine!.Contains("enabled", StringComparison.OrdinalIgnoreCase);
        var expectedPackage = spec.CommandArguments
            .FirstOrDefault(argument => argument.StartsWith("@", StringComparison.Ordinal));
        var pinnedPackageMatches = string.IsNullOrWhiteSpace(expectedPackage) ||
                                   matchingLine?.Contains(expectedPackage, StringComparison.OrdinalIgnoreCase) == true;
        var codexVerified = installed && enabled && pinnedPackageMatches;

        // `claude mcp list` exits non-zero when any configured server is unhealthy or needs
        // authentication, so the exit code cannot be used to decide whether the list is usable.
        var claudeListUsable = claudeMcpList.Success ||
                               claudeMcpList.Output.Contains(':', StringComparison.Ordinal);
        var claudeRelevant = spec.InstallIntoClaude && claudeAvailable && claudeListUsable;
        var claudeInstalled = claudeRelevant &&
                              claudeMcpList.Output
                                  .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                  .Any(line => line.StartsWith($"{spec.McpName}:", StringComparison.OrdinalIgnoreCase));

        item.IsInstalled = installed;
        item.IsEnabled = enabled;
        item.IsVerified = codexVerified && (!claudeRelevant || claudeInstalled);
        item.HasVerificationError = codexAvailable && !mcpList.Success;
        item.VerificationDetail = !codexAvailable
            ? "Install or repair Codex CLI first."
            : item.IsVerified
                ? claudeRelevant
                    ? "Both Codex and Claude Code report this MCP server as configured."
                    : spec.InstallIntoClaude
                        ? "Codex reports the pinned MCP command as configured and enabled. Claude Code state could not be checked."
                        : "Codex reports the pinned MCP command as configured and enabled."
                : codexVerified && claudeRelevant && !claudeInstalled
                    ? "Codex is configured, but Claude Code does not report this MCP server yet. Reinstall to add it to Claude Code."
                    : installed && !pinnedPackageMatches
                        ? "An MCP entry with this name exists, but it does not match AIHelper's pinned package."
                        : mcpList.Success
                            ? "Codex does not report this MCP server as configured."
                            : mcpList.Output;
    }

    private void ApplySkillStatus(AiExtensionItem item, ProvisioningSpec spec)
    {
        var targetPath = Path.Combine(_environmentService.SharedSkillsFolder, spec.Selector);
        var installed = File.Exists(Path.Combine(targetPath, "SKILL.md"));

        item.IsInstalled = installed;
        item.IsEnabled = installed;
        item.IsVerified = installed;
        item.VerificationDetail = installed
            ? $"SKILL.md is present in the shared skills folder ({targetPath}). Both Codex and Claude Code can use it."
            : "The skill is not in the shared skills folder yet.";
    }

    private void ApplyBuiltInStatus(AiExtensionItem item)
    {
        var openCodePath = _environmentService.OpenCodeCommandPath;
        item.IsInstalled = true;
        item.IsEnabled = File.Exists(openCodePath);
        item.IsVerified = item.IsEnabled;
        item.VerificationDetail = item.IsVerified
            ? "The AIHelper bridge is built in and OpenCode was found."
            : "The bridge is built in, but OpenCode is not installed yet.";
    }

    private static void ApplyEndpointStatus(AiExtensionItem item, bool endpointReady)
    {
        item.IsInstalled = endpointReady;
        item.IsEnabled = endpointReady;
        item.IsVerified = endpointReady;
        item.VerificationDetail = endpointReady
            ? "The local LM Studio /v1/models endpoint answered successfully."
            : "Start the Local Server in LM Studio, then check the connection again.";
    }

    private async Task<CommandExecutionResult> VerifyNpmPackageAsync(
        ProvisioningSpec spec,
        CancellationToken cancellationToken)
    {
        var package = spec.CommandArguments
            .FirstOrDefault(argument => argument.StartsWith("@", StringComparison.Ordinal));

        if (string.IsNullOrWhiteSpace(package))
        {
            return CommandExecutionResult.Missing("The pinned npm package is missing from the installer specification.");
        }

        var npmPath = ResolveExecutable("npm.cmd", "npm.exe", "npm");
        if (string.IsNullOrWhiteSpace(npmPath))
        {
            return CommandExecutionResult.Missing("Node.js/npm is required before this MCP server can be installed.");
        }

        var result = await RunProcessAsync(
            npmPath,
            ["view", package, "version", "--json"],
            cancellationToken);

        if (!result.Success)
        {
            return result;
        }

        return result.Output.Contains(spec.Version, StringComparison.OrdinalIgnoreCase)
            ? result
            : CommandExecutionResult.Missing(
                $"The npm registry did not confirm pinned package version {spec.Version}. Nothing was changed.");
    }

    private async Task<bool> ProbeLmStudioEndpointAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var response = await _localHttpClient.GetAsync(
                "http://127.0.0.1:1234/v1/models",
                cancellationToken);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private Task<CommandExecutionResult> RunCodexAsync(
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        var codexPath = _environmentService.CodexCommandPath;
        return File.Exists(codexPath)
            ? RunProcessAsync(codexPath, arguments, cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Codex CLI was not found."));
    }

    private Task<CommandExecutionResult> RunClaudeAsync(
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        var claudePath = _environmentService.ClaudeCommandPath;
        return File.Exists(claudePath)
            ? RunProcessAsync(claudePath, arguments, cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Claude Code CLI was not found."));
    }

    private async Task<CommandExecutionResult> InstallSharedSkillAsync(
        ProvisioningSpec spec,
        CancellationToken cancellationToken)
    {
        var gitPath = ResolveExecutable("git.exe", "git");
        if (string.IsNullOrWhiteSpace(gitPath))
        {
            return CommandExecutionResult.Missing("Git is required to install skills. Install Git first, then try again.");
        }

        var repositoryUrl = spec.CommandArguments.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(repositoryUrl) ||
            !repositoryUrl.StartsWith("https://github.com/", StringComparison.OrdinalIgnoreCase))
        {
            return CommandExecutionResult.Missing("The skill source repository is not configured or is not a trusted GitHub URL.");
        }

        var temporaryRoot = Path.Combine(Path.GetTempPath(), $"aihelper-skill-{Guid.NewGuid():N}");

        try
        {
            var clone = await RunProcessAsync(
                gitPath,
                ["clone", "--depth", "1", repositoryUrl, temporaryRoot],
                cancellationToken);

            if (!clone.Success)
            {
                return clone;
            }

            var skillDirectory = Directory
                .EnumerateDirectories(temporaryRoot, spec.Selector, SearchOption.AllDirectories)
                .FirstOrDefault(directory => File.Exists(Path.Combine(directory, "SKILL.md")));

            if (skillDirectory is null)
            {
                return CommandExecutionResult.Missing(
                    $"The repository does not contain a '{spec.Selector}' folder with SKILL.md. Nothing was changed.");
            }

            var targetPath = Path.Combine(_environmentService.SharedSkillsFolder, spec.Selector);

            if (Directory.Exists(targetPath))
            {
                var removal = RemoveSharedSkill(spec);
                if (!removal.Success)
                {
                    return removal;
                }
            }

            CopyDirectory(skillDirectory, targetPath);
            return new CommandExecutionResult(true, $"The skill was installed to {targetPath}.");
        }
        finally
        {
            TryDeleteDirectory(temporaryRoot);
        }
    }

    private CommandExecutionResult RemoveSharedSkill(ProvisioningSpec spec)
    {
        var skillsRoot = Path.GetFullPath(_environmentService.SharedSkillsFolder);
        var targetPath = Path.GetFullPath(Path.Combine(skillsRoot, spec.Selector));

        if (!targetPath.StartsWith(skillsRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        {
            return CommandExecutionResult.Missing("The skill path is outside the shared skills folder. Nothing was changed.");
        }

        if (!Directory.Exists(targetPath))
        {
            return new CommandExecutionResult(true, "The skill folder is already absent.");
        }

        var attributes = File.GetAttributes(targetPath);
        if (attributes.HasFlag(FileAttributes.ReparsePoint))
        {
            return CommandExecutionResult.Missing("The skill folder is a link, not a regular folder. Remove it manually.");
        }

        try
        {
            DeleteDirectoryRobust(targetPath);
            return new CommandExecutionResult(true, "The skill folder was removed from the shared skills folder.");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return CommandExecutionResult.Missing($"Could not remove the skill folder: {exception.Message}");
        }
    }

    private static void CopyDirectory(string sourcePath, string targetPath)
    {
        Directory.CreateDirectory(targetPath);

        foreach (var file in Directory.EnumerateFiles(sourcePath, "*", SearchOption.AllDirectories))
        {
            var relativePath = Path.GetRelativePath(sourcePath, file);
            var destination = Path.Combine(targetPath, relativePath);
            var destinationDirectory = Path.GetDirectoryName(destination);

            if (!string.IsNullOrWhiteSpace(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            File.Copy(file, destination, overwrite: true);
        }
    }

    private static void DeleteDirectoryRobust(string path)
    {
        foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
        {
            File.SetAttributes(file, FileAttributes.Normal);
        }

        Directory.Delete(path, recursive: true);
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                DeleteDirectoryRobust(path);
            }
        }
        catch
        {
            // Leftover temporary clone directories are cleaned up by Windows temp maintenance.
        }
    }

    private static async Task<CommandExecutionResult> RunProcessAsync(
        string fileName,
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        try
        {
            var startInfo = CreateStartInfo(fileName, arguments);
            using var process = new Process { StartInfo = startInfo };
            process.Start();
            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(CommandTimeoutMilliseconds);

            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                TryKill(process);
                return CommandExecutionResult.Missing("The command timed out. No successful state was recorded.");
            }

            var stdout = await stdoutTask;
            var stderr = await stderrTask;
            var output = string.IsNullOrWhiteSpace(stdout) ? stderr.Trim() : stdout.Trim();

            if (process.ExitCode != 0 && !string.IsNullOrWhiteSpace(stderr) &&
                !output.Contains(stderr.Trim(), StringComparison.Ordinal))
            {
                output = $"{output}{Environment.NewLine}{stderr.Trim()}".Trim();
            }

            return new CommandExecutionResult(
                process.ExitCode == 0,
                string.IsNullOrWhiteSpace(output) ? "Command completed." : output);
        }
        catch (Exception exception)
        {
            return CommandExecutionResult.Missing(exception.Message);
        }
    }

    private static ProcessStartInfo CreateStartInfo(
        string fileName,
        IReadOnlyList<string> arguments)
    {
        var isCommandScript = fileName.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
                              fileName.EndsWith(".bat", StringComparison.OrdinalIgnoreCase);
        var startInfo = new ProcessStartInfo
        {
            FileName = isCommandScript ? "cmd.exe" : fileName,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        if (isCommandScript)
        {
            startInfo.ArgumentList.Add("/d");
            startInfo.ArgumentList.Add("/s");
            startInfo.ArgumentList.Add("/c");
            startInfo.ArgumentList.Add(BuildCmdCommand(fileName, arguments));
        }
        else
        {
            foreach (var argument in arguments)
            {
                startInfo.ArgumentList.Add(argument);
            }
        }

        return startInfo;
    }

    private static string BuildCmdCommand(string fileName, IEnumerable<string> arguments)
    {
        return string.Join(
            " ",
            new[] { QuoteCmdArgument(fileName) }.Concat(arguments.Select(QuoteCmdArgument)));
    }

    private static string QuoteCmdArgument(string value)
    {
        var safe = value
            .Replace("^", "^^", StringComparison.Ordinal)
            .Replace("&", "^&", StringComparison.Ordinal)
            .Replace("|", "^|", StringComparison.Ordinal)
            .Replace("<", "^<", StringComparison.Ordinal)
            .Replace(">", "^>", StringComparison.Ordinal)
            .Replace("%", "%%", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal);
        return $"\"{safe}\"";
    }

    private static string? ResolveExecutable(params string[] names)
    {
        foreach (var name in names)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "where.exe",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                startInfo.ArgumentList.Add(name);
                using var process = Process.Start(startInfo);

                if (process is null)
                {
                    continue;
                }

                var output = process.StandardOutput.ReadToEnd();
                process.WaitForExit(3000);
                var path = output
                    .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .FirstOrDefault(File.Exists);

                if (!string.IsNullOrWhiteSpace(path))
                {
                    return path;
                }
            }
            catch
            {
                // Try the next known executable name.
            }
        }

        return null;
    }

    private static string EnsureWorkspaceDirectory()
    {
        var path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "AIHelper Workspaces");
        Directory.CreateDirectory(path);
        return path;
    }

    private static AiExtensionOperationResult FailAndMark(AiExtensionItem item, string detail)
    {
        item.IsVerified = false;
        item.HasVerificationError = true;
        item.VerificationDetail = detail;
        return AiExtensionOperationResult.Fail(detail);
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch
        {
            // The process has already exited.
        }
    }

    private sealed record ProvisioningSpec(
        string Kind,
        string Selector,
        string McpName,
        string Version,
        IReadOnlyList<string> CommandArguments,
        bool InstallIntoClaude = false);

    private readonly record struct CommandExecutionResult(bool Success, string Output)
    {
        public static CommandExecutionResult Missing(string output) => new(false, output);
    }
}

public sealed record AiExtensionOperationResult(bool Success, string Detail)
{
    public static AiExtensionOperationResult Ok(string detail) => new(true, detail);

    public static AiExtensionOperationResult Fail(string detail) => new(false, detail);
}
