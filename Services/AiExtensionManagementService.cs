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
                ["-y", "@playwright/mcp@0.0.78"]),
            ["preset-mcp-filesystem"] = new(
                "mcp",
                string.Empty,
                "aihelper-filesystem",
                "2026.7.10",
                ["-y", "@modelcontextprotocol/server-filesystem@2026.7.10", "{workspace}"]),
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
        var pluginListTask = codexAvailable
            ? RunCodexAsync(["plugin", "list"], cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Codex CLI was not found."));
        var mcpListTask = codexAvailable
            ? RunCodexAsync(["mcp", "list"], cancellationToken)
            : Task.FromResult(CommandExecutionResult.Missing("Codex CLI was not found."));
        var lmStudioTask = ProbeLmStudioEndpointAsync(cancellationToken);

        await Task.WhenAll(pluginListTask, mcpListTask, lmStudioTask);
        var pluginList = await pluginListTask;
        var mcpList = await mcpListTask;
        var lmStudioReady = await lmStudioTask;
        _logService.Info(
            nameof(AiExtensionManagementService),
            $"Managed extension verification completed. Codex={_environmentService.CodexCommandPath}; PluginsOk={pluginList.Success}; PluginsLength={pluginList.Output.Length}; McpOk={mcpList.Success}; McpLength={mcpList.Output.Length}; LmStudio={lmStudioReady}.");

        foreach (var item in managedItems)
        {
            var spec = Specs[item.Id];
            item.ManagementKind = spec.Kind;
            item.PackageVersion = spec.Version;
            item.CanProvision = spec.Kind is "plugin" or "mcp" or "endpoint";
            item.CanUninstall = spec.Kind is "plugin" or "mcp";
            item.HasVerificationError = false;

            switch (spec.Kind)
            {
                case "plugin":
                ApplyPluginStatus(item, spec, pluginList, codexAvailable);
                break;
                case "mcp":
                    ApplyMcpStatus(item, spec, mcpList, codexAvailable);
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

                    var arguments = new List<string> { "mcp", "add", spec.McpName, "--", "npx", "-y" };
                    arguments.AddRange(
                        spec.CommandArguments
                            .Skip(1)
                            .Select(value => string.Equals(value, "{workspace}", StringComparison.Ordinal)
                                ? EnsureWorkspaceDirectory()
                                : value));
                    operation = await RunCodexAsync(arguments, cancellationToken);
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
        if (!Specs.TryGetValue(item.Id, out var spec) || spec.Kind is not ("plugin" or "mcp"))
        {
            return AiExtensionOperationResult.Fail("This entry has no trusted automatic removal path.");
        }

        item.IsBusy = true;
        item.HasVerificationError = false;

        try
        {
            var result = spec.Kind == "plugin"
                ? await RunCodexAsync(["plugin", "remove", spec.Selector, "--json"], cancellationToken)
                : await RunCodexAsync(["mcp", "remove", spec.McpName], cancellationToken);

            if (!result.Success)
            {
                return FailAndMark(item, result.Output);
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
        bool codexAvailable)
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

        item.IsInstalled = installed;
        item.IsEnabled = enabled;
        item.IsVerified = installed && enabled && pinnedPackageMatches;
        item.HasVerificationError = codexAvailable && !mcpList.Success;
        item.VerificationDetail = !codexAvailable
            ? "Install or repair Codex CLI first."
            : item.IsVerified
                ? "Codex reports the pinned MCP command as configured and enabled."
                : installed && !pinnedPackageMatches
                    ? "An MCP entry with this name exists, but it does not match AIHelper's pinned package."
                    : mcpList.Success
                        ? "Codex does not report this MCP server as configured."
                        : mcpList.Output;
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
        IReadOnlyList<string> CommandArguments);

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
