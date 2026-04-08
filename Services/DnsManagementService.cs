using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using LaptopSessionViewer.Models;
using Microsoft.Win32;

namespace LaptopSessionViewer.Services;

public sealed class DnsManagementService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly AppLogService _logService = new();

    public string BackupFilePath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AIHelper",
            "dns-backups.json");

    public IReadOnlyList<DnsAdapterRecord> GetAdapters()
    {
        var backups = LoadBackups().Select(item => item.InterfaceIndex).ToHashSet();
        var adapters = new List<(DnsAdapterRecord Adapter, int Priority, long Speed)>();

        foreach (var networkInterface in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (!ShouldShowAdapter(networkInterface))
            {
                continue;
            }

            var ipv4Properties = TryGetIpv4Properties(networkInterface);

            if (ipv4Properties is null)
            {
                continue;
            }

            var currentDns = GetCurrentDnsServers(networkInterface);
            var isAutomatic = IsAutomaticDnsConfiguration(networkInterface.Id);

            adapters.Add(
                (
                    new DnsAdapterRecord
                    {
                        InterfaceIndex = ipv4Properties.Index,
                        InterfaceAlias = networkInterface.Name,
                        Description = networkInterface.Description ?? string.Empty,
                        Status = networkInterface.OperationalStatus.ToString(),
                        DnsServers = currentDns,
                        IsAutomatic = isAutomatic,
                        HasSavedBackup = backups.Contains(ipv4Properties.Index)
                    },
                    GetSelectionPriority(networkInterface),
                    networkInterface.Speed
                ));
        }

        _logService.Info(nameof(DnsManagementService), $"Loaded {adapters.Count} network adapters.");

        return adapters
            .OrderByDescending(item => item.Priority)
            .ThenByDescending(item => item.Speed)
            .ThenBy(item => item.Adapter.InterfaceAlias, StringComparer.OrdinalIgnoreCase)
            .Select(item => item.Adapter)
            .ToList();
    }

    public IReadOnlyList<string> GetCurrentDnsServers(NetworkInterface adapter)
    {
        try
        {
            return adapter.GetIPProperties()
                .DnsAddresses
                .Where(address => address.AddressFamily == AddressFamily.InterNetwork)
                .Select(address => address.ToString())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception exception)
        {
            _logService.Error(
                nameof(DnsManagementService),
                $"Failed to read current DNS servers for adapter '{adapter.Name}'.",
                exception);

            return [];
        }
    }

    public void ApplyPreset(DnsAdapterRecord adapter, DnsPreset preset)
    {
        if (preset.IsAutomatic)
        {
            ResetToAutomatic(adapter);
            return;
        }

        ApplyCustomDns(
            adapter,
            preset.PrimaryDns,
            preset.SecondaryDns,
            preset.EnableDoh,
            preset.DohTemplate);
    }

    public void SaveBackup(DnsAdapterRecord adapter)
    {
        var backups = LoadBackups().ToDictionary(item => item.InterfaceIndex);
        var dohState = GetCurrentDohState(adapter.DnsServers);

        backups[adapter.InterfaceIndex] = new DnsBackupRecord
        {
            InterfaceIndex = adapter.InterfaceIndex,
            InterfaceAlias = adapter.InterfaceAlias,
            DnsServers = adapter.DnsServers.ToList(),
            EnableDoh = dohState.EnableDoh,
            DohTemplate = dohState.DohTemplate,
            SavedAtUtc = DateTime.UtcNow
        };

        SaveBackups(backups.Values);
        _logService.Info(nameof(DnsManagementService), $"Saved DNS backup for '{adapter.InterfaceAlias}'.");
    }

    public bool HasBackup(int interfaceIndex)
    {
        return LoadBackups().Any(item => item.InterfaceIndex == interfaceIndex);
    }

    public void ApplyCustomDns(
        DnsAdapterRecord adapter,
        string primaryDns,
        string? secondaryDns,
        bool enableDoh,
        string? dohTemplate)
    {
        var primary = ValidateDnsAddress(primaryDns);
        var secondary = string.IsNullOrWhiteSpace(secondaryDns) ? null : ValidateDnsAddress(secondaryDns);
        var servers = new List<string> { primary };

        if (!string.IsNullOrWhiteSpace(secondary))
        {
            servers.Add(secondary);
        }

        var normalizedDohTemplate = enableDoh ? ValidateDohTemplate(dohTemplate) : string.Empty;
        SaveBackup(adapter);

        var script = BuildApplyScript(adapter.InterfaceAlias, servers, enableDoh, normalizedDohTemplate);
        RunElevatedPowerShellOrThrow(adapter.InterfaceAlias, script);
    }

    public void ResetToAutomatic(DnsAdapterRecord adapter)
    {
        SaveBackup(adapter);
        var script = BuildResetAutomaticScript(adapter.InterfaceAlias);
        RunElevatedPowerShellOrThrow(adapter.InterfaceAlias, script);
    }

    public void RestoreBackup(DnsAdapterRecord adapter)
    {
        var backup = LoadBackups().FirstOrDefault(item => item.InterfaceIndex == adapter.InterfaceIndex);

        if (backup is null)
        {
            throw new InvalidOperationException("DNS backup was not found for the selected adapter.");
        }

        string script;

        if (backup.DnsServers.Count == 0)
        {
            script = BuildResetAutomaticScript(adapter.InterfaceAlias);
        }
        else
        {
            script = BuildApplyScript(
                adapter.InterfaceAlias,
                backup.DnsServers,
                backup.EnableDoh,
                backup.DohTemplate);
        }

        RunElevatedPowerShellOrThrow(adapter.InterfaceAlias, script);
    }

    private IReadOnlyList<DnsBackupRecord> LoadBackups()
    {
        if (!File.Exists(BackupFilePath))
        {
            return [];
        }

        try
        {
            var json = File.ReadAllText(BackupFilePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<List<DnsBackupRecord>>(json) ?? [];
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(DnsManagementService), "Failed to load DNS backup file.", exception);
            return [];
        }
    }

    private void SaveBackups(IEnumerable<DnsBackupRecord> backups)
    {
        var directory = Path.GetDirectoryName(BackupFilePath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(
            BackupFilePath,
            JsonSerializer.Serialize(backups.OrderBy(item => item.InterfaceAlias), JsonOptions),
            Encoding.UTF8);
    }

    private DnsDohState GetCurrentDohState(IReadOnlyList<string> servers)
    {
        if (servers.Count == 0)
        {
            return DnsDohState.Disabled;
        }

        try
        {
            var quotedServers = string.Join(", ", servers.Select(server => $"'{EscapePowerShellString(server)}'"));
            var script = $"""
$records = Get-DnsClientDohServerAddress -ServerAddress @({quotedServers}) -ErrorAction SilentlyContinue |
    Select-Object ServerAddress, DohTemplate, AutoUpgrade, AllowFallbackToUdp
$records | ConvertTo-Json -Depth 4 -Compress
""";

            var result = RunPowerShell(script, requiresElevation: false);

            if (!result.Success)
            {
                return DnsDohState.Disabled;
            }

            var json = result.Output.Trim();

            if (string.IsNullOrWhiteSpace(json) || json == "null")
            {
                return DnsDohState.Disabled;
            }

            List<DnsDohShellRecord> records;

            using (var document = JsonDocument.Parse(json))
            {
                records = document.RootElement.ValueKind switch
                {
                    JsonValueKind.Array => JsonSerializer.Deserialize<List<DnsDohShellRecord>>(json) ?? [],
                    JsonValueKind.Object => JsonSerializer.Deserialize<DnsDohShellRecord>(json) is { } single
                        ? [single]
                        : [],
                    _ => []
                };
            }

            var activeTemplate = records
                .Where(record => record.AutoUpgrade && !string.IsNullOrWhiteSpace(record.DohTemplate))
                .Select(record => record.DohTemplate!)
                .FirstOrDefault();

            return string.IsNullOrWhiteSpace(activeTemplate)
                ? DnsDohState.Disabled
                : new DnsDohState(true, activeTemplate);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(DnsManagementService), "Failed to read current DoH configuration.", exception);
            return DnsDohState.Disabled;
        }
    }

    private bool IsAutomaticDnsConfiguration(string interfaceId)
    {
        try
        {
            using var key =
                Registry.LocalMachine.OpenSubKey(
                    $@"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{{{interfaceId}}}");

            var nameServer = key?.GetValue("NameServer")?.ToString();
            return string.IsNullOrWhiteSpace(nameServer);
        }
        catch (Exception exception)
        {
            _logService.Error(
                nameof(DnsManagementService),
                $"Failed to detect DNS mode for interface id '{interfaceId}'.",
                exception);
            return false;
        }
    }

    private static IPv4InterfaceProperties? TryGetIpv4Properties(NetworkInterface adapter)
    {
        try
        {
            return adapter.GetIPProperties().GetIPv4Properties();
        }
        catch
        {
            return null;
        }
    }

    private static bool ShouldShowAdapter(NetworkInterface adapter)
    {
        return adapter.NetworkInterfaceType is not NetworkInterfaceType.Loopback and not NetworkInterfaceType.Tunnel &&
               !string.IsNullOrWhiteSpace(adapter.Name);
    }

    private static int GetSelectionPriority(NetworkInterface adapter)
    {
        var priority = 0;

        if (adapter.OperationalStatus == OperationalStatus.Up)
        {
            priority += 1000;
        }

        if (HasIpv4DefaultGateway(adapter))
        {
            priority += 400;
        }

        if (HasIpv4Address(adapter))
        {
            priority += 200;
        }

        priority += adapter.NetworkInterfaceType switch
        {
            NetworkInterfaceType.Wireless80211 => 120,
            NetworkInterfaceType.Ethernet => 110,
            NetworkInterfaceType.GigabitEthernet => 110,
            NetworkInterfaceType.FastEthernetFx => 100,
            NetworkInterfaceType.FastEthernetT => 100,
            NetworkInterfaceType.Ppp => 90,
            _ => 0
        };

        if (IsLikelyVirtualAdapter(adapter))
        {
            priority -= 500;
        }

        return priority;
    }

    private static bool HasIpv4DefaultGateway(NetworkInterface adapter)
    {
        try
        {
            return adapter.GetIPProperties()
                .GatewayAddresses
                .Any(gateway => gateway.Address.AddressFamily == AddressFamily.InterNetwork &&
                                !IPAddress.Any.Equals(gateway.Address) &&
                                !IPAddress.None.Equals(gateway.Address));
        }
        catch
        {
            return false;
        }
    }

    private static bool HasIpv4Address(NetworkInterface adapter)
    {
        try
        {
            return adapter.GetIPProperties()
                .UnicastAddresses
                .Any(address => address.Address.AddressFamily == AddressFamily.InterNetwork);
        }
        catch
        {
            return false;
        }
    }

    private static bool IsLikelyVirtualAdapter(NetworkInterface adapter)
    {
        var text = $"{adapter.Name} {adapter.Description}".ToLowerInvariant();

        return text.Contains("virtual") ||
               text.Contains("hyper-v") ||
               text.Contains("vmware") ||
               text.Contains("vethernet") ||
               text.Contains("wintun") ||
               text.Contains("zerotier") ||
               text.Contains("tailscale");
    }

    private static string ValidateDnsAddress(string value)
    {
        var trimmed = value.Trim();

        if (!IPAddress.TryParse(trimmed, out _))
        {
            throw new InvalidOperationException($"Invalid DNS address: {value}");
        }

        return trimmed;
    }

    private static string ValidateDohTemplate(string? value)
    {
        var trimmed = (value ?? string.Empty).Trim();

        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) ||
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Invalid DoH template URL.");
        }

        return trimmed;
    }

    private static string BuildApplyScript(
        string adapterAlias,
        IReadOnlyList<string> servers,
        bool enableDoh,
        string? dohTemplate)
    {
        var script = new StringBuilder();
        script.AppendLine("$ErrorActionPreference = 'Stop'");
        script.AppendLine(
            $"& netsh interface ipv4 set dns name=\"{EscapePowerShellString(adapterAlias)}\" static {servers[0]} primary");
        script.AppendLine("if ($LASTEXITCODE -ne 0) { throw 'netsh failed to set primary DNS.' }");

        if (servers.Count > 1)
        {
            script.AppendLine(
                $"& netsh interface ipv4 add dns name=\"{EscapePowerShellString(adapterAlias)}\" {servers[1]} index=2");
            script.AppendLine("if ($LASTEXITCODE -ne 0) { throw 'netsh failed to add secondary DNS.' }");
        }

        script.AppendLine(BuildDohConfigurationScript(servers, enableDoh, dohTemplate));
        return script.ToString();
    }

    private static string BuildResetAutomaticScript(string adapterAlias)
    {
        return
            "$ErrorActionPreference = 'Stop'" + Environment.NewLine +
            $"& netsh interface ipv4 set dns name=\"{EscapePowerShellString(adapterAlias)}\" dhcp" + Environment.NewLine +
            "if ($LASTEXITCODE -ne 0) { throw 'netsh failed to reset DNS to DHCP.' }";
    }

    private static string BuildDohConfigurationScript(
        IReadOnlyList<string> servers,
        bool enableDoh,
        string? dohTemplate)
    {
        if (servers.Count == 0)
        {
            return string.Empty;
        }

        var quotedServers = string.Join(", ", servers.Select(server => $"'{EscapePowerShellString(server)}'"));
        var template = EscapePowerShellString(dohTemplate ?? string.Empty);
        var script = new StringBuilder();
        script.AppendLine($"$servers = @({quotedServers})");
        script.AppendLine("foreach ($server in $servers) {");
        script.AppendLine("    $existing = Get-DnsClientDohServerAddress -ServerAddress $server -ErrorAction SilentlyContinue");

        if (enableDoh)
        {
            script.AppendLine("    if ($existing) {");
            script.AppendLine(
                $"        Set-DnsClientDohServerAddress -ServerAddress $server -DohTemplate '{template}' -AutoUpgrade $true -AllowFallbackToUdp $true | Out-Null");
            script.AppendLine("    }");
            script.AppendLine("    else {");
            script.AppendLine(
                $"        Add-DnsClientDohServerAddress -ServerAddress $server -DohTemplate '{template}' -AutoUpgrade $true -AllowFallbackToUdp $true | Out-Null");
            script.AppendLine("    }");
        }
        else
        {
            script.AppendLine("    if ($existing) {");
            script.AppendLine(
                "        Set-DnsClientDohServerAddress -ServerAddress $server -AutoUpgrade $false -AllowFallbackToUdp $true | Out-Null");
            script.AppendLine("    }");
        }

        script.AppendLine("}");
        return script.ToString();
    }

    private void RunElevatedPowerShellOrThrow(string adapterAlias, string script)
    {
        var result = RunPowerShell(script, requiresElevation: true);

        if (!result.Success)
        {
            _logService.Error(nameof(DnsManagementService), $"Failed to apply DNS change to '{adapterAlias}': {result.Output}");
            throw new InvalidOperationException(result.Output);
        }
    }

    private static CommandResult RunPowerShell(string script, bool requiresElevation)
    {
        var tempScriptPath = Path.Combine(Path.GetTempPath(), $"aihelper-dns-{Guid.NewGuid():N}.ps1");

        try
        {
            File.WriteAllText(tempScriptPath, script, new UTF8Encoding(false));

            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempScriptPath}\"",
                UseShellExecute = requiresElevation,
                CreateNoWindow = !requiresElevation,
                WindowStyle = requiresElevation ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden,
                RedirectStandardOutput = !requiresElevation,
                RedirectStandardError = !requiresElevation,
                Verb = requiresElevation ? "runas" : string.Empty
            };

            process.Start();

            var output = string.Empty;
            var error = string.Empty;

            if (!requiresElevation)
            {
                output = process.StandardOutput.ReadToEnd();
                error = process.StandardError.ReadToEnd();
            }

            process.WaitForExit();

            var combined = string.IsNullOrWhiteSpace(output)
                ? error.Trim()
                : output.Trim();

            if (!string.IsNullOrWhiteSpace(error) && !string.IsNullOrWhiteSpace(output))
            {
                combined = $"{output.Trim()}{Environment.NewLine}{error.Trim()}";
            }

            if (process.ExitCode != 0)
            {
                return new CommandResult(false, string.IsNullOrWhiteSpace(combined) ? "Command failed." : combined);
            }

            return new CommandResult(true, string.IsNullOrWhiteSpace(combined) ? "OK" : combined);
        }
        catch (Win32Exception exception) when (exception.NativeErrorCode == 1223)
        {
            return new CommandResult(false, "The operation was cancelled.");
        }
        finally
        {
            try
            {
                if (File.Exists(tempScriptPath))
                {
                    File.Delete(tempScriptPath);
                }
            }
            catch
            {
            }
        }
    }

    private static string EscapePowerShellString(string value)
    {
        return value.Replace("'", "''");
    }

    private sealed class DnsDohShellRecord
    {
        public string? ServerAddress { get; init; }
        public string? DohTemplate { get; init; }
        public bool AutoUpgrade { get; init; }
        public bool AllowFallbackToUdp { get; init; }
    }

    private readonly record struct DnsDohState(bool EnableDoh, string DohTemplate)
    {
        public static DnsDohState Disabled => new(false, string.Empty);
    }

    private readonly record struct CommandResult(bool Success, string Output);
}
