param(
    [string]$InstallDir = "",
    [switch]$SkipShortcuts,
    [switch]$NoSelfElevate,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-Argument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

function New-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$WorkingDirectory,
        [string]$Description
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = $TargetPath
    $shortcut.Description = $Description
    $shortcut.Save()
}

if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    $InstallDir = Join-Path $env:ProgramFiles "AIHelper"
}

if (-not (Test-IsAdministrator)) {
    if ($NoSelfElevate) {
        throw "Administrator rights are required."
    }

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", (Quote-Argument $PSCommandPath),
        "-InstallDir", (Quote-Argument $InstallDir),
        "-NoSelfElevate"
    )

    if ($SkipShortcuts) {
        $arguments += "-SkipShortcuts"
    }

    if ($Quiet) {
        $arguments += "-Quiet"
    }

    $process = Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arguments -Wait -PassThru
    exit $process.ExitCode
}

$scriptRoot = Split-Path -Parent $PSCommandPath
$appSourcePath = Join-Path $scriptRoot "AIHelper.exe"
$exampleJsonSourcePath = Join-Path $scriptRoot "dns-presets-example.json"
$uninstallScriptSourcePath = Join-Path $scriptRoot "Uninstall-AIHelper.ps1"
$uninstallCmdSourcePath = Join-Path $scriptRoot "uninstall.cmd"

if (-not (Test-Path $appSourcePath)) {
    throw "Installer payload is incomplete: AIHelper.exe was not found."
}

$resolvedInstallDir = [IO.Path]::GetFullPath($InstallDir)
$assetsInstallDir = Join-Path $resolvedInstallDir "Assets"
$appTargetPath = Join-Path $resolvedInstallDir "AIHelper.exe"
$uninstallKeyPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\AIHelper"

Get-Process AIHelper -ErrorAction SilentlyContinue | Stop-Process -Force

New-Item -ItemType Directory -Path $resolvedInstallDir -Force | Out-Null
New-Item -ItemType Directory -Path $assetsInstallDir -Force | Out-Null

Copy-Item -Path $appSourcePath -Destination $appTargetPath -Force

if (Test-Path $exampleJsonSourcePath) {
    Copy-Item -Path $exampleJsonSourcePath -Destination (Join-Path $assetsInstallDir "dns-presets-example.json") -Force
}

Copy-Item -Path $uninstallScriptSourcePath -Destination (Join-Path $resolvedInstallDir "Uninstall-AIHelper.ps1") -Force
Copy-Item -Path $uninstallCmdSourcePath -Destination (Join-Path $resolvedInstallDir "uninstall.cmd") -Force

$fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($appTargetPath).ProductVersion
if ([string]::IsNullOrWhiteSpace($fileVersion)) {
    $fileVersion = "1.0.0"
}

New-Item -Path $uninstallKeyPath -Force | Out-Null
Set-ItemProperty -Path $uninstallKeyPath -Name "DisplayName" -Value "AIHelper"
Set-ItemProperty -Path $uninstallKeyPath -Name "Publisher" -Value "AIHelper"
Set-ItemProperty -Path $uninstallKeyPath -Name "DisplayVersion" -Value $fileVersion
Set-ItemProperty -Path $uninstallKeyPath -Name "InstallLocation" -Value $resolvedInstallDir
Set-ItemProperty -Path $uninstallKeyPath -Name "DisplayIcon" -Value $appTargetPath
Set-ItemProperty -Path $uninstallKeyPath -Name "UninstallString" -Value ('"' + (Join-Path $resolvedInstallDir "uninstall.cmd") + '"')
Set-ItemProperty -Path $uninstallKeyPath -Name "QuietUninstallString" -Value ('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $resolvedInstallDir "Uninstall-AIHelper.ps1") + '" -InstallDir "' + $resolvedInstallDir + '" -Quiet')
Set-ItemProperty -Path $uninstallKeyPath -Name "NoModify" -Value 1 -Type DWord
Set-ItemProperty -Path $uninstallKeyPath -Name "NoRepair" -Value 1 -Type DWord

if (-not $SkipShortcuts) {
    $desktopShortcutPath = Join-Path $env:Public "Desktop\AIHelper.lnk"
    $startMenuShortcutPath = Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs\AIHelper.lnk"
    New-Shortcut -ShortcutPath $desktopShortcutPath -TargetPath $appTargetPath -WorkingDirectory $resolvedInstallDir -Description "AIHelper"
    New-Shortcut -ShortcutPath $startMenuShortcutPath -TargetPath $appTargetPath -WorkingDirectory $resolvedInstallDir -Description "AIHelper"
}

if (-not $Quiet) {
    Write-Host "AIHelper installed to $resolvedInstallDir"
}
