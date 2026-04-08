param(
    [string]$InstallDir = "",
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

    if ($Quiet) {
        $arguments += "-Quiet"
    }

    $process = Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arguments -Wait -PassThru
    exit $process.ExitCode
}

$resolvedInstallDir = [IO.Path]::GetFullPath($InstallDir)
$uninstallKeyPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\AIHelper"
$desktopShortcutPath = Join-Path $env:Public "Desktop\AIHelper.lnk"
$startMenuShortcutPath = Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs\AIHelper.lnk"

Get-Process AIHelper -ErrorAction SilentlyContinue | Stop-Process -Force

if (Test-Path $desktopShortcutPath) {
    Remove-Item $desktopShortcutPath -Force
}

if (Test-Path $startMenuShortcutPath) {
    Remove-Item $startMenuShortcutPath -Force
}

if (Test-Path $uninstallKeyPath) {
    Remove-Item $uninstallKeyPath -Recurse -Force
}

if (Test-Path $resolvedInstallDir) {
    Remove-Item $resolvedInstallDir -Recurse -Force
}

if (-not $Quiet) {
    Write-Host "AIHelper removed from $resolvedInstallDir"
}
