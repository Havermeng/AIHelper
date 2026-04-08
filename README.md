# AIHelper

AIHelper is a Windows WPF desktop app for working with local Codex sessions.

## Features

- View local Codex sessions from `%USERPROFILE%\.codex\sessions`
- Search, inspect, favorite, annotate, delete, and resume sessions
- Launch a new Codex session from the app
- Check the local Codex environment
- Manage Windows DNS settings with presets, custom presets, import/export, and rollback
- Switch the UI language between English and Russian

## Requirements

- Windows 10/11 x64
- For session management and launch features: Codex CLI installed locally
- For DNS changes: administrator rights

## Installation

Download the latest `AIHelper-Setup.exe` from the GitHub Releases page and run it.

If Windows SmartScreen appears, use `More info` -> `Run anyway`. The app is locally built and not code-signed.

## Build From Source

```powershell
dotnet build .\LaptopSessionViewer.csproj -c Release
dotnet publish .\LaptopSessionViewer.csproj -c Release -r win-x64 --self-contained false -o .\publish\win-x64
```

## Build Installer

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\installer\Build-Installer.ps1
```

The generated installer is written to:

```text
dist\AIHelper-Setup.exe
```

## Notes

- DNS presets are stored in `%AppData%\AIHelper\dns-presets.json`
- Session favorites and notes are stored in `%USERPROFILE%\.codex`
- The example DNS preset JSON file is included in `Assets\dns-presets-example.json`
