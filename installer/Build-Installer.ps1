param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$installerRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path -Parent $installerRoot
$projectFile = Join-Path $projectRoot "LaptopSessionViewer.csproj"
$assetsExampleJson = Join-Path $projectRoot "Assets\dns-presets-example.json"
$supportRoot = Join-Path $installerRoot "Support"
$artifactsRoot = Join-Path $installerRoot "artifacts"
$publishRoot = Join-Path $artifactsRoot "publish-single"
$packageRoot = Join-Path $artifactsRoot "package"
$distRoot = Join-Path $projectRoot "dist"
$sedPath = Join-Path $artifactsRoot "AIHelper-Setup.sed"
$outputSetupPath = Join-Path $distRoot "AIHelper-Setup.exe"

Remove-Item $publishRoot -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $packageRoot -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $sedPath -Force -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Path $publishRoot -Force | Out-Null
New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
New-Item -ItemType Directory -Path $distRoot -Force | Out-Null

& dotnet publish $projectFile `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $publishRoot

$appExePath = Join-Path $publishRoot "AIHelper.exe"

if (-not (Test-Path $appExePath)) {
    throw "Self-contained publish did not produce AIHelper.exe."
}

Copy-Item $appExePath -Destination (Join-Path $packageRoot "AIHelper.exe") -Force
Copy-Item $assetsExampleJson -Destination (Join-Path $packageRoot "dns-presets-example.json") -Force
Copy-Item (Join-Path $supportRoot "install.cmd") -Destination (Join-Path $packageRoot "install.cmd") -Force
Copy-Item (Join-Path $supportRoot "Install-AIHelper.ps1") -Destination (Join-Path $packageRoot "Install-AIHelper.ps1") -Force
Copy-Item (Join-Path $supportRoot "uninstall.cmd") -Destination (Join-Path $packageRoot "uninstall.cmd") -Force
Copy-Item (Join-Path $supportRoot "Uninstall-AIHelper.ps1") -Destination (Join-Path $packageRoot "Uninstall-AIHelper.ps1") -Force

$sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles
[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=AIHelper setup has finished.
TargetName=$outputSetupPath
FriendlyName=AIHelper Setup
AppLaunched=cmd.exe /d /s /c ""install.cmd""
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
FILE0="AIHelper.exe"
FILE1="dns-presets-example.json"
FILE2="install.cmd"
FILE3="Install-AIHelper.ps1"
FILE4="uninstall.cmd"
FILE5="Uninstall-AIHelper.ps1"
[SourceFiles]
SourceFiles0=$packageRoot\
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
%FILE3%=
%FILE4%=
%FILE5%=
"@

[IO.File]::WriteAllText($sedPath, $sedContent, [Text.Encoding]::ASCII)

& iexpress.exe /N /Q $sedPath

if (-not (Test-Path $outputSetupPath)) {
    throw "IExpress did not create the setup executable."
}

Write-Host "Installer created: $outputSetupPath"
