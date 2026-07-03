param(
    [switch]$DesktopShortcut,
    [switch]$TestMode
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Source = Join-Path $Root "dist\ScrmDailyExporter"
$InstallRoot = Join-Path $env:LOCALAPPDATA "ScrmDailyExporter"
$AppDir = Join-Path $InstallRoot "app"
$UiExe = Join-Path $AppDir "scrm-exporter-ui.exe"
$CliExe = Join-Path $AppDir "scrm-exporter.exe"

if (-not (Test-Path -LiteralPath (Join-Path $Source "scrm-exporter.exe"))) {
    throw "Build output not found: $Source. Run .\build_release.ps1 first."
}

New-Item -ItemType Directory -Force -Path $AppDir | Out-Null
Copy-Item -Path (Join-Path $Source "*") -Destination $AppDir -Recurse -Force

$ProgramsDir = [Environment]::GetFolderPath("Programs")
$ShortcutDir = Join-Path $ProgramsDir "SCRM Daily Exporter"
New-Item -ItemType Directory -Force -Path $ShortcutDir | Out-Null
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut((Join-Path $ShortcutDir "SCRM Daily Exporter.lnk"))
$Shortcut.TargetPath = $UiExe
$Shortcut.WorkingDirectory = $AppDir
$Shortcut.Save()

if ($DesktopShortcut) {
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $DesktopLink = $Shell.CreateShortcut((Join-Path $Desktop "SCRM Daily Exporter.lnk"))
    $DesktopLink.TargetPath = $UiExe
    $DesktopLink.WorkingDirectory = $AppDir
    $DesktopLink.Save()
}

$InstallTaskArgs = @("install-task")
if ($TestMode) {
    $InstallTaskArgs += "--test-mode"
}
& $CliExe @InstallTaskArgs

Write-Host ""
Write-Host "Install complete: $AppDir"
Write-Host "Start menu: SCRM Daily Exporter"
