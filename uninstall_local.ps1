$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$InstallRoot = Join-Path $env:LOCALAPPDATA "ScrmDailyExporter"
$AppDir = Join-Path $InstallRoot "app"
$CliExe = Join-Path $AppDir "scrm-exporter.exe"

if (Test-Path -LiteralPath $CliExe) {
    & $CliExe uninstall-task
}
else {
    Write-Host "CLI not found; scheduled task removal skipped."
}

$ProgramsDir = [Environment]::GetFolderPath("Programs")
$ShortcutDir = Join-Path $ProgramsDir "SCRM Daily Exporter"
if (Test-Path -LiteralPath $ShortcutDir) {
    Remove-Item -LiteralPath $ShortcutDir -Recurse -Force
}

$Desktop = [Environment]::GetFolderPath("Desktop")
$DesktopLink = Join-Path $Desktop "SCRM Daily Exporter.lnk"
if (Test-Path -LiteralPath $DesktopLink) {
    Remove-Item -LiteralPath $DesktopLink -Force
}

if (Test-Path -LiteralPath $AppDir) {
    Remove-Item -LiteralPath $AppDir -Recurse -Force
}

Write-Host "Program removed. Runtime data is kept at: $InstallRoot"
