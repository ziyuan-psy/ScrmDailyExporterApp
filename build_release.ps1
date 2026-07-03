$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Root
try {
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        throw "python was not found. Install Python 3.10+ or run this on a build machine with Python."
    }

    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    python -m PyInstaller --clean .\ScrmDailyExporter.spec

    Write-Host ""
    Write-Host "Build complete: $Root\dist\ScrmDailyExporter"
    Write-Host "CLI: $Root\dist\ScrmDailyExporter\scrm-exporter.exe"
    Write-Host "UI: $Root\dist\ScrmDailyExporter\scrm-exporter-ui.exe"
}
finally {
    Pop-Location
}
