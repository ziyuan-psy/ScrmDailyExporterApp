$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$IsccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)
$Iscc = $IsccCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $Iscc) {
    throw "Inno Setup 6 was not found. Install Inno Setup before running this script."
}

Push-Location $Root
try {
    if (-not (Test-Path -LiteralPath ".\dist\ScrmDailyExporter\scrm-exporter.exe")) {
        & .\build_release.ps1
    }
    & $Iscc ".\installer.iss"
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup compiler failed with exit code $LASTEXITCODE."
    }
    Write-Host ""
    Write-Host "Installer generated: $Root\installer-output\ScrmDailyExporterSetup.exe"
}
finally {
    Pop-Location
}
