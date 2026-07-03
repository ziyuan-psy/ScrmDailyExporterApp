$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Root
try {
    $Python = Join-Path $Root ".venv\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $Python)) {
        $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if ($PythonCommand) {
            $Python = $PythonCommand.Source
        }
    }
    if (-not (Test-Path -LiteralPath $Python)) {
        throw "python was not found. Install Python 3.10+ or run this on a build machine with Python."
    }

    & $Python -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw "pip upgrade failed with exit code $LASTEXITCODE."
    }
    & $Python -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        throw "pip install failed with exit code $LASTEXITCODE."
    }
    & $Python -m PyInstaller --clean --noconfirm .\ScrmDailyExporter.spec
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE."
    }

    Write-Host ""
    Write-Host "Build complete: $Root\dist\ScrmDailyExporter"
    Write-Host "CLI: $Root\dist\ScrmDailyExporter\scrm-exporter.exe"
    Write-Host "UI: $Root\dist\ScrmDailyExporter\scrm-exporter-ui.exe"
}
finally {
    Pop-Location
}
