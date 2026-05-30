# PL Research Tool — Phase 1 launcher
# Usage:  .\run.ps1

Set-Location $PSScriptRoot

$activate = "$PSScriptRoot\venv\Scripts\Activate.ps1"
if (Test-Path $activate) {
    & $activate
} else {
    Write-Host "Virtual environment not found. Run .\setup.ps1 first." -ForegroundColor Red
    exit 1
}

python run.py
