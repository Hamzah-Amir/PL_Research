# PL Research Tool — environment setup
# Run once before first use:  .\setup.ps1

Set-Location $PSScriptRoot

Write-Host "Creating virtual environment..." -ForegroundColor Cyan
python -m venv venv
if (-not $?) { Write-Host "ERROR: python not found. Install Python 3.10+ and try again." -ForegroundColor Red; exit 1 }

Write-Host "Activating virtual environment..." -ForegroundColor Cyan
& "$PSScriptRoot\venv\Scripts\Activate.ps1"

Write-Host "Installing dependencies..." -ForegroundColor Cyan
pip install --upgrade pip
pip install -r requirements.txt

Write-Host ""
Write-Host "Setup complete. To run Phase 1:" -ForegroundColor Green
Write-Host "  .\run.ps1" -ForegroundColor Yellow
