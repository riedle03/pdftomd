$ErrorActionPreference = "Stop"

Write-Host "[1/2] Installing Python packages..."
py -m pip install -r requirements.txt

Write-Host "[2/2] Installing hwpjs globally..."
npm install -g @ohah/hwpjs

Write-Host "Setup completed."
