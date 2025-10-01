# Update-Huntcore.ps1
# This script updates the Huntcore.extension Git repo and refreshes pyRevit

# Path to your Huntcore.extension
$extensionPath = "C:\pyRevit\Git_Repository\Huntcore.extension"

# Navigate to the repo folder
Set-Location $extensionPath

# Pull latest changes from Git
Write-Host "Pulling latest changes from Git..." -ForegroundColor Cyan
git pull

# Refresh the pyRevit extension
Write-Host "Refreshing pyRevit extension..." -ForegroundColor Cyan
pyrevit extensions update Huntcore

Write-Host "Huntcore.extension is now up-to-date!" -ForegroundColor Green
