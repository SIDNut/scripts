# ===== SafeClean-PatchCleanerPS.ps1 =====
param(
  [string]$Quarantine = "D:\InstallerQuarantine"  # change if needed
)

$ErrorActionPreference = 'Stop'
Write-Host "Quarantine: $Quarantine"
if (!(Test-Path $Quarantine)) { New-Item -ItemType Directory -Path $Quarantine | Out-Null }

# Download + load PatchCleanerPS (PowerShell version)
$pcUrl = "https://raw.githubusercontent.com/jackharvest/PatchCleanerPS/main/patchcleanerscript.ps1"
Write-Host "Fetching PatchCleanerPS script..."
$pcScript = Join-Path $env:TEMP "PatchCleanerPS.ps1"
Invoke-WebRequest -UseBasicParsing -Uri $pcUrl -OutFile $pcScript

. $pcScript   # dot-source functions

# Dry run first (prints findings)
Write-Host "`n--- Dry run (no changes) ---`n" -ForegroundColor Yellow
PatchCleanerPS -AutoDryAll

# Move orphans to quarantine (safe)
Write-Host "`n--- Moving orphaned MSI/MSP to quarantine ---`n" -ForegroundColor Cyan
PatchCleanerPS -MoveTo $Quarantine

Write-Host "`nDone. Keep the quarantine for a few days; delete later if no issues." -ForegroundColor Green
