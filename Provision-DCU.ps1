  <# 
    Dell Command Update: Install, Configure, Drivers & ApplyUpdates
    - Self-elevates if needed
    - Loads Garytown DCU functions
    - Installs DCU if missing
    - Applies a few safe DCU settings
    - Scans and applies all available updates with visible output
#>

# --- Self-elevate ---
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "Re-launching with administrative rights..."
    $psi = @{
        FilePath        = "powershell.exe"
        Verb            = "RunAs"
        ArgumentList    = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        WindowStyle     = "Normal"
    }
    Start-Process @psi
    exit
}

# --- Prep: Robust TLS & error display ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor `
                                              [Net.SecurityProtocolType]::Tls13

$ErrorActionPreference = "Stop"

# --- Load Garytown DCU functions (Install-DCU, Set-DCUSettings, Invoke-DCU, etc.) ---
# Source: https://garytown.com/dell-command-update-install-manage-via-powershell
# The blog provides a one-liner:  iex (irm dell.garytown.com)
Write-Host "Loading Garytown DCU helper functions..."
Invoke-Expression (Invoke-RestMethod -UseBasicParsing -Uri "https://dell.garytown.com")

# --- Helper: show DCU version if present ---
function Show-DCUVersion {
    try {
        $ver = Get-DCUVersion
        if ($ver) { Write-Host "Current DCU version: $ver" }
    } catch { Write-Host "DCU not detected yet." }
}

Show-DCUVersion

# --- Install DCU if missing ---
try {
    if (-not (Get-Command dcu-cli.exe -ErrorAction SilentlyContinue)) {
        Write-Host "Installing Dell Command | Update..."
        Install-DCU   # Uses Garytown function to resolve the correct installer for this model
    } else {
        Write-Host "DCU appears installed; verifying/updating if needed..."
        Install-DCU   # Function is safe to run to reconcile version
    }
} catch {
    Write-Warning "Install-DCU reported an issue: $($_.Exception.Message)"
}

Show-DCUVersion

# --- Configure (conservative/safe defaults) ---
# Note: Garytown’s Set-DCUSettings wraps DCU’s policy. These calls are defensive:
# If a parameter isn’t available in your version, it will be skipped with a warning.
Write-Host "Applying DCU settings..."
$settingsAttempts = @(
    { Set-DCUSettings -AutoSuspendBitLocker $true }, # let DCU handle BitLocker for BIOS/firmware safely
    { Set-DCUSettings -UpdateSeverity "recommended" }, # prefer recommended updates by default
    { Set-DCUSettings -ScheduleEnabled $false },       # no auto schedule; we’re running on-demand
    { Set-DCUSettings -RestartBehavior "required" }    # allow DCU to restart if the update requires it
)
foreach ($attempt in $settingsAttempts) {
    try { & $attempt } catch { Write-Warning "Skipping a DCU setting: $($_.Exception.Message)" }
}

# --- Run: Scan & Apply Updates (drivers/BIOS/firmware) ---
# Invoke-DCU wraps dcu-cli; we run a Scan first, then ApplyUpdates so techs can observe output.
Write-Host "`n===== DCU SCAN ====="
try {
    # Scan
    Invoke-DCU -Scan
} catch {
    Write-Warning "Invoke-DCU -Scan failed: $($_.Exception.Message)"
}

Write-Host "`n===== DCU APPLY UPDATES ====="
try {
    # ApplyUpdates (all applicable). DCU will handle restarts if required by settings above.
    Invoke-DCU -ApplyUpdates
} catch {
    Write-Warning "Invoke-DCU -ApplyUpdates failed: $($_.Exception.Message)"
}

Write-Host "`nAll done. Check output above for any items requiring a reboot or follow-up."
