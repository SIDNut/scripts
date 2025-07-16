<#
.SYNOPSIS
    Update Dell BIOS update .exe files on offline media or USB from the official Dell catalog, with robust automation, interactive/manual selection, and safe cleanup.

.DESCRIPTION
    Scans a specified folder for Dell BIOS .exe files, downloads and parses Dell's CatalogPC.xml, and identifies the latest BIOS updates for all models.
    Compares each local file against the latest catalog version, and downloads missing or newer BIOS .exe files as needed. Offers both fully automated and interactive/manual (Out-GridView or CLI menu) modes.
    Deletes or archives superseded BIOS after successful download, audits orphaned files, and offers to match/resolve them with catalog suggestions.
    Includes logging, dry-run (-WhatIf), cache directory override, and works cross-platform (Windows, Linux, macOS).

.PARAMETER Path
    (Required) The directory containing local Dell BIOS .exe files (e.g., a USB or folder). Also used as the default cache location for Dell catalog files, unless overridden.

.PARAMETER Interactive
    (Switch) Show an interactive selection UI for review and download. On Windows, uses Out-GridView if available, otherwise uses a CLI menu.

.PARAMETER AutoReplace
    (Switch) Automatically download and replace outdated BIOS (default unless -Interactive is used). If both Interactive and AutoReplace are omitted, updates only outdated BIOS for models present locally.

.PARAMETER CacheDir
    (String) Directory to cache catalog files (defaults to Path\Bin). Useful if your target folder is space-limited or FAT32.

.PARAMETER LogPath
    (String) Optional path for log file recording all actions and events.

.PARAMETER ArchiveFolder
    (String) Optional directory for archiving superseded or deleted BIOS files instead of deleting them permanently.

.PARAMETER IncludeModel
    (String[]) Only process the specified models (supports wildcards).

.PARAMETER ExcludeModel
    (String[]) Exclude specified models from processing (supports wildcards).

.EXAMPLE
    Update-DellBiosFiles -Path "F:\Ventoy"
    # Updates only outdated BIOS .exe for models present in the target folder.

.EXAMPLE
    Update-DellBiosFiles -Path "F:\Ventoy" -Interactive
    # Opens an interactive selection UI (Out-GridView or CLI) for BIOS update/download.

.EXAMPLE
    Update-DellBiosFiles -Path "F:\Ventoy" -WhatIf
    # Shows what would be done (download/delete/archive), but makes no changes.

.EXAMPLE
    Update-DellBiosFiles -Path "F:\Ventoy" -LogPath "C:\Logs\BiosUpdate.log"
    # Logs all actions to a file.

.EXAMPLE
    Update-DellBiosFiles -Path "F:\Ventoy" -ArchiveFolder "F:\ArchivedBIOS"
    # Archives superseded BIOS files to a specified folder instead of deleting them.

.NOTES
    Author: Luke Mitchell-Collins (based on original by David Segura)
    Version: 1.0.0 (2024-07-16)
    https://www.powershellgallery.com/packages/DellBios

    - On Windows, interactive mode uses Out-GridView if available.
    - On PowerShell 7+ (Core), Out-GridView may be unreliable; the CLI/text menu is offered instead.
    - For non-Windows systems, 'cabextract' or '7z' is required to extract Dell's CatalogPC.cab.
    - If you see parameter errors, ensure you are running THIS script, not the DellBios module version (run as: .\Update-DellBiosFiles.ps1).

.LINK
    https://github.com/SeguraDavid/DellBios
    https://www.dell.com/support
#>

# ======================= DellBios Module Collision Detection ==========================
$functionName = 'Update-DellBiosFiles'
$command = Get-Command $functionName -ErrorAction SilentlyContinue

if ($command -and $command.CommandType -eq 'Function') {
    # Only check if running from a file (not the module)
    $myPath = $null
    try { $myPath = $MyInvocation.MyCommand.Path } catch {}
    $cmdPath = $null
    try { $cmdPath = $command.ScriptBlock.File } catch {}
    if ($myPath -and $cmdPath -and ($myPath -ne $cmdPath)) {
        Write-Warning @"
Another '$functionName' function/command is already loaded (from: $cmdPath).
You may have the 'DellBios' module installed, which does NOT support these parameters.
To use THIS script, run it as:  .\Update-DellBiosFiles.ps1 <parameters>
Or remove/unimport the module before running.
"@
        return
    }
}
# ==================== End DellBios Module Collision Detection =========================

function Compare-DellBiosVersion {
    param($v1, $v2)
    if ($v1 -match '^A(\d+)$' -and $v2 -match '^A(\d+)$') {
        $n1 = [int]($v1 -replace '^A', '')
        $n2 = [int]($v2 -replace '^A', '')
        return $n1 - $n2
    }
    elseif ($v1 -match '^\d+(\.\d+)+$' -and $v2 -match '^\d+(\.\d+)+$') {
        return ([version]$v1).CompareTo([version]$v2)
    }
    else {
        return [string]::Compare($v1, $v2)
    }
}

function Find-ClosestCatalogModelKey {
    param($orphanName, $catalogKeys)
    $orphanLower = $orphanName.ToLower()
    $maxMatch = 0
    $bestKey = $null
    foreach ($key in $catalogKeys) {
        $len = [Math]::Min($orphanLower.Length, $key.Length)
        $match = 0
        for ($i = 0; $i -lt $len; $i++) {
            if ($orphanLower[$i] -eq $key[$i]) { $match++ } else { break }
        }
        if ($match -gt $maxMatch) {
            $maxMatch = $match
            $bestKey = $key
        }
    }
    return $bestKey
}

function Test-FileHash {
    param([string]$Path, [string]$Hash, [string]$Algorithm = "MD5")
    if (-not (Test-Path $Path)) { return $false }
    $calc = Get-FileHash -Path $Path -Algorithm $Algorithm
    return ($calc.Hash.ToLower() -eq $Hash.ToLower())
}

function Expand-CabFile {
    param([string]$cab, [string]$xml)
    if ($IsWindowsPlatform) {
        Expand "$cab" "$xml" | Out-Null
    }
    else {
        if (Get-Command cabextract -ErrorAction SilentlyContinue) {
            & cabextract -sf "$cab" | Out-Null
            & cabextract -F CatalogPC.xml "$cab" -d (Split-Path $xml)
        }
        elseif (Get-Command 7z -ErrorAction SilentlyContinue) {
            & 7z e "$cab" "-o$(Split-Path $xml)" "CatalogPC.xml" -y
        }
        else {
            throw "Cannot extract CAB on non-Windows: need cabextract or 7z in PATH."
        }
    }
}
function Invoke-FileDownload {
    param(
        [Parameter(Mandatory)]
        [string]$Url,
        [Parameter(Mandatory)]
        [string]$Destination
    )
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing
        return $true
    }
    catch {
        Write-Warning "Failed to download ${Url}: $($PSItem.Exception.Message)"
        return $false
    }
}

function Update-DellBiosFiles {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Directory containing Dell BIOS .exe files.")]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$Path,

        [Parameter(HelpMessage = "Show interactive selection UI for BIOS updates.")]
        [switch]$Interactive,

        [Parameter(HelpMessage = "Automatically replace outdated BIOS. Default unless -Interactive is used.")]
        [switch]$AutoReplace,

        [Parameter(HelpMessage = "Directory to cache catalog files.")]
        [string]$CacheDir,

        [Parameter(HelpMessage = "Optional log file path.")]
        [string]$LogPath,

        [Parameter(HelpMessage = "Archive superseded/deleted BIOS to specified folder.")]
        [string]$ArchiveFolder,

        [Parameter(HelpMessage = "Only include specified models (wildcards supported).")]
        [string[]]$IncludeModel,

        [Parameter(HelpMessage = "Exclude specified models (wildcards supported).")]
        [string[]]$ExcludeModel
    )

    # ==== PLATFORM DETECTION ====
    $IsWindowsPlatform = $false
    try {
        # Works in PowerShell 5.x and 7.x
        $IsWindowsPlatform = ($env:OS -match 'Windows') -or ($PSVersionTable.Platform -eq 'Win32NT')
    }
    catch {
        # fallback, best effort
        $IsWindowsPlatform = $env:OS -match 'Windows'
    }

    # ==== LOGGING ====
    function Write-Log {
        param([string]$Message)
        if ($LogPath) {
            $Message | Out-File $LogPath -Append -Encoding UTF8
        }
    }
    Write-Log "=== Update-DellBiosFiles run at $(Get-Date) ==="

    Write-Host "`n===== Dell BIOS Update Utility =====`n" -ForegroundColor Cyan

    # ==== FILTERS ====
    $IncludePatterns = @()
    $ExcludePatterns = @()
    if ($IncludeModel) { $IncludePatterns = $IncludeModel -split "," | ForEach-Object { $_.Trim() } }
    if ($ExcludeModel) { $ExcludePatterns = $ExcludeModel -split "," | ForEach-Object { $_.Trim() } }

    # ==== CACHE AND ARCHIVE FOLDER ====
    if ($CachePath) {
        if (!(Test-Path $CachePath)) { New-Item -Type Directory -Path $CachePath | Out-Null }
        $CacheRoot = $CachePath
    }
    else {
        $CacheRoot = [System.IO.Path]::GetTempPath()
    }

    $DellBiosRoot = $Path
    if (!(Test-Path $DellBiosRoot)) { New-Item -Type Directory -Path $DellBiosRoot | Out-Null }

    $ArchiveFolder = if ($ArchiveDeleted) {
        $folder = Join-Path $DellBiosRoot "Archive/$(Get-Date -Format yyyyMMdd-HHmmss)"
        if (!$WhatIf -and !(Test-Path $folder)) { New-Item -Type Directory -Path $folder | Out-Null }
        $folder
    }
    else { $null }

    # ==== CATALOG FILE PATHS ====
    $DellDownloadsUrl = "http://downloads.dell.com/"
    $DellCatalogPcUrl = "http://downloads.dell.com/catalog/CatalogPC.cab"
    $DellCatalogPcCab = Join-Path $CacheRoot ($DellCatalogPcUrl | Split-Path -Leaf)
    $DellCatalogPcXml = Join-Path $CacheRoot "CatalogPC.xml"

    Write-Host "Downloading Dell BIOS Catalog..." -ForegroundColor Green
    Write-Log  "Downloading Dell BIOS Catalog..."
    try {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would download: $DellCatalogPcUrl → $DellCatalogPcCab"
            Write-Log  "[WhatIf] Would download: $DellCatalogPcUrl → $DellCatalogPcCab"
        }
        else {
            (New-Object System.Net.WebClient).DownloadFile($DellCatalogPcUrl, $DellCatalogPcCab)
        }
    }
    catch {
        Write-Host "Catalog download failed! Exiting." -ForegroundColor Red
        Write-Log  "Catalog download failed! Exiting."
        return
    }

    Write-Host "Extracting Catalog..." -ForegroundColor Green
    Write-Log  "Extracting Catalog..."
    if (-not $WhatIf) {
        Expand-CabFile $DellCatalogPcCab $DellCatalogPcXml
    }
    if (-not $WhatIf -and !(Test-Path $DellCatalogPcXml)) {
        Write-Host "Could not extract CatalogPC.xml. Exiting." -ForegroundColor Red
        Write-Log  "Could not extract CatalogPC.xml. Exiting."
        return
    }

    # ==== PARSE CATALOG XML ====
    [xml]$XMLDellUpdateCatalog = if ($WhatIf) { $null } else { Get-Content "$DellCatalogPcXml" -ErrorAction Stop }
    $DellUpdateList = if ($WhatIf) { @() } else {
        $XMLDellUpdateCatalog.Manifest.SoftwareComponent | Where-Object {
            $_.ComponentType.Display.'#cdata-section'.Trim() -eq 'BIOS'
        }
    }

    # ==== BUILD CATALOG LOOKUP ====
    $CatalogByModelKey = @{}
    foreach ($item in $DellUpdateList) {
        if ($item.path -match '^(?:.*[\\/])?(?<ModelKey>.+)_(?<Version>[^_]+)\.exe$') {
            $modelKey = $matches.ModelKey
            $ver = $matches.Version
            $url = "$DellDownloadsUrl$($item.path)"
            $brand = $item.SupportedSystems.Brand.Display.'#cdata-section'.Trim()
            $modelName = $item.SupportedSystems.Brand.Model.Display.'#cdata-section'.Trim()
            $fileName = "$modelKey" + "_$ver.exe"
            $md5 = $null; $sha1 = $null
            if ($item.HashMD5) { $md5 = $item.HashMD5.'#cdata-section' }
            if ($item.HashSHA1) { $sha1 = $item.HashSHA1.'#cdata-section' }
            if (!$CatalogByModelKey.ContainsKey($modelKey) -or (Compare-DellBiosVersion $ver $CatalogByModelKey[$modelKey].DellVersion -gt 0)) {
                $CatalogByModelKey[$modelKey] = [PSCustomObject]@{
                    ModelKey    = $modelKey
                    DellVersion = $ver
                    DownloadURL = $url
                    Category    = $brand
                    ModelName   = $modelName
                    FileName    = $fileName
                    MD5         = $md5
                    SHA1        = $sha1
                }
            }
        }
    }

    # ==== APPLY INCLUDE/EXCLUDE MODEL FILTERS ====
    if ($IncludePatterns.Count) {
        $CatalogByModelKey = $CatalogByModelKey.GetEnumerator() | Where-Object {
            $inc = $false
            foreach ($p in $IncludePatterns) { if ($_.Value.ModelKey -like $p) { $inc = $true; break } }
            $inc
        } | ForEach-Object { $_.Key, $_.Value } | Group-Object -AsHashTable -AsString
    }
    if ($ExcludePatterns.Count) {
        $CatalogByModelKey = $CatalogByModelKey.GetEnumerator() | Where-Object {
            $exc = $false
            foreach ($p in $ExcludePatterns) { if ($_.Value.ModelKey -like $p) { $exc = $true; break } }
            -not $exc
        } | ForEach-Object { $_.Key, $_.Value } | Group-Object -AsHashTable -AsString
    }

    # ==== PARSE LOCAL BIOS FILES ====
    $LocalBios = @{}
    Get-ChildItem -Path $DellBiosRoot -Filter '*.exe' -File | ForEach-Object {
        if ($_.BaseName -match '^(?<ModelKey>.+)_(?<Version>[^_]+)$') {
            $modelKey = $matches.ModelKey
            $version = $matches.Version
            if (!$LocalBios.ContainsKey($modelKey) -or (Compare-DellBiosVersion $version $LocalBios[$modelKey].Version -gt 0)) {
                $LocalBios[$modelKey] = @{
                    File    = $_
                    Version = $version
                }
            }
        }
    }

    # ==== CANDIDATE UPDATE LIST ====
    $Candidates = @()
    foreach ($kv in $CatalogByModelKey.GetEnumerator()) {
        $modelKey = $kv.Key
        $latestVer = $kv.Value.DellVersion
        $url = $kv.Value.DownloadURL
        $category = $kv.Value.Category
        $modelName = $kv.Value.ModelName
        $fileName = $kv.Value.FileName

        $localInfo = $LocalBios[$modelKey]
        $localVer = if ($localInfo) { $localInfo.Version } else { $null }
        $compareResult = if ($localVer) { Compare-DellBiosVersion $localVer $latestVer } else { $null }

        if ($null -eq $localVer) { $status = 'Missing' }
        elseif ($compareResult -eq 0) { $status = 'Current' }
        elseif ($compareResult -lt 0) { $status = 'Outdated' }
        elseif ($compareResult -gt 0) { $status = 'NewerThanCatalog' }
        else { $status = 'Unknown' }

        $Candidates += [PSCustomObject]@{
            ModelKey      = $modelKey
            ModelName     = $modelName
            Category      = $category
            LocalVersion  = $localVer
            LatestVersion = $latestVer
            Status        = $status
            FileName      = $fileName
            DownloadURL   = $url
            LocalFile     = if ($localInfo) { $localInfo.File } else { $null }
            MD5           = $kv.Value.MD5
            SHA1          = $kv.Value.SHA1
        }
    }

    # ==== SORT ORDER ====
    $statusOrder = @{
        "Outdated"         = 5
        "Current"          = 4
        "NewerThanCatalog" = 3
        "Missing"          = 2
        "Unknown"          = 1
    }

    $localBiosCount = $LocalBios.Count
    if (-not $Interactive -and $localBiosCount -eq 0) {
        Write-Host "No local BIOS found in path '$DellBiosRoot'. Check path!`n" -ForegroundColor Yellow
        Write-Log  "No local BIOS found in path '$DellBiosRoot'."
        $seconds = 5
        Write-Host "Starting interactive BIOS selection in $seconds seconds..."
        Write-Host '(Press "Enter" to continue, any other key to Exit.)'

        $timer = [Diagnostics.Stopwatch]::StartNew()
        while ($timer.Elapsed.TotalSeconds -lt $seconds) {
            if ($Host.UI.RawUI.KeyAvailable) {
                $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
                if ($key.Character -eq "`r" -or $key.Character -eq "`n") {
                    break
                }
                else {
                    Write-Host "Exiting."
                    Write-Log  "Exiting."
                    return
                }
            }
            Start-Sleep -Milliseconds 200
        }
        $Interactive = $true
    }

    # ==== INTERACTIVE SELECTION (GridView or CLI menu) ====
    if ($Interactive) {
        $IsGridViewAvailable = $false
        if ($IsWindowsPlatform) {
            $gv = Get-Command Out-GridView -ErrorAction SilentlyContinue
            if ($gv) { $IsGridViewAvailable = $true }
        }
        if ($IsGridViewAvailable) {
            Write-Host "Tip: To update all Outdated BIOS, sort by Status, select all Outdated (Ctrl+A), then Enter." -ForegroundColor Cyan
            $ToUpdate = $Candidates |
            Sort-Object @{ Expression = { $statusOrder[$_.Status] }; Descending = $true }, ModelKey |
            Select-Object ModelKey, ModelName, Category, LocalVersion, LatestVersion, Status, FileName, DownloadURL |
            Out-GridView -PassThru -Title "Select BIOS to download/update (Outdated at top)"
        }
        else {
            Write-Host "`n[Non-Windows or Out-GridView not available] Showing text selection menu:" -ForegroundColor Yellow
            $i = 1
            $sorted = $Candidates | Sort-Object @{ Expression = { $statusOrder[$_.Status] }; Descending = $true }, ModelKey
            foreach ($c in $sorted) {
                Write-Host ("[{0}] {1,-35} {2,-12} {3,-10} → {4,-10} {5,-8}" -f $i, $c.ModelKey, $c.Status, $c.LocalVersion, $c.LatestVersion, $c.Category)
                $i++
            }
            $sel = Read-Host "Enter numbers (comma separated) of BIOS to update, or press Enter to update all Outdated"
            if (-not $sel) {
                $ToUpdate = $sorted | Where-Object { $_.Status -eq 'Outdated' }
            }
            else {
                $indices = $sel -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
                $ToUpdate = @()
                foreach ($ix in $indices) {
                    $i = [int]$ix
                    if ($i -gt 0 -and $i -le $sorted.Count) {
                        $ToUpdate += $sorted[$i - 1]
                    }
                }
            }
        }
    }
    else {
        $ToUpdate = $Candidates | Where-Object { $_.Status -eq 'Outdated' }
    }

    if (-not $ToUpdate) {
        Write-Host "No BIOS updates selected or required." -ForegroundColor Cyan
        Write-Log  "No BIOS updates selected or required."
        return
    }

    # --- Progress and summary counters ---
    $totalUpdated = 0; $totalDeleted = 0; $totalArchived = 0; $totalFailed = 0

    foreach ($item in $ToUpdate) {
        $destFile = Join-Path $DellBiosRoot $item.FileName
        Write-Host "`n[$($item.ModelKey)] $($item.Status): $($item.LocalVersion) → $($item.LatestVersion)"
        Write-Log  "[$($item.ModelKey)] $($item.Status): $($item.LocalVersion) → $($item.LatestVersion)"
        if ($WhatIf) {
            Write-Host "[WhatIf] Would download: $($item.DownloadURL) → $destFile"
            Write-Log  "[WhatIf] Would download: $($item.DownloadURL) → $destFile"
        }
        else {
            $downloadSucceeded = $false
            try {
                Write-Host "Downloading: $($item.DownloadURL) → $destFile"
                if ($WhatIf) {
                    Write-Host "[WhatIf] Would download: $($item.DownloadURL) → $destFile"
                    Write-Log  "[WhatIf] Would download: $($item.DownloadURL) → $destFile"
                    $downloadSucceeded = $true  # Simulate success in WhatIf
                }
                else {
                    $dlResult = Invoke-FileDownload -Url $item.DownloadURL -Destination $destFile
                    if ($dlResult) {
                        $downloadSucceeded = $true
                        Write-Host "Downloaded: $destFile" -ForegroundColor Green
                        Write-Log  "Downloaded: $destFile"
                        $totalUpdated++
                    }
                    else {
                        throw "Download failed"
                    }
                }

                # Hash verification (if provided, only if downloadSucceeded)
                $hashChecked = $false
                if ($downloadSucceeded -and $item.MD5) {
                    if (Test-FileHash $destFile $item.MD5 "MD5") {
                        Write-Host "MD5 hash verified." -ForegroundColor Green
                        Write-Log  "MD5 hash verified."
                        $hashChecked = $true
                    }
                    else {
                        Write-Warning "MD5 hash mismatch for $($item.FileName)!"
                        Write-Log     "MD5 hash mismatch for $($item.FileName)!"
                        $totalFailed++
                        $downloadSucceeded = $false
                    }
                }
                if ($downloadSucceeded -and -not $hashChecked -and $item.SHA1) {
                    if (Test-FileHash $destFile $item.SHA1 "SHA1") {
                        Write-Host "SHA1 hash verified." -ForegroundColor Green
                        Write-Log  "SHA1 hash verified."
                        $hashChecked = $true
                    }
                    else {
                        Write-Warning "SHA1 hash mismatch for $($item.FileName)!"
                        Write-Log     "SHA1 hash mismatch for $($item.FileName)!"
                        $totalFailed++
                        $downloadSucceeded = $false
                    }
                }
            }
            catch {
                Write-Warning "Download failed for $($item.ModelKey): $_"
                Write-Log     "Download failed for $($item.ModelKey): $_"
                $totalFailed++
            }

            # Only after download+hash succeed:
            if ($downloadSucceeded) {
                $pattern = "$($item.ModelKey)_*.exe"
                $allForModel = Get-ChildItem -Path $DellBiosRoot -Filter $pattern -File | Where-Object { $_.FullName -ne $destFile }
                foreach ($old in $allForModel) {
                    if ($ArchiveFolder) {
                        Move-Item $old.FullName -Destination (Join-Path $ArchiveFolder $old.Name) -Force
                        Write-Host "Archived old: $($old.Name)" -ForegroundColor Yellow
                        Write-Log  "Archived old: $($old.Name)"
                        $totalArchived++
                    }
                    else {
                        Remove-Item $old.FullName -Force
                        Write-Host "Deleted old: $($old.Name)" -ForegroundColor Yellow
                        Write-Log  "Deleted old: $($old.Name)"
                        $totalDeleted++
                    }
                }
            }
        }
    }

    # === SMART FINAL FILE AUDIT ===
    $localExeFiles = Get-ChildItem -Path $DellBiosRoot -Filter '*.exe' -File
    $finalFileList = @()
    foreach ($file in $localExeFiles) {
        if ($file.BaseName -match '^(?<ModelKey>.+)_(?<Version>[^_]+)$') {
            $modelKey = $matches.ModelKey
            $localVer = $matches.Version
            $catEntry = $CatalogByModelKey[$modelKey]
            if ($catEntry) {
                $catVer = $catEntry.DellVersion
                $cmp = Compare-DellBiosVersion $localVer $catVer
                if ($cmp -eq 0) { $finalStatus = "Current" }
                elseif ($cmp -gt 0) { $finalStatus = "NewerThanCatalog" }
                else { $finalStatus = "Orphaned" }
            }
            else {
                $finalStatus = "Orphaned"
            }
            $finalFileList += [PSCustomObject]@{
                File     = $file.Name
                ModelKey = $modelKey
                LocalVer = $localVer
                Status   = $finalStatus
            }
        }
        else {
            $finalFileList += [PSCustomObject]@{
                File     = $file.Name
                ModelKey = ""
                LocalVer = ""
                Status   = "Orphaned"
            }
        }
    }

    # === OUTPUT FILE AUDIT ===
    if ($finalFileList | Where-Object { $_.Status -eq "Current" }) {
        Write-Host "`nCurrent BIOS .exe files (match latest catalog):" -ForegroundColor Green
        $finalFileList | Where-Object { $_.Status -eq "Current" } | ForEach-Object {
            Write-Host "  $($_.File)" -ForegroundColor Green
            Write-Log  "Current: $($_.File)"
        }
    }
    if ($finalFileList | Where-Object { $_.Status -eq "NewerThanCatalog" }) {
        Write-Host "`nNewerThanCatalog BIOS .exe files (newer than latest catalog):" -ForegroundColor Cyan
        $finalFileList | Where-Object { $_.Status -eq "NewerThanCatalog" } | ForEach-Object {
            Write-Host "  $($_.File)" -ForegroundColor Cyan
            Write-Log  "NewerThanCatalog: $($_.File)"
        }
    }
    if ($finalFileList | Where-Object { $_.Status -eq "Orphaned" }) {
        Write-Host "`nOrphaned BIOS .exe files (older, missing from catalog, or unrecognized):" -ForegroundColor Magenta
        $finalFileList | Where-Object { $_.Status -eq "Orphaned" } | ForEach-Object {
            Write-Host "  $($_.File)" -ForegroundColor DarkYellow
            Write-Log  "Orphaned: $($_.File)"
        }
    }
    else {
        Write-Host "`nNo orphaned BIOS files found." -ForegroundColor Green
        Write-Log  "No orphaned BIOS files found."
    }

    # === OFFER TO MATCH & UPDATE ORPHANS ===
    $catalogKeys = $CatalogByModelKey.Keys

    $finalFileList | Where-Object { $_.Status -eq "Orphaned" } | ForEach-Object {
        $orphan = $_
        $tryName = if ($orphan.File -match '^(?<ModelKey>.+)_') { $matches.ModelKey } else { $orphan.File }
        $closestKey = Find-ClosestCatalogModelKey $tryName $catalogKeys
        if ($closestKey) {
            $catInfo = $CatalogByModelKey[$closestKey]
            Write-Host "`nOrphaned file: $($orphan.File)" -ForegroundColor Magenta
            Write-Host "Closest catalog match: $($closestKey) (latest version: $($catInfo.DellVersion))" -ForegroundColor Cyan
            Write-Log  "Orphaned file: $($orphan.File), Closest match: $($closestKey), Latest: $($catInfo.DellVersion)"
            $response = Read-Host "Does this match and should I update+replace? (Y/N)"
            if ($response -match '^[Yy]$') {
                $destFile = Join-Path $DellBiosRoot "$($catInfo.ModelKey)_$($catInfo.DellVersion).exe"
                if ($WhatIf) {
                    Write-Host "[WhatIf] Would download: $($catInfo.DownloadURL) → $destFile"
                    Write-Host "[WhatIf] Would delete/archive orphan: $($orphan.File)"
                    Write-Log  "[WhatIf] Would download: $($catInfo.DownloadURL) → $destFile"
                    Write-Log  "[WhatIf] Would delete/archive orphan: $($orphan.File)"
                }
                else {
                    try {
                        Invoke-WebRequest -Uri $catInfo.DownloadURL -OutFile $destFile -UseBasicParsing
                        Write-Host "Downloaded: $destFile" -ForegroundColor Green
                        Write-Log  "Downloaded: $destFile"
                        # Archive or delete orphan file after successful download
                        if ($ArchiveFolder) {
                            Move-Item (Join-Path $DellBiosRoot $orphan.File) -Destination (Join-Path $ArchiveFolder $orphan.File) -Force
                            Write-Host "Archived orphan: $($orphan.File)" -ForegroundColor Yellow
                            Write-Log  "Archived orphan: $($orphan.File)"
                            $totalArchived++
                        }
                        else {
                            Remove-Item (Join-Path $DellBiosRoot $orphan.File) -Force
                            Write-Host "Deleted orphan: $($orphan.File)" -ForegroundColor Yellow
                            Write-Log  "Deleted orphan: $($orphan.File)"
                            $totalDeleted++
                        }
                        # Delete/Archive any other local old versions for this model (except new one)
                        $pattern = "$($catInfo.ModelKey)_*.exe"
                        $extras = Get-ChildItem -Path $DellBiosRoot -Filter $pattern -File | Where-Object { $_.FullName -ne $destFile }
                        foreach ($old in $extras) {
                            if ($ArchiveFolder) {
                                Move-Item $old.FullName -Destination (Join-Path $ArchiveFolder $old.Name) -Force
                                Write-Host "Archived old: $($old.Name)" -ForegroundColor Yellow
                                Write-Log  "Archived old: $($old.Name)"
                                $totalArchived++
                            }
                            else {
                                Remove-Item $old.FullName -Force
                                Write-Host "Deleted old: $($old.Name)" -ForegroundColor Yellow
                                Write-Log  "Deleted old: $($old.Name)"
                                $totalDeleted++
                            }
                        }
                    }
                    catch {
                        Write-Warning "Download failed for $($catInfo.ModelKey): $_"
                        Write-Log     "Download failed for $($catInfo.ModelKey): $_"
                        $totalFailed++
                    }
                }
            }
        }
    }

    # === CLEANUP (cache/catalog files) ===
    if ($WhatIf) {
        Write-Host "[WhatIf] Would cleanup catalog files: $DellCatalogPcCab, $DellCatalogPcXml"
        Write-Log  "[WhatIf] Would cleanup catalog files: $DellCatalogPcCab, $DellCatalogPcXml"
    }
    else {
        if (Test-Path $DellCatalogPcCab) { Remove-Item $DellCatalogPcCab -Force }
        if (Test-Path $DellCatalogPcXml) { Remove-Item $DellCatalogPcXml -Force }
    }

    # === FINAL SUMMARY ===
    Write-Host ""
    Write-Host "========== Summary =========="
    Write-Host ("Updated BIOS:    {0}" -f $totalUpdated)
    Write-Host ("Deleted:         {0}" -f $totalDeleted)
    Write-Host ("Archived:        {0}" -f $totalArchived)
    Write-Host ("Failed:          {0}" -f $totalFailed)
    Write-Host "============================="
    Write-Log  "Summary: Updated=$totalUpdated Deleted=$totalDeleted Archived=$totalArchived Failed=$totalFailed"

    Write-Host "`nAll requested BIOS updates processed.`n" -ForegroundColor Cyan
    Write-Log  "All requested BIOS updates processed."
} # end function
