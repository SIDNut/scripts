# --- SAFER LANGUAGE / KEYBOARD CLEANUP (removes en-US, ensures a valid default, fixes override) ---

$ErrorActionPreference = 'Stop'

function Get-SafeInputTip($langs){
    # Pick the first valid TIP from the language list
    foreach ($l in $langs) {
        if ($l.InputMethodTips -and $l.InputMethodTips.Count -gt 0) {
            return $l.InputMethodTips[0]
        }
    }
    return $null
}

try {
    # 1) Read current user's language list
    $langs = Get-WinUserLanguageList

    # 2) If removing en-US would make the list empty, add en-AU first (adjust to your preferred language)
    $haveOnlyEnUS = ($langs.Count -eq 1 -and $langs[0].LanguageTag -eq 'en-US')
    if ($haveOnlyEnUS) {
        $langs.Add([Windows.Globalization.Language]::new('en-AU'))
    }

    # 3) Remove en-US if present
    $langs2 = $langs | Where-Object LanguageTag -ne 'en-US'

    # 4) Apply the language list
    if ($langs2.Count -gt 0) {
        Set-WinUserLanguageList -LanguageList $langs2 -Force
        Write-Host "Updated user language list: $($langs2.LanguageTag -join ', ')"
    } else {
        Write-Warning "Refusing to apply an empty language list. Add at least one language first."
    }

    # 5) Ensure Windows has a readable default input method (repair corrupt override)
    # Try to compute a safe TIP from the current list
    $langsNow = Get-WinUserLanguageList
    $safeTip  = Get-SafeInputTip $langsNow

    # If we couldn't compute a TIP, fall back to a generic en-AU TIP (Windows will coerce as needed)
    if (-not $safeTip) { $safeTip = '0c09:00000409' }  # en-AU langID with a common layout TIP

    try {
        # Clear any broken override first
        Remove-ItemProperty -Path 'HKCU:\Control Panel\International' -Name 'InputMethodOverride' -ErrorAction SilentlyContinue
    } catch {}

    # Set a good override
    Set-WinDefaultInputMethodOverride -InputTip $safeTip
    Write-Host "Default input method set to: $safeTip"

} catch {
    Write-Warning "Could not adjust WinUserLanguageList / default TIP: $_"
}

# 6) Purge US (0409) keyboard entries from current user + Default profile
$keys = @(
    'HKCU:\Keyboard Layout\Preload',
    'HKCU:\Keyboard Layout\Substitutes',
    'HKU:\.DEFAULT\Keyboard Layout\Preload',
    'HKU:\.DEFAULT\Keyboard Layout\Substitutes'
)

$removed = $false
foreach ($k in $keys) {
    if (-not (Test-Path $k)) { continue }
    try {
        $propNames = (Get-Item $k).Property
        foreach ($name in $propNames) {
            $val = (Get-ItemProperty -Path $k -Name $name -ErrorAction SilentlyContinue).$name
            if ($val -match '^(00000409|.*0409)$') {
                Remove-ItemProperty -Path $k -Name $name -ErrorAction SilentlyContinue
                $removed = $true
            }
        }
    } catch {
        Write-Warning "Error processing $k : $_"
    }
}

if ($removed) {
    Write-Host "US (0409) layouts removed. Sign out/in to refresh the tray."
} else {
    Write-Host "No 0409 layouts found; nothing to remove."
}

# 7) (Optional but helpful) Nuke cached Intl files if things still look odd
# Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Control Panel\International\*" -ErrorAction SilentlyContinue
# They will be recreated next sign-in.
