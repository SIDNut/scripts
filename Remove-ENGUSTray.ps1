# Remove en-US (0409) language/keyboard remnants for current user and default profile

# 1) Remove en-US from current user's language list (if present)
try {
    $L = Get-WinUserLanguageList
    $L2 = $L | Where-Object LanguageTag -ne 'en-US'
    if ($L2.Count -ne $L.Count) {
        Set-WinUserLanguageList $L2 -Force
        Write-Host "Removed en-US from language list."
    }
} catch {
    Write-Warning "Could not adjust WinUserLanguageList: $_"
}

# 2) Purge 0409 keyboard entries from common registry hives
$keys = @(
    'HKCU:\Keyboard Layout\Preload',
    'HKCU:\Keyboard Layout\Substitutes',
    'HKU\.DEFAULT\Keyboard Layout\Preload',
    'HKU\.DEFAULT\Keyboard Layout\Substitutes'
)

$removed = $false
foreach ($k in $keys) {
    if (-not (Test-Path $k)) { continue }
    try {
        $propNames = (Get-Item $k).Property
        foreach ($name in $propNames) {
            $val = (Get-ItemProperty -Path $k -Name $name).$name
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
