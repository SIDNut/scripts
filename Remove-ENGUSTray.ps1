# Remove en-US from the language list, and purge 0409 keyboard entries from common registry spots
$L=Get-WinUserLanguageList; Set-WinUserLanguageList ($L | ? LanguageTag -ne 'en-US') -Force
$keys=@(
 'HKCU:\Keyboard Layout\Preload',
 'HKCU:\Keyboard Layout\Substitutes',
 'HKU\.DEFAULT\Keyboard Layout\Preload',
 'HKU\.DEFAULT\Keyboard Layout\Substitutes'
)
foreach($k in $keys){
  if(Test-Path $k){
    (Get-ItemProperty -Path $k | Select-Object -Expand PSObject).Properties |
      ? { $_.MemberType -eq 'NoteProperty' -and ($_.Value -match '^00000409$' -or $_.Value -match '0409$') } |
      % { Remove-ItemProperty -Path $k -Name $_.Name -ErrorAction SilentlyContinue }
  }
}
Write-Host 'US (0409) layouts removed. Sign out/in to refresh the tray.'
