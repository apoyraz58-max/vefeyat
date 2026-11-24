@echo off
rem {{NAME}}: Word'den kaldirir. Belgelerinizdeki tarihler yerinde kalir.
setlocal
set "VEFEYAT_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:VEFEYAT_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf('#'+'POWERSHELL'); Invoke-Expression $t.Substring($i)"
echo.
pause
exit /b
#POWERSHELL
$ErrorActionPreference = 'Stop'
$Id = '{{ID}}'
$Name = '{{NAME}}'
$Channel = '{{CHANNEL}}'
$Folder = Join-Path $env:LOCALAPPDATA "Vefeyat\$Channel"
$Developer = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$CatalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$Id}"
Write-Host ''
Write-Host "  $Name kaldiriliyor" -ForegroundColor Cyan
try {
    while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        Write-Host '  Word acik. Calismanizi kaydedip Word''un butun pencerelerini kapatin,' -ForegroundColor Yellow
        Read-Host '  sonra Enter''a basin' | Out-Null
    }
    $found = $false
    # The lasting install (Paylasilan Klasor catalog) and the earlier one (Gelistirici Eklentileri).
    if (Test-Path $CatalogKey) { Remove-Item -Path $CatalogKey -Recurse -Force; $found = $true }
    if (Get-ItemProperty -Path $Developer -Name $Id -ErrorAction SilentlyContinue) { Remove-ItemProperty -Path $Developer -Name $Id; $found = $true }
    if (Test-Path (Join-Path $Developer $Id)) { Remove-Item (Join-Path $Developer $Id) -Recurse -Force }
    if (Test-Path $Folder) { Remove-Item -Path $Folder -Recurse -Force }
    if ($found) {
        Write-Host "  Tamam. $Name Word'den kaldirildi; sekme Word bir sonraki acilista gorunmez." -ForegroundColor Green
    } else {
        Write-Host "  $Name bu bilgisayarda kayitli degil."
    }
    Write-Host '  Belgelerinizdeki tarihler yerinde kalir.'
} catch {
    Write-Host ('  Kaldirilamadi: ' + $_.Exception.Message) -ForegroundColor Red
}
