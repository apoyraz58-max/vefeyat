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
$Key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
Write-Host ''
Write-Host "  $Name kaldiriliyor" -ForegroundColor Cyan
try {
    while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        Write-Host '  Word acik. Calismanizi kaydedip Word''un butun pencerelerini kapatin,' -ForegroundColor Yellow
        Read-Host '  sonra Enter''a basin' | Out-Null
    }
    if (Get-ItemProperty -Path $Key -Name $Id -ErrorAction SilentlyContinue) {
        Remove-ItemProperty -Path $Key -Name $Id
        Write-Host "  Tamam. $Name Word'den kaldirildi." -ForegroundColor Green
    } else {
        Write-Host "  $Name bu bilgisayarda kayitli degil."
    }
    if (Test-Path (Join-Path $Key $Id)) { Remove-Item (Join-Path $Key $Id) -Recurse -Force }
    Write-Host '  Belgelerinizdeki tarihler yerinde kalir.'
} catch {
    Write-Host ('  Kaldirilamadi: ' + $_.Exception.Message) -ForegroundColor Red
}
