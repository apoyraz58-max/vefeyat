@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall.ps1"
set "vefeyat_exit=%ERRORLEVEL%"
if not "%vefeyat_exit%"=="0" echo Kaldirma tamamlanamadi. Yukaridaki aciklamayi okuyun.
pause
exit /b %vefeyat_exit%
