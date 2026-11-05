@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
set "vefeyat_exit=%ERRORLEVEL%"
if not "%vefeyat_exit%"=="0" echo Kurulum tamamlanamadi. Yukaridaki aciklamayi okuyun.
pause
exit /b %vefeyat_exit%
