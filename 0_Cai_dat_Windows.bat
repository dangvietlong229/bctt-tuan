@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup_windows.ps1"
set "status=%ERRORLEVEL%"
echo.
pause
exit /b %status%
