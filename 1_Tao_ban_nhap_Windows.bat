@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\run_weekly_report_windows.ps1" -ReportCommand all
set "status=%ERRORLEVEL%"
echo.
if "%status%"=="0" (
  echo Da tao ban nhap. Xem output_reports\latest_run.json de tim file review.
) else (
  echo Co loi. Vui long xem thong bao phia tren.
)
echo.
pause
exit /b %status%
