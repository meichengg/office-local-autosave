@echo off
setlocal
cd /d "%~dp0"

echo Installing Office Local AutoSave...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
    echo.
    echo Install failed.
    pause
    exit /b 1
)

echo.
echo Done. Office Local AutoSave is installed for all Windows users.
echo It will start automatically every time any Windows user logs in.
echo It runs in the background without showing a PowerShell window.
echo You do not need to run this installer again after reboot.
echo.
pause
