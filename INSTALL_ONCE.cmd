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
echo Done. Office Local AutoSave is installed for this Windows user.
echo It will start automatically every time this user logs in.
echo You do not need to run this installer again after reboot.
echo.
pause
