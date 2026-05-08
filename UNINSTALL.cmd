@echo off
setlocal
cd /d "%~dp0"

echo Uninstalling Office Local AutoSave...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
if errorlevel 1 (
    echo.
    echo Uninstall failed.
    pause
    exit /b 1
)

echo.
echo Done. Office Local AutoSave has been removed from startup.
echo Existing backup/log files are kept in %%LOCALAPPDATA%%\OfficeLocalAutoSave.
echo.
pause
