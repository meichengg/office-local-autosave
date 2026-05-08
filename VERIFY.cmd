@echo off
setlocal
cd /d "%~dp0"

echo Verifying Office Local AutoSave...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0verify.ps1"
if errorlevel 1 (
    echo.
    echo Verify failed.
    pause
    exit /b 1
)

echo.
echo Verify completed.
echo.
pause
