@echo off
setlocal

rem ===============================================
rem Lab-Aid Embedded Python Setup Launcher
rem ===============================================

set "SCRIPT_DIR=%~dp0"
set "POWERSHELL=PowerShell -NoLogo -NoProfile -ExecutionPolicy Bypass"

echo.
echo Starting Lab-Aid runtime setup.
echo Please keep this window open until the process finishes.
echo.

%POWERSHELL% -File "%SCRIPT_DIR%setup_lab_aid.ps1" %*
if errorlevel 1 (
    echo.
    echo [ERROR] Setup failed.
) else (
    echo.
    echo Setup completed successfully.
    echo Please edit windows\lab_aid_input.xlsx and run windows\run_lab_aid.cmd to execute Lab-Aid.
)

echo.
pause
endlocal
