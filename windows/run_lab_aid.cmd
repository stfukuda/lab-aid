@echo off
setlocal

rem ===============================================
rem Lab-Aid execution launcher
rem ===============================================

set "ROOT=%~dp0.."
set "PYTHON_EXE=%ROOT%\windows\runtime\python\python.exe"
set "LAUNCHER=%ROOT%\lab_aid.bat"

if not exist "%PYTHON_EXE%" (
    echo.
    echo [ERROR] Embedded Python runtime was not found.
    echo         Please run windows\setup_lab_aid.cmd first.
    echo.
    pause
    exit /b 1
)

if exist "%LAUNCHER%" (
    call "%LAUNCHER%" %*
) else (
    "%PYTHON_EXE%" -m lab_aid.excel_cli %*
)

if errorlevel 1 (
    echo.
    echo [ERROR] Lab-Aid execution failed.
) else (
    echo.
    echo Lab-Aid finished successfully.
)

echo.
pause
endlocal
