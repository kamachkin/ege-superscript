@echo off
:: Launcher for SuperScript - bypasses execution policy and elevates to admin
:: Place this file alongside superscript_modified.ps1 or in the 'soft' folder

:: Get the directory where this .cmd file is located
set "SCRIPT_DIR=%~dp0"

:: Check if superscript_modified.ps1 exists in this directory
if exist "%SCRIPT_DIR%superscript_modified.ps1" (
    set "PS1_PATH=%SCRIPT_DIR%superscript_modified.ps1"
) else (
    echo ERROR: superscript_modified.ps1 not found in %SCRIPT_DIR%
    pause
    exit /b 1
)

:: Launch PowerShell with Bypass policy and admin elevation
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%PS1_PATH%\"' -Verb RunAs"
