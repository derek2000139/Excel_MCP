@echo off
chcp 65001 >nul
title ExcelForge Installer
color 0A

REM Check if PowerShell exists
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] PowerShell not found
    pause
    exit /b 1
)

REM Run PowerShell script with bypass policy
call powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0install.ps1"

set EXITCODE=%errorlevel%
echo.
echo [INFO] Exit code: %EXITCODE%

if %EXITCODE% neq 0 (
    echo.
    echo [ERROR] Installation failed
    echo.
)

echo.
pause
exit /b %EXITCODE%
