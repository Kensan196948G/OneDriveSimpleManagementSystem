@echo off
chcp 65001 > nul
setlocal

echo =====================================================
echo   PowerShell Script Encoding Fix Tool
echo =====================================================
echo.
echo This tool fixes script encoding to UTF-8
echo.

:: Get current directory
set CURRENT_DIR=%~dp0

:: Check if EncodingFixer.ps1 exists
if not exist "%CURRENT_DIR%EncodingFixer.ps1" (
    echo Encoding fix tool not found.
    echo Running CreateFixerScript.ps1 to create the fix tool...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CURRENT_DIR%CreateFixerScript.ps1"
    echo.
    if not exist "%CURRENT_DIR%EncodingFixer.ps1" (
        echo Failed to create the fix tool.
        echo Please reconfigure manually.
        pause
        exit /b 1
    )
)


:: Run PowerShell script directly with error messages suppressed
echo Running PowerShell script...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { . '%CURRENT_DIR%EncodingFixer.ps1' }" 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo An error occurred. Error code: %ERRORLEVEL%
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Process completed.
echo.

pause
exit /b 0
