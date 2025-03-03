@echo off
REM 文字コードをUTF-8に設定
chcp 65001 > nul
setlocal EnableDelayedExpansion

title OneDrive Tool Setup - Fixed Version

cls
echo =====================================================
echo   OneDrive Management Tool Setup (Fixed Version)
echo =====================================================
echo.

REM 管理者権限のチェック
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Administrator privileges required.
    echo Please run as administrator.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo Creating necessary folders...

if not exist "%SCRIPT_DIR%logs" (
    mkdir "%SCRIPT_DIR%logs" 2>nul
    echo - Created logs folder
) else (
    echo - Logs folder already exists
)

if not exist "%SCRIPT_DIR%diagnostic" (
    mkdir "%SCRIPT_DIR%diagnostic" 2>nul
    echo - Created diagnostic folder
) else (
    echo - Diagnostic folder already exists
)

if not exist "%SCRIPT_DIR%tools" (
    mkdir "%SCRIPT_DIR%tools" 2>nul
    echo - Created tools folder
) else (
    echo - Tools folder already exists
)

echo.
echo Folder creation completed. Press any key to continue...
pause > nul

echo.
echo Creating encoding fix files...

echo @echo off > "%SCRIPT_DIR%fix-encoding-cmd.bat"
echo chcp 65001 >> "%SCRIPT_DIR%fix-encoding-cmd.bat"
echo echo Command Prompt encoding set to UTF-8 (CP 65001). >> "%SCRIPT_DIR%fix-encoding-cmd.bat"
echo echo Current code page: >> "%SCRIPT_DIR%fix-encoding-cmd.bat"
echo chcp >> "%SCRIPT_DIR%fix-encoding-cmd.bat"
echo pause >> "%SCRIPT_DIR%fix-encoding-cmd.bat"

echo # PowerShell encoding fix > "%SCRIPT_DIR%fix-encoding-ps1.ps1"
echo [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 >> "%SCRIPT_DIR%fix-encoding-ps1.ps1"
echo $OutputEncoding = [System.Text.Encoding]::UTF8 >> "%SCRIPT_DIR%fix-encoding-ps1.ps1"
echo Write-Host "PowerShell encoding set to UTF-8" >> "%SCRIPT_DIR%fix-encoding-ps1.ps1"

echo.
echo Setup completed successfully!
echo.
echo Note: If you encounter character encoding issues, use:
echo  - fix-encoding-cmd.bat for Command Prompt
echo  - fix-encoding-ps1.ps1 for PowerShell
echo.
echo Press any key to exit...
pause > nul
exit /b 0
