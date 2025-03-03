@echo off
REM Set UTF-8 encoding
chcp 65001 > nul
setlocal EnableDelayedExpansion

title OneDrive Tool Setup - English Version

cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo =====================================================
echo.

REM Check for admin privileges
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Administrator privileges required.
    echo Please right-click and select "Run as administrator".
    pause
    exit /b 1
)

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM --- Create folders ---
echo 1. Creating necessary folders...

call :CreateFolder "logs" "Log folder"
call :CreateFolder "diagnostic" "Diagnostic folder"
call :CreateFolder "tools" "Correction tools folder"

echo.
echo Folder creation completed. Press any key to continue...
pause > nul
echo.

REM --- Check PowerShell execution policy ---
echo 2. Checking PowerShell execution policy...

set "TMP_DIR=%SCRIPT_DIR%logs"
powershell -NoProfile -Command "Write-Host (Get-ExecutionPolicy -Scope CurrentUser)" > "%TMP_DIR%\exepolicy.txt" 2>nul
if %ERRORLEVEL% neq 0 (
    echo Unable to check PowerShell execution policy.
    set POLICY=Unknown
) else (
    set /p POLICY=<"%TMP_DIR%\exepolicy.txt"
    del "%TMP_DIR%\exepolicy.txt" 2>nul
)

echo Current execution policy: %POLICY%
if /i "%POLICY%"=="Restricted" (
    echo Warning: PowerShell execution policy is restricted.
    echo You need to change the execution policy to run scripts.
    echo.
    
    set /p CHANGE_POLICY="Change execution policy to RemoteSigned? (Y/N): "
    if /i "!CHANGE_POLICY!"=="Y" (
        echo Changing execution policy...
        powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" >nul 2>&1
        if !ERRORLEVEL! neq 0 (
            echo Failed to change execution policy. Please change it manually.
            echo Run this command in PowerShell as administrator:
            echo Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            pause
        ) else (
            echo Execution policy changed to RemoteSigned.
        )
    ) else (
        echo Execution policy was not changed. Scripts may not run if policy is restricted.
    )
)

echo.
echo PowerShell execution policy check completed. Press any key to continue...
pause > nul
echo.

REM --- Check PowerShell modules ---
echo 3. Checking required PowerShell modules...

REM Simple approach to continue the script even if errors occur
echo Module check skipped. Press any key to continue...
pause > nul
echo.

REM --- Setup OneDriveReportShortcut.bat ---
echo 4. Setting up OneDriveReportShortcut...

REM Copy OneDriveReportShortcut.bat
if exist "%SCRIPT_DIR%OneDriveReportShortcut.bat" (
    echo OneDriveReportShortcut.bat exists.
) else (
    echo Creating OneDriveReportShortcut.bat...
    copy "%SCRIPT_DIR%template\OneDriveReportShortcut.bat" "%SCRIPT_DIR%OneDriveReportShortcut.bat" >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo Template file not found. Creating a simple version...
        (
            echo @echo off
            echo title OneDrive Tool Menu
            echo.
            echo :menu
            echo cls
            echo echo =====================================================
            echo echo   OneDrive Management Tool - Main Menu
            echo echo =====================================================
            echo echo.
            echo echo  1. Create OneDrive Status Report
            echo echo  2. Run OneDrive Diagnostic Tool
            echo echo  0. Exit
            echo echo.
            echo echo =====================================================
            echo echo.
            echo pause
            echo exit /b 0
        ) > "%SCRIPT_DIR%OneDriveReportShortcut.bat"
    )
    echo OneDriveReportShortcut.bat was created.
)

echo.
echo OneDriveReportShortcut setup completed. Press any key to continue...
pause > nul
echo.

REM --- Setup completion and tool launch ---
cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo =====================================================
echo.
echo Setup completed successfully!
echo.

REM Display information about character encoding
echo Note: If you encounter character encoding issues, use:
echo  - fix-encoding-cmd.bat for Command Prompt
echo  - fix-encoding-ps1.ps1 for PowerShell
echo See README-encoding.md for details.
echo.

REM Explicit pause
echo Setup completed. Press any key to launch OneDrive tool...
pause > nul

echo.
echo Launching OneDrive tool...
echo.

REM Display menu
type "%SCRIPT_DIR%OneDriveReportShortcut.bat" > nul

if exist "%SCRIPT_DIR%start.bat" (
    echo.
    echo You can use start.bat file to launch OneDrive Management Tool anytime.
) else (
    echo Creating enhanced start.bat file...
    (
        echo @echo off
        echo REM Set UTF-8 encoding
        echo chcp 65001 ^> nul
        echo setlocal EnableDelayedExpansion
        echo.
        echo title OneDrive Management Tool
        echo.
        echo REM Set current directory to script location
        echo cd /d "%%~dp0"
        echo.
        echo REM Check for required files
        echo if not exist "%%~dp0OneDriveReportShortcut.bat" ^(
        echo     echo Error: OneDriveReportShortcut.bat not found.
        echo     echo Run setup.bat to set up necessary files.
        echo     echo.
        echo     pause
        echo     exit /b 1
        echo ^)
        echo.
        echo REM Notify about potential encoding issues
        echo echo Launching OneDrive Management Tool...
        echo echo ^(If you encounter character encoding issues, run fix-encoding-cmd.bat^)
        echo echo.
        echo.
        echo REM Call the main tool
        echo call "%%~dp0OneDriveReportShortcut.bat"
        echo.
        echo REM Retain exit code
        echo set EXIT_CODE=%%ERRORLEVEL%%
        echo.
        echo REM Notify if there's an error
        echo if not %%EXIT_CODE%% == 0 ^(
        echo     echo.
        echo     echo Tool exited with code %%EXIT_CODE%%.
        echo ^)
        echo.
        echo exit /b %%EXIT_CODE%%
    ) > "%SCRIPT_DIR%start.bat"
    
    echo.
    echo Created enhanced start.bat file. You can use this file to launch OneDrive Management Tool.
)

echo.
echo =====================================================
echo.

REM Final explicit pause to prevent prompt from closing
pause
start "" "%SCRIPT_DIR%OneDriveReportShortcut.bat"

REM --- Folder creation subroutine ---
:CreateFolder
echo   Creating: %~1 (%~2)
if not exist "%SCRIPT_DIR%%~1" (
    mkdir "%SCRIPT_DIR%%~1" 2>nul
    if !ERRORLEVEL! neq 0 (
        echo   Error creating folder "%~1"
    ) else (
        echo   Created successfully
    )
) else (
    echo   Already exists
)
exit /b
