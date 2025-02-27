@echo off
setlocal enabledelayedexpansion
chcp 932 > nul

REM PowerShellスクリプトを管理者権限で実行するためのランチャー
title OneDrive運用ツール 管理者実行ランチャー
color 17

REM 管理者権限チェック
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ==========================================
    echo   このツールは管理者権限で実行する必要があります。
    echo   右クリックから「管理者として実行」を選択してください。
    echo ==========================================
    echo.
    echo 終了するには何かキーを押してください...
    pause >nul
    exit /b 1
)

:menu
cls
echo ==========================================
echo   OneDrive運用ツール 管理者実行ランチャー
echo ==========================================
echo.
echo 管理者権限で実行するスクリプトを選択してください:
echo.
echo [1] OneDriveステータスレポート作成
echo [2] OneDrive接続診断ツール実行
echo [3] 文字化け修正ツール起動
echo [0] 終了
echo.
echo ==========================================
echo.

set /p OPTION=選択してください (0-3): 

if "%OPTION%"=="1" goto :run_report
if "%OPTION%"=="2" goto :run_debug 
if "%OPTION%"=="3" goto :run_encoding_fix
if "%OPTION%"=="0" goto :end

echo 無効な選択です。
timeout /t 2 >nul
goto :menu

:run_report
cls
echo 管理者権限で OneDriveステータスレポート作成を実行します...
echo.

set SCRIPT_PATH=%~dp0OneDriveStatusCheck.ps1

REM 明示的にUTF-8エンコーディングを設定してスクリプト実行
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%SCRIPT_PATH%'"

echo.
echo レポート作成が完了しました。
echo.
pause
goto :menu

:run_debug
cls
echo 管理者権限で OneDrive接続診断ツールを実行します...
echo.

set DEBUG_SCRIPT_PATH=%~dp0OneDriveDebugLog.ps1

REM 明示的にUTF-8エンコーディングを設定してスクリプト実行
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%DEBUG_SCRIPT_PATH%'"

echo.
echo 診断が完了しました。
echo.
pause
goto :menu

:run_encoding_fix
cls
echo 管理者権限で 文字化け修正ツールを起動します...
echo.

REM 文字化け修正ツールも明示的にエンコーディングを指定して起動
set GUI_TOOL=%~dp0CharacterEncodingFixer.ps1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%GUI_TOOL%'"

echo.
echo 文字化け修正ツールを終了しました。
echo.
pause
goto :menu

:end
echo.
echo ツールを終了します...
timeout /t 2 >nul
endlocal
exit /b 0
