@echo off
setlocal enabledelayedexpansion

REM OneDriveレポート作成ショートカット
REM このバッチファイルはPowerShellスクリプトを起動するためのものです

REM コンソールをUTF-8に設定
chcp 65001 > nul

REM タイトルとコンソール設定
title OneDriveステータスレポート作成ツール
color 1F

REM バナー表示
echo ================================================
echo       OneDrive ステータスレポート作成ツール
echo ================================================
echo.
echo このツールはOneDriveの使用状況レポートを生成します
echo.

REM スクリプトパスの設定
set SCRIPT_PATH=%~dp0OneDriveStatusCheck.ps1
set DEBUG_SCRIPT_PATH=%~dp0OneDriveDebugLog.ps1

REM スクリプト存在チェック
if not exist "%SCRIPT_PATH%" (
    echo エラー: スクリプトファイルが見つかりません:
    echo %SCRIPT_PATH%
    echo.
    echo 終了するには何かキーを押してください...
    pause >nul
    exit /b 1
)

REM メニュー表示
:menu
cls
echo ================================================
echo       OneDrive ステータスレポート作成ツール
echo ================================================
echo.
echo [1] OneDriveステータスレポート作成
echo [2] OneDrive接続診断ツール実行
echo [3] 文字化け修正ツール起動
echo [0] 終了
echo.
echo ================================================
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
echo OneDriveステータスレポート作成を実行します...
echo.

REM PowerShellコマンドを実行（文字化け対策付き）
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '%SCRIPT_PATH%' }"

echo.
echo レポート作成が完了しました。
echo.
pause
goto :menu

:run_debug
cls
echo OneDrive接続診断ツールを実行します...
echo.

REM PowerShellコマンドを実行（文字化け対策付き）
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '%DEBUG_SCRIPT_PATH%' }"

echo.
echo 診断が完了しました。
echo.
pause
goto :menu

:run_encoding_fix
cls
echo 文字化け修正ツールを起動します...
echo.

call "%~dp0文字化け修正.bat"

echo.
echo 文字化け修正ツールを終了しました。
echo.
pause
goto :menu

:end
echo.
echo OneDriveステータスレポート作成ツールを終了します...
timeout /t 2 >nul
endlocal
exit /b 0
