@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

:: OneDrive運用ツールのメインバッチファイル
:: このファイルは各種機能へのショートカットを提供します

title OneDrive Tool Menu

:menu
cls
echo =====================================================
echo   OneDrive Management Tool - Main Menu
echo   OneDrive運用ツール - メインメニュー
echo =====================================================
echo.
echo  1. Create OneDrive Status Report
echo     (OneDriveステータスレポート作成)
echo  2. Run OneDrive Diagnostic Tool
echo     (OneDrive接続診断ツール実行)
echo  3. Launch Encoding Fix Tool
echo     (文字化け修正ツール起動)
echo  0. Exit
echo     (終了)
echo.
echo =====================================================
echo.

set /p choice="Select an option / 選択してください (0-3): "

if "%choice%"=="1" goto run_report
if "%choice%"=="2" goto run_diagnostic
if "%choice%"=="3" goto run_encoding_fixer
if "%choice%"=="0" goto end

echo Invalid selection. Please try again.
echo 無効な選択です。もう一度お試しください。
timeout /t 2 > nul
goto menu

:run_report
cls
echo Creating OneDrive Status Report...
echo OneDriveステータスレポートを作成します...
echo.
if exist "%~dp0OneDriveStatusCheck.ps1" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1"
) else (
    echo =====================================================
    echo   OneDrive Status Report
    echo   OneDriveステータスレポート
    echo =====================================================
    echo.
    echo This tool displays the OneDrive usage status.
    echo このツールはOneDriveの使用状況を表示します。
    echo.
    echo [Note] This is a placeholder file. Please implement the actual OneDrive reporting logic.
    echo [注意] これはプレースホルダーファイルです。実際のOneDriveレポート機能を実装してください。
    echo.
    echo =====================================================
    echo.
    pause
)
pause
goto menu

:run_diagnostic
cls
echo Running OneDrive Diagnostic Tool...
echo OneDrive接続診断ツールを実行します...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0診断\OneDriveDebugLog.ps1"
pause
goto menu

:run_encoding_fixer
cls
echo Launching Encoding Fix Tool...
echo 文字化け修正ツールを起動します...
echo.
if exist "%~dp0修正ツール\EncodingFixer.bat" (
    call "%~dp0修正ツール\EncodingFixer.bat"
) else (
    echo EncodingFixer.bat not found.
    echo EncodingFixer.batが見つかりません。
    echo Please check the repair tools directory.
    echo 修正ツールディレクトリを確認してください。
    pause
)
goto menu

:end
echo.
echo Exiting OneDrive Management Tool.
echo OneDrive運用ツールを終了します。
echo.
timeout /t 2 > nul
exit /b 0