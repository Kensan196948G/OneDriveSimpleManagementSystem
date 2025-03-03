@echo off
REM UTF-8エンコーディングを設定
chcp 65001 > nul
setlocal EnableDelayedExpansion

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
echo  0. Exit
echo     (終了)
echo.
echo =====================================================
echo.

set /p CHOICE="Enter your choice / 選択してください (0-2): "

if "%CHOICE%"=="1" goto report
if "%CHOICE%"=="2" goto diagnostic
if "%CHOICE%"=="0" goto end

echo.
echo Invalid choice. Please try again.
echo 無効な選択です。もう一度試してください。
timeout /t 2 > nul
goto menu

:report
cls
echo.
echo Creating OneDrive Status Report...
echo OneDriveステータスレポートを作成しています...
echo.
echo This function is not implemented yet.
echo この機能はまだ実装されていません。
echo.
pause
goto menu

:diagnostic
cls
echo.
echo Running OneDrive Diagnostic Tool...
echo OneDrive診断ツールを実行しています...
echo.
echo This function is not implemented yet.
echo この機能はまだ実装されていません。
echo.
pause
goto menu

:end
echo.
echo Exiting OneDrive Management Tool...
echo OneDrive運用ツールを終了します。
echo.
exit /b 0