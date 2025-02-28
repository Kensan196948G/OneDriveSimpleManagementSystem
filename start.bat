@echo off
REM OneDrive運用ツール 起動用バッチファイル
title OneDrive運用ツール

setlocal
set SCRIPT_DIR=%~dp0

cls
echo =====================================================
echo   OneDrive Management Tool Launcher
echo   OneDrive運用ツール ランチャー
echo =====================================================
echo.
echo Starting OneDrive Management Tool...
echo OneDrive運用ツールを起動しています...
echo.

if exist "%SCRIPT_DIR%OneDriveReportShortcut.bat" (
    call "%SCRIPT_DIR%OneDriveReportShortcut.bat"
) else (
    echo エラー: OneDriveReportShortcut.batが見つかりません。
    echo Error: OneDriveReportShortcut.bat not found.
    echo.
    echo セットアップを実行してください。
    echo Please run the setup first.
    echo.
    pause
)

exit /b
