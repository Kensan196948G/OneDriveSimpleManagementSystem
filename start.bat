@echo off
REM 文字コードをUTF-8に設定（文字化け解消）
chcp 65001 > nul
setlocal EnableDelayedExpansion

title OneDrive運用ツール

REM 現在のディレクトリをスクリプトのある場所に設定
cd /d "%~dp0"

REM 必要なファイルの存在確認
if not exist "%~dp0OneDriveReportShortcut.bat" (
    echo エラー: OneDriveReportShortcut.bat が見つかりません。
    echo setup.bat を実行して必要なファイルをセットアップしてください。
    echo.
    pause
    exit /b 1
)

REM 文字化け問題の可能性を通知
echo OneDrive運用ツールを起動しています...
echo （文字化けが発生した場合は fix-encoding-cmd.bat を実行してください）
echo.

REM メインツールを呼び出し
call "%~dp0OneDriveReportShortcut.bat"

REM 終了コードを保持
set EXIT_CODE=%ERRORLEVEL%

REM エラーがあれば通知
if not %EXIT_CODE% == 0 (
    echo.
    echo ツールが終了コード %EXIT_CODE% で終了しました。
)

exit /b %EXIT_CODE%
