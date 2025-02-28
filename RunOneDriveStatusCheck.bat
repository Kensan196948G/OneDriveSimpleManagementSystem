@echo off
cls
echo ===================================================
echo  OneDrive ステータスチェックツール
echo ===================================================
echo.
echo このツールは、OneDriveの使用状況を取得し、
echo レポートを生成します。
echo.
echo 実行中はウィンドウを閉じないでください...
echo.
echo ※Microsoft認証画面が表示されたら、
echo  自分のアカウントでログインしてください。
echo ===================================================
echo.

cd /d "%~dp0"

echo PowerShellスクリプトを実行しています...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1"

echo.
if %errorlevel% neq 0 (
    echo エラーが発生しました。エラーログを確認してください。
    pause
    exit /b %errorlevel%
)

echo 完了しました！
echo.
echo ウィンドウを閉じてください。
pause
