@echo off
setlocal

echo ======================================================
echo  OneDriveステータスレポート (認証リセット版)
echo ======================================================
echo.
echo このスクリプトは認証キャッシュをクリアしてから
echo OneDriveステータスレポートを生成します。
echo.
echo 「Access denied」または「アクセス拒否」エラーが
echo 発生した場合に実行してください。
echo.
echo ======================================================
echo.

set POWERSHELL=powershell.exe -NoProfile -ExecutionPolicy Bypass

echo 認証キャッシュをクリアしています...
%POWERSHELL% -Command "Disconnect-MgGraph -ErrorAction SilentlyContinue"

echo.
echo エラー発生時にもユーザーをスキップして処理を継続します
echo.

%POWERSHELL% -File "C:\kitting\OneDrive運用ツール\OneDriveStatusCheck.ps1" -SkipErrorUsers $true -ClearAuth

echo.
echo 処理が完了しました。
pause
