@echo off
REM OneDrive運用ツールの起動バッチファイル
REM 2025/02/27 更新

echo OneDriveステータス確認ツールを起動しています...
echo 管理者権限が必要です。UAC確認があれば「はい」を選択してください。

REM 管理者権限で実行
powershell -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0OneDriveStatusCheckLatestVersion20250227.ps1""' -Verb RunAs"

echo.
echo プログラムが起動しました。認証画面でサインインしてください。
echo 処理が完了するまでこの画面は閉じないでください。
