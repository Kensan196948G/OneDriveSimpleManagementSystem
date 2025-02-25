@echo off
REM 管理者権限チェックと昇格
NET FILE 1>NUL 2>NUL
if not '%errorlevel%' == '0' (
    powershell -Command "Start-Process '%~f0' -Verb RunAs -Wait"
    exit /b
)

REM PowerShellスクリプトのパスを設定
set "PS_SCRIPT=%~dp0OneDriveStatusCheckLatestVersion20250225.ps1"

REM PowerShellスクリプトを直接実行（シンプルな実行方法に変更）
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

REM スクリプト完了メッセージ
echo.
echo ● スクリプトの実行が完了しました。Enterキーを押して終了してください。
pause > nul
