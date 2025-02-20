@echo off
REM filepath: /c:/kitting/OneDrive運用ツール/OneDriveStatus.bat


setlocal enabledelayedexpansion

REM 文字化け防止のためのコードページ設定
chcp 65001 > nul

REM 管理者権限チェックと昇格
NET FILE 1>NUL 2>NUL
if not '%errorlevel%' == '0' (
    powershell -Command "Start-Process '%~f0' -Verb RunAs -Wait"
    exit /b
)

REM PowerShellを使用して日本語メッセージを表示
powershell -Command "Write-Host '● PowerShell実行ポリシーを確認しています...' -ForegroundColor Yellow"

REM 実行ポリシーの確認
powershell -Command "Get-ExecutionPolicy" > "%TEMP%\psexecpolicy.txt"
set /p CURRENT_POLICY=<"%TEMP%\psexecpolicy.txt"
del "%TEMP%\psexecpolicy.txt"

powershell -Command "Write-Host ('● 現在のポリシー: ' + '%CURRENT_POLICY%') -ForegroundColor Yellow"
powershell -Command "Write-Host '● PowerShellスクリプトを実行します...' -ForegroundColor Yellow"

REM スクリプトのパスを設定
set "PS_SCRIPT=%~dp0OneDriveStatusCheckLatestVersion20250220.ps1"

REM PowerShellスクリプトを実行（コマンドを1行にまとめ、ReadKey の引数をダブルクォートで指定）
powershell -NoProfile -NoExit -ExecutionPolicy Bypass -NoProfile -Command "$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '%PS_SCRIPT%'; Write-Host '● スクリプトの実行が完了しました。Enterキーを押して終了してください。' -ForegroundColor Green; $null = $Host.UI.RawUI.ReadKey(\"NoEcho,IncludeKeyDown\")"

endlocal
