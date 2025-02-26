@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion

echo =====================================================
echo   OneDrive運用ツール セットアップ
echo =====================================================
echo.

REM 管理者権限のチェック
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo 管理者権限が必要です。
    echo 右クリックして「管理者として実行」でこのバッチファイルを実行してください。
    echo.
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%

echo 1. 必要なフォルダを作成しています...
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"
if not exist "%SCRIPT_DIR%診断" mkdir "%SCRIPT_DIR%診断"
if not exist "%SCRIPT_DIR%修正ツール" mkdir "%SCRIPT_DIR%修正ツール"

echo 2. PowerShellスクリプト実行ポリシーを確認しています...
powershell.exe -Command "Get-ExecutionPolicy -Scope CurrentUser" > %temp%\exepolicy.txt
set /p POLICY=<%temp%\exepolicy.txt
del %temp%\exepolicy.txt

echo 現在の実行ポリシー: %POLICY%
if /i "%POLICY%"=="Restricted" (
    echo 警告: PowerShellの実行ポリシーが制限されています。
    echo スクリプトを実行するには実行ポリシーを変更する必要があります。
    echo.
    
    set /p CHANGE_POLICY="実行ポリシーをRemoteSigned に変更しますか？(Y/N): "
    if /i "!CHANGE_POLICY!"=="Y" (
        echo 実行ポリシーを変更しています...
        powershell.exe -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" >nul 2>&1
        if !ERRORLEVEL! neq 0 (
            echo 実行ポリシーの変更に失敗しました。手動で変更してください。
        ) else (
            echo 実行ポリシーをRemoteSigned に変更しました。
        )
    ) else (
        echo 実行ポリシーは変更されませんでした。スクリプトが実行できない場合は手動で変更してください。
    )
)

echo.
echo 3. 必要なPowerShellモジュールを確認しています...
powershell.exe -Command "if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { Write-Host 'Microsoft.Graph.Authentication モジュールがインストールされていません。インストールを開始します...'; Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force }"
powershell.exe -Command "if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) { Write-Host 'Microsoft.Graph.Users モジュールがインストールされていません。インストールを開始します...'; Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force }"
powershell.exe -Command "if (-not (Get-Module -ListAvailable Microsoft.Graph.Files)) { Write-Host 'Microsoft.Graph.Files モジュールがインストールされていません。インストールを開始します...'; Install-Module Microsoft.Graph.Files -Scope CurrentUser -Force }"

echo.
echo 4. スクリプトファイルのエンコーディングをUTF-8に修正しています...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%修正ツール\EncodingCheck.ps1" -FixAutomatically

echo.
echo 5. OneDrive運用ツールを起動します...
start "" "%SCRIPT_DIR%OneDriveReportShortcut.bat"

echo.
echo セットアップが完了しました！
echo.
echo =====================================================
echo   OneDrive運用ツールの使い方
echo =====================================================
echo 1. OneDriveステータスレポート：OneDriveの使用状況を確認
echo 2. OneDrive接続診断：接続問題のトラブルシューティング
echo 3. 文字化け修正：スクリプト文字化け問題の解決
echo 4. PowerShellエンコード設定：エンコーディング環境設定
echo 5. 認証キャッシュクリア：認証問題の解決
echo.
echo 詳細は MANUAL.md ファイルを参照してください。
echo =====================================================
echo.
pause
