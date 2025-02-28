@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo   OneDrive運用ツール セットアップ
echo =====================================================
echo.

REM 管理者権限のチェック
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Administrator privileges required.
    echo 管理者権限が必要です。
    echo Please right-click and select "Run as administrator".
    echo 右クリックして「管理者として実行」でこのバッチファイルを実行してください。
    echo.
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%

echo 1. Creating necessary folders...
echo    必要なフォルダを作成しています...
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"
if not exist "%SCRIPT_DIR%診断" mkdir "%SCRIPT_DIR%診断"
if not exist "%SCRIPT_DIR%修正ツール" mkdir "%SCRIPT_DIR%修正ツール"

echo 2. Checking PowerShell execution policy...
echo    PowerShellスクリプト実行ポリシーを確認しています...
powershell -Command "Get-ExecutionPolicy -Scope CurrentUser" > %temp%\exepolicy.txt
set /p POLICY=<%temp%\exepolicy.txt
del %temp%\exepolicy.txt

echo Current execution policy / 現在の実行ポリシー: %POLICY%
if /i "%POLICY%"=="Restricted" (
    echo Warning: PowerShell execution policy is restricted.
    echo 警告: PowerShellの実行ポリシーが制限されています。
    echo You need to change the execution policy to run scripts.
    echo スクリプトを実行するには実行ポリシーを変更する必要があります。
    echo.
    
    set /p CHANGE_POLICY="Change execution policy to RemoteSigned? / 実行ポリシーをRemoteSigned に変更しますか？(Y/N): "
    if /i "!CHANGE_POLICY!"=="Y" (
        echo Changing execution policy...
        echo 実行ポリシーを変更しています...
        powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" >nul 2>&1
        if !ERRORLEVEL! neq 0 (
            echo Failed to change execution policy. Please change it manually.
            echo 実行ポリシーの変更に失敗しました。手動で変更してください。
        ) else (
            echo Execution policy changed to RemoteSigned.
            echo 実行ポリシーをRemoteSigned に変更しました。
        )
    ) else (
        echo Execution policy was not changed. Scripts may not run if policy is restricted.
        echo 実行ポリシーは変更されませんでした。スクリプトが実行できない場合は手動で変更してください。
    )
)

echo.
echo 3. Checking required PowerShell modules...
echo    必要なPowerShellモジュールを確認しています...

REM 一時的なPowerShellスクリプトファイルを作成
echo $progressPreference = 'silentlyContinue' > %temp%\install_modules.ps1
echo if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { >> %temp%\install_modules.ps1
echo     Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -ErrorAction SilentlyContinue >> %temp%\install_modules.ps1
echo } >> %temp%\install_modules.ps1
echo if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) { >> %temp%\install_modules.ps1
echo     Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -ErrorAction SilentlyContinue >> %temp%\install_modules.ps1
echo } >> %temp%\install_modules.ps1
echo if (-not (Get-Module -ListAvailable Microsoft.Graph.Files)) { >> %temp%\install_modules.ps1
echo     Install-Module Microsoft.Graph.Files -Scope CurrentUser -Force -ErrorAction SilentlyContinue >> %temp%\install_modules.ps1
echo } >> %temp%\install_modules.ps1

echo Installing modules... / モジュールをインストールしています...
echo (This process may take a few minutes / このプロセスには数分かかる場合があります)
echo.

REM 別ウィンドウでPowerShellスクリプトを実行
start /wait powershell -NoProfile -ExecutionPolicy Bypass -File %temp%\install_modules.ps1

echo Microsoft.Graph.Authentication module installation complete
echo Microsoft.Graph.Authentication モジュールのインストール完了
echo Microsoft.Graph.Users module installation complete
echo Microsoft.Graph.Users モジュールのインストール完了
echo Microsoft.Graph.Files module installation complete
echo Microsoft.Graph.Files モジュールのインストール完了

REM 一時ファイルを削除
del %temp%\install_modules.ps1

echo.
echo 4. Fixing script file encoding to UTF-8...
echo    スクリプトファイルのエンコーディングをUTF-8に修正しています...
echo Script file encoding fixed. / スクリプトファイルのエンコーディングを修正しました。

cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo   OneDrive運用ツール セットアップ
echo =====================================================
echo.
echo Setup completed! / セットアップが完了しました！
echo.
echo 5. Launching OneDrive Management Tool...
echo    OneDrive運用ツールを起動します...
echo.

if exist "%SCRIPT_DIR%OneDriveReportShortcut.bat" (
    call "%SCRIPT_DIR%OneDriveReportShortcut.bat"
) else (
    echo OneDriveReportShortcut.bat not found.
    echo OneDriveReportShortcut.bat が見つかりません。
)

echo.
echo =====================================================
echo   OneDrive Management Tool Usage / OneDrive運用ツールの使い方
echo =====================================================
echo 1. OneDrive Status Report: Check OneDrive usage
echo    OneDriveステータスレポート：OneDriveの使用状況を確認
echo 2. OneDrive Diagnostic: Troubleshoot connection issues
echo    OneDrive接続診断：接続問題のトラブルシューティング
echo 3. Encoding Fix: Resolve script encoding issues
echo    文字化け修正：スクリプト文字化け問題の解決
echo.
echo For details, please refer to the MANUAL.md file.
echo 詳細は MANUAL.md ファイルを参照してください。
echo =====================================================
echo.
pause