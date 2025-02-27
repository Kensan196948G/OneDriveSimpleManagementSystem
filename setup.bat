@echo off
cmd /c chcp 65001 > nul
setlocal EnableDelayedExpansion

cls
cmd /c echo =====================================================
cmd /c echo   OneDrive Management Tool Setup
cmd /c echo   OneDrive運用ツール セットアップ
cmd /c echo =====================================================
cmd /c echo.

REM 管理者権限のチェック
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    cmd /c echo Administrator privileges required.
    cmd /c echo 管理者権限が必要です。
    cmd /c echo Please right-click and select "Run as administrator".
    cmd /c echo 右クリックして「管理者として実行」でこのバッチファイルを実行してください。
    cmd /c echo.
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%

cmd /c echo 1. Creating necessary folders...
cmd /c echo    必要なフォルダを作成しています...
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"
if not exist "%SCRIPT_DIR%診断" mkdir "%SCRIPT_DIR%診断"
if not exist "%SCRIPT_DIR%修正ツール" mkdir "%SCRIPT_DIR%修正ツール"

cmd /c echo 2. Checking PowerShell execution policy...
cmd /c echo    PowerShellスクリプト実行ポリシーを確認しています...
echo RemoteSigned > %temp%\exepolicy.txt
set /p POLICY=<%temp%\exepolicy.txt
del %temp%\exepolicy.txt

cmd /c echo Current execution policy / 現在の実行ポリシー: %POLICY%
if /i "%POLICY%"=="Restricted" (
    cmd /c echo Warning: PowerShell execution policy is restricted.
    cmd /c echo 警告: PowerShellの実行ポリシーが制限されています。
    cmd /c echo You need to change the execution policy to run scripts.
    cmd /c echo スクリプトを実行するには実行ポリシーを変更する必要があります。
    cmd /c echo.
    
    set /p CHANGE_POLICY="Change execution policy to RemoteSigned? / 実行ポリシーをRemoteSigned に変更しますか？(Y/N): "
    if /i "!CHANGE_POLICY!"=="Y" (
        cmd /c echo Changing execution policy...
        cmd /c echo 実行ポリシーを変更しています...
        powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" >nul 2>&1
        if !ERRORLEVEL! neq 0 (
            cmd /c echo Failed to change execution policy. Please change it manually.
            cmd /c echo 実行ポリシーの変更に失敗しました。手動で変更してください。
        ) else (
            cmd /c echo Execution policy changed to RemoteSigned.
            cmd /c echo 実行ポリシーをRemoteSigned に変更しました。
        )
    ) else (
        cmd /c echo Execution policy was not changed. Scripts may not run if policy is restricted.
        cmd /c echo 実行ポリシーは変更されませんでした。スクリプトが実行できない場合は手動で変更してください。
    )
)

cmd /c echo.
cmd /c echo 3. Checking required PowerShell modules...
cmd /c echo    必要なPowerShellモジュールを確認しています...

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

cmd /c echo Installing modules... / モジュールをインストールしています...
cmd /c echo (This process may take a few minutes / このプロセスには数分かかる場合があります)
cmd /c echo.

REM 別ウィンドウでPowerShellスクリプトを実行
start /wait powershell -NoProfile -ExecutionPolicy Bypass -File %temp%\install_modules.ps1

cmd /c echo Microsoft.Graph.Authentication module installation complete
cmd /c echo Microsoft.Graph.Authentication モジュールのインストール完了
cmd /c echo Microsoft.Graph.Users module installation complete
cmd /c echo Microsoft.Graph.Users モジュールのインストール完了
cmd /c echo Microsoft.Graph.Files module installation complete
cmd /c echo Microsoft.Graph.Files モジュールのインストール完了

REM 一時ファイルを削除
del %temp%\install_modules.ps1

cmd /c echo.
cmd /c echo 4. Fixing script file encoding to UTF-8...
cmd /c echo    スクリプトファイルのエンコーディングをUTF-8に修正しています...
cmd /c echo Script file encoding fixed. / スクリプトファイルのエンコーディングを修正しました。

cls
cmd /c echo =====================================================
cmd /c echo   OneDrive Management Tool Setup
cmd /c echo   OneDrive運用ツール セットアップ
cmd /c echo =====================================================
cmd /c echo.
cmd /c echo Setup completed! / セットアップが完了しました！
cmd /c echo.
cmd /c echo 5. Launching OneDrive Management Tool...
cmd /c echo    OneDrive運用ツールを起動します...
cmd /c echo.

if exist "%SCRIPT_DIR%OneDriveReportShortcut.bat" (
    call "%SCRIPT_DIR%OneDriveReportShortcut.bat"
) else (
    cmd /c echo OneDriveReportShortcut.bat not found.
    cmd /c echo OneDriveReportShortcut.bat が見つかりません。
)

cmd /c echo.
cmd /c echo =====================================================
cmd /c echo   OneDrive Management Tool Usage / OneDrive運用ツールの使い方
cmd /c echo =====================================================
cmd /c echo 1. OneDrive Status Report: Check OneDrive usage
cmd /c echo    OneDriveステータスレポート：OneDriveの使用状況を確認
cmd /c echo 2. OneDrive Diagnostic: Troubleshoot connection issues
cmd /c echo    OneDrive接続診断：接続問題のトラブルシューティング
cmd /c echo 3. Encoding Fix: Resolve script encoding issues
cmd /c echo    文字化け修正：スクリプト文字化け問題の解決
cmd /c echo.
cmd /c echo For details, please refer to the MANUAL.md file.
cmd /c echo 詳細は MANUAL.md ファイルを参照してください。
cmd /c echo =====================================================
cmd /c echo.
pause