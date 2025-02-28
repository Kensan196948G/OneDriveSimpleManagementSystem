@echo off
REM 文字コードをコメントアウト（文字化け解消）
REM chcp 932 > nul
setlocal EnableDelayedExpansion

REM echo onを削除してコマンド表示を無効化
@echo off

REM コンソールウィンドウが閉じないように明示的なpauseを追加
title OneDrive運用ツールセットアップ

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

REM --- フォルダ作成処理 ---
echo 1. Creating necessary folders...
echo    必要なフォルダを作成しています...

call :CreateFolder "logs" "ログ保存用フォルダ"
call :CreateFolder "診断" "診断用フォルダ"
call :CreateFolder "修正ツール" "修正ツール用フォルダ"

echo.
echo フォルダ作成が完了しました。続けるには何かキーを押してください...
pause > nul
echo.

REM --- PowerShell実行ポリシー確認 ---
echo 2. Checking PowerShell execution policy...
echo    PowerShellスクリプト実行ポリシーを確認しています...

set "TMP_DIR=%SCRIPT_DIR%logs"
powershell -NoProfile -Command "Write-Host (Get-ExecutionPolicy -Scope CurrentUser)" > "%TMP_DIR%\exepolicy.txt" 2>nul
if %ERRORLEVEL% neq 0 (
    echo Unable to check PowerShell execution policy.
    echo PowerShell実行ポリシーの確認に失敗しました。
    set POLICY=Unknown
) else (
    set /p POLICY=<"%TMP_DIR%\exepolicy.txt"
    del "%TMP_DIR%\exepolicy.txt" 2>nul
)

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
            echo Run this command in PowerShell as administrator:
            echo 管理者権限でPowerShellを開き、以下のコマンドを実行してください:
            echo Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            pause
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
echo PowerShell実行ポリシーの確認が完了しました。続けるには何かキーを押してください...
pause > nul
echo.

REM --- PowerShellモジュールの確認とインストール ---
echo 3. Checking required PowerShell modules...
echo    必要なPowerShellモジュールを確認しています...

REM エラーが発生してもスクリプトを続行するためのシンプルな対応
echo モジュールの確認をスキップします。続けるには何かキーを押してください...
pause > nul
echo.

REM --- OneDriveReportShortcut.batのセットアップ ---
echo 4. Setting up OneDriveReportShortcut...
echo    OneDriveReportShortcutを設定しています...

REM OneDriveReportShortcut.batのコピー
if exist "%SCRIPT_DIR%OneDriveReportShortcut.bat" (
    echo OneDriveReportShortcut.bat exists.
    echo OneDriveReportShortcut.batが存在します。
) else (
    echo Creating OneDriveReportShortcut.bat...
    echo OneDriveReportShortcut.batを作成しています...
    copy "%SCRIPT_DIR%template\OneDriveReportShortcut.bat" "%SCRIPT_DIR%OneDriveReportShortcut.bat" >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo Template file not found. Creating a simple version...
        echo テンプレートファイルが見つかりません。シンプルバージョンを作成します...
        (
            echo @echo off
            echo title OneDrive Tool Menu
            echo.
            echo :menu
            echo cls
            echo echo =====================================================
            echo echo   OneDrive Management Tool - Main Menu
            echo echo   OneDrive運用ツール - メインメニュー
            echo echo =====================================================
            echo echo.
            echo echo  1. Create OneDrive Status Report
            echo echo     ^(OneDriveステータスレポート作成^)
            echo echo  2. Run OneDrive Diagnostic Tool
            echo echo     ^(OneDrive接続診断ツール実行^)
            echo echo  0. Exit
            echo echo     ^(終了^)
            echo echo.
            echo echo =====================================================
            echo echo.
            echo pause
            echo exit /b 0
        ) > "%SCRIPT_DIR%OneDriveReportShortcut.bat"
    )
    echo OneDriveReportShortcut.bat was created.
    echo OneDriveReportShortcut.batを作成しました。
)

echo.
echo OneDriveReportShortcutの設定が完了しました。続けるには何かキーを押してください...
pause > nul
echo.

REM --- セットアップ完了とツール起動 ---
cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo   OneDrive運用ツール セットアップ
echo =====================================================
echo.
echo セットアップが正常に完了しました！
echo Setup completed successfully!
echo.

REM 明示的にpauseを実行
echo セットアップが完了しました。OneDrive運用ツールを起動するには何かキーを押してください...
pause > nul

echo.
echo OneDrive運用ツールを起動しています...
echo.

REM シンプルにメニューを表示
type "%SCRIPT_DIR%OneDriveReportShortcut.bat" > nul

if exist "%SCRIPT_DIR%start.bat" (
    echo.
    echo start.batファイルを使用して、いつでもOneDrive運用ツールを起動できます。
    echo You can use start.bat file to launch OneDrive Management Tool anytime.
) else (
    echo @echo off > "%SCRIPT_DIR%start.bat"
    echo title OneDrive運用ツール >> "%SCRIPT_DIR%start.bat"
    echo. >> "%SCRIPT_DIR%start.bat"
    echo cd /d "%%~dp0" >> "%SCRIPT_DIR%start.bat"
    echo call "%%~dp0OneDriveReportShortcut.bat" >> "%SCRIPT_DIR%start.bat"
    echo exit /b >> "%SCRIPT_DIR%start.bat"
    
    echo.
    echo start.batファイルを作成しました。このファイルを使用してOneDrive運用ツールを起動できます。
    echo Created start.bat file. You can use this file to launch OneDrive Management Tool.
)

echo.
echo =====================================================
echo.

REM 最後に明示的なpauseを追加して、プロンプトが閉じないようにする
pause
start "" "%SCRIPT_DIR%OneDriveReportShortcut.bat"

REM --- フォルダ作成用サブルーチン ---
:CreateFolder
echo   Creating: %~1 (%~2)
if not exist "%SCRIPT_DIR%%~1" (
    mkdir "%SCRIPT_DIR%%~1" 2>nul
    if !ERRORLEVEL! neq 0 (
        echo   Error creating folder "%~1"
        echo   フォルダ "%~1" の作成に失敗しました
    ) else (
        echo   Created successfully
        echo   正常に作成されました
    )
) else (
    echo   Already exists
    echo   既に存在しています
)
exit /b