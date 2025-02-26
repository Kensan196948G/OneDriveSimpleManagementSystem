@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

echo =====================================================
echo   OneDriveツール緊急修復ユーティリティ
echo =====================================================
echo.
echo このツールはOneDrive運用ツールの緊急修復を実行します。
echo システムが実行エラーを起こしている場合に使用してください。
echo.

REM 管理者権限チェック
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 警告: このスクリプトは管理者権限なしで実行されています。
    echo 一部の操作が実行できない可能性があります。
    echo 管理者として再実行することを推奨します。
    echo.
    pause
)

REM 現在のディレクトリとプロジェクトルートを設定
set TOOL_DIR=%~dp0
cd /d %TOOL_DIR%
cd ..
set PROJECT_ROOT=%CD%

echo 作業ディレクトリ: %PROJECT_ROOT%
echo.

echo ステップ 1: PowerShellの実行ポリシーを確認・設定中...
powershell -Command "Get-ExecutionPolicy" > %TEMP%\ExecPolicy.txt
set /p EXEC_POLICY=<%TEMP%\ExecPolicy.txt
del %TEMP%\ExecPolicy.txt
echo 現在の実行ポリシー: %EXEC_POLICY%

if "%EXEC_POLICY%"=="Restricted" (
    echo 実行ポリシーを一時的に変更します...
    powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force"
    echo 実行ポリシーを変更しました。
)

echo.
echo ステップ 2: UTF-8エンコードを確認しています...
powershell -Command "[Console]::OutputEncoding.CodePage" > %TEMP%\Encoding.txt
set /p ENCODING=<%TEMP%\Encoding.txt
del %TEMP%\Encoding.txt

echo 現在のエンコーディング: %ENCODING%
if not "%ENCODING%"=="65001" (
    echo 警告: システムのエンコーディングがUTF-8ではありません。
)

echo.
echo ステップ 3: 必要なフォルダの存在確認・作成...
if not exist "%PROJECT_ROOT%\logs" (
    mkdir "%PROJECT_ROOT%\logs"
    echo logs フォルダを作成しました
)

if not exist "%PROJECT_ROOT%\診断" (
    mkdir "%PROJECT_ROOT%\診断"
    echo 診断 フォルダを作成しました
)

if not exist "%PROJECT_ROOT%\修正ツール" (
    mkdir "%PROJECT_ROOT%\修正ツール"
    echo 修正ツール フォルダを作成しました
)

echo.
echo ステップ 4: スクリプトエンコーディングの修正...

echo [A] 必須スクリプトのエンコーディングを修正しています...
set IMPORTANT_SCRIPTS=OneDriveStatusCheck.ps1

for %%s in (%IMPORTANT_SCRIPTS%) do (
    if exist "%PROJECT_ROOT%\%%s" (
        echo     %%s を修正中...
        powershell -NoProfile -Command "$content = [System.IO.File]::ReadAllText('%PROJECT_ROOT%\%%s', [System.Text.Encoding]::Default); [System.IO.File]::WriteAllText('%PROJECT_ROOT%\%%s', $content, [System.Text.Encoding]::UTF8)"
        echo     %%s を修正しました
    ) else (
        echo     警告: %%s が見つかりません
    )
)

echo [B] バッチファイルのエンコーディングを修正しています...
for %%f in ("%PROJECT_ROOT%\*.bat") do (
    echo     %%~nxf を修正中...
    powershell -NoProfile -Command "$content = [System.IO.File]::ReadAllText('%%f', [System.Text.Encoding]::Default); [System.IO.File]::WriteAllText('%%f', $content, [System.Text.Encoding]::UTF8)"
)

echo.
echo ステップ 5: エンコーディング修正ツールの準備...
echo CreateFixerScript.ps1を実行して修正ツールを生成しています...
powershell -NoProfile -ExecutionPolicy Bypass -File "%TOOL_DIR%CreateFixerScript.ps1"

echo.
echo ステップ 6: PowerShellモジュールの状態を確認...
powershell -NoProfile -Command "Get-Module -ListAvailable Microsoft.Graph.* | Format-Table Name, Version"

echo.
echo =====================================================
echo   緊急修復処理が完了しました
echo =====================================================
echo.
echo 問題が解決しない場合は以下を試してください:
echo.
echo 1. EncodingFixer.batを実行してスクリプトを修正
echo 2. PowerShellを管理者として実行し、以下のコマンドを実行:
echo    Install-Module Microsoft.Graph.Users -Force
echo    Install-Module Microsoft.Graph.Files -Force
echo    Install-Module Microsoft.Graph.Authentication -Force
echo.
echo 3. OneDriveStatusCheck.ps1の実行時に以下のパラメータを追加:
echo    -ClearAuth
echo.

pause