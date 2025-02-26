@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

title OneDrive緊急修復ツール

cls
echo =====================================================
echo   OneDrive緊急修復ツール
echo =====================================================
echo.
echo このツールはOneDriveの深刻な問題を修復するためのものです。
echo 実行する前にデータのバックアップを取ることをお勧めします。
echo.
echo 主な機能:
echo 1. OneDriveの再起動
echo 2. OneDrive設定のリセット
echo 3. OneDriveの完全再インストール
echo.
echo =====================================================
echo.

set /p choice="実行する操作を選択してください (1-3、キャンセルは0): "

if "%choice%"=="0" goto end
if "%choice%"=="1" goto restart_onedrive
if "%choice%"=="2" goto reset_settings
if "%choice%"=="3" goto reinstall_onedrive

echo 無効な選択です。もう一度お試しください。
timeout /t 2 > nul
goto end

:restart_onedrive
echo.
echo OneDriveを再起動しています...
echo.
taskkill /f /im OneDrive.exe > nul 2>&1
timeout /t 2 > nul
start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
echo OneDriveを再起動しました。
goto end

:reset_settings
echo.
echo 警告: OneDrive設定をリセットしますか？
echo この操作は再同期が必要になる場合があります。
set /p confirm="続行しますか？ (Y/N): "
if /i not "%confirm%"=="Y" goto end

echo.
echo OneDriveの設定をリセットしています...
echo.
taskkill /f /im OneDrive.exe > nul 2>&1
timeout /t 2 > nul
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive" /f > nul 2>&1
echo 設定がリセットされました。OneDriveを再起動します...
start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
goto end

:reinstall_onedrive
echo.
echo 警告: OneDriveを完全に再インストールしますか？
echo この操作は設定とキャッシュを削除し、再同期が必要になります。
set /p confirm="続行しますか？ (Y/N): "
if /i not "%confirm%"=="Y" goto end

echo.
echo OneDriveを完全に再インストールしています...
echo.
taskkill /f /im OneDrive.exe > nul 2>&1
timeout /t 2 > nul

:: OneDriveをアンインストール
if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
    echo OneDriveをアンインストールしています...
    "%SYSTEMROOT%\System32\OneDriveSetup.exe" /uninstall
) else (
    echo OneDriveのアンインストーラーが見つかりません。
    goto end
)

timeout /t 5 > nul

:: OneDrive設定をクリア
echo OneDrive設定をクリアしています...
rd "%LOCALAPPDATA%\Microsoft\OneDrive" /s /q > nul 2>&1
rd "%PROGRAMDATA%\Microsoft OneDrive" /s /q > nul 2>&1
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive" /f > nul 2>&1

timeout /t 2 > nul

:: OneDriveを再インストール
if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
    echo OneDriveを再インストールしています...
    "%SYSTEMROOT%\System32\OneDriveSetup.exe"
) else (
    echo OneDriveのインストーラーが見つかりません。
    echo Microsoft公式サイトからOneDriveをダウンロードしてインストールしてください。
)

echo 処理が完了しました。
goto end

:end
echo.
echo プログラムを終了します...
echo.
timeout /t 3 > nul
exit /b 0
