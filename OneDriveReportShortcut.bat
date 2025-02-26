@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

:: OneDrive運用ツールのメインバッチファイル
:: このファイルは各種機能へのショートカットを提供します

title OneDrive運用ツール メニュー

:menu
cls
echo =====================================================
echo   OneDrive運用ツール - メインメニュー
echo =====================================================
echo.
echo  1. OneDriveステータスレポート作成
echo  2. OneDrive接続診断ツール実行
echo  3. 文字化け修正ツール起動
echo  4. PowerShellエンコーディング設定
echo  5. 認証キャッシュをクリア
echo  0. 終了
echo.
echo =====================================================
echo.

set /p choice="選択してください (0-5): "

if "%choice%"=="1" goto run_report
if "%choice%"=="2" goto run_diagnostic
if "%choice%"=="3" goto run_encoding_fixer
if "%choice%"=="4" goto run_ps_encoding
if "%choice%"=="5" goto clear_auth
if "%choice%"=="0" goto end

echo 無効な選択です。もう一度お試しください。
timeout /t 2 > nul
goto menu

:run_report
cls
echo OneDriveステータスレポートを作成します...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1"
pause
goto menu

:run_diagnostic
cls
echo OneDrive接続診断ツールを実行します...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0診断\OneDriveDebugLog.ps1"
pause
goto menu

:run_encoding_fixer
cls
echo 文字化け修正ツールを起動します...
echo.
call "%~dp0修正ツール\EncodingFixer.bat"
goto menu

:run_ps_encoding
cls
echo PowerShellエンコーディング設定を実行します...
echo.
call "%~dp0PowerShellEncoding.bat"
goto menu

:clear_auth
cls
echo 認証キャッシュをクリアしています...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Disconnect-MgGraph -ErrorAction SilentlyContinue; Write-Host '認証キャッシュをクリアしました。' -ForegroundColor Green"
echo.
echo OneDriveステータスレポートを実行しますか？(Y/N)
set /p run_report_choice=
if /i "%run_report_choice%"=="Y" (
    echo.
    echo 認証キャッシュクリア後にレポートを実行します...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1" -ClearAuth
)
pause
goto menu

:end
echo.
echo OneDrive運用ツールを終了します。
echo.
timeout /t 2 > nul
exit /b 0
