@echo off
setlocal enabledelayedexpansion

:: タイトル表示
REM chcp 65001 > nul
cls
echo =====================================================
echo   PowerShellスクリプト文字化け修正ツール
echo =====================================================
echo.
echo 文字化け修正用GUIツールを起動しています...
echo.

:: GUIツールのパス
set GUI_TOOL=%~dp0CharacterEncodingFixer.ps1

:: スクリプト存在チェック
if not exist "%GUI_TOOL%" (
    echo エラー: 文字化け修正ツールが見つかりません:
    echo %GUI_TOOL%
    echo.
    echo 終了するには何かキーを押してください...
    pause >nul
    exit /b 1
)

:: PowerShellの実行ポリシーを確認
powershell -Command "Get-ExecutionPolicy" > "%TEMP%\pspolicy.txt"
set /p PS_POLICY=<"%TEMP%\pspolicy.txt"
del "%TEMP%\pspolicy.txt"

:: 必要に応じて実行ポリシーを一時的に変更
if /i "%PS_POLICY%"=="Restricted" (
    echo 実行ポリシーを一時的に変更します...
    powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force"
)

:: GUIツール実行（管理者権限でも正しく日本語表示されるように設定）
echo PowerShellスクリプトを実行中...
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '%GUI_TOOL%' }"

endlocal
