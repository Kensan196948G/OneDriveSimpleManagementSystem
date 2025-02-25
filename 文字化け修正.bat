@echo off
setlocal

:: タイトル表示
cls
echo =====================================================
echo   PowerShellスクリプト文字化け修正ツール
echo =====================================================
echo.
echo 文字化け診断・修正GUIを起動しています...
echo.

:: GUIツールのパス
set GUI_TOOL=%~dp0CharacterEncodingFixer.ps1

:: スクリプト存在チェック
if not exist "%GUI_TOOL%" (
    echo エラー: 文字化け修正ツールが見つかりません:
    echo %GUI_TOOL%
    echo.
    echo 続行するには何かキーを押してください...
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

:: GUIツール実行
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%GUI_TOOL%'"

endlocal
