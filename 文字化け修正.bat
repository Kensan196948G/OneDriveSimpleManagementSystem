@echo off
REM 文字化け修正ツールを起動するバッチファイル

title 文字化け診断・修正ツール

echo 文字化け診断・修正ツールを起動しています...
powershell -ExecutionPolicy Bypass -File "%~dp0CharacterEncodingFixer.ps1"

echo.
echo 処理が完了しました。
pause
