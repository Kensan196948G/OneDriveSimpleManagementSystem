@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion

title OneDrive運用ツール デバッグ

echo =====================================================
echo   OneDrive Management Tool Debug
echo   OneDrive運用ツール デバッグ
echo =====================================================
echo.

set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%

echo 現在のコードページ（文字コード）:
chcp

echo.
echo システム情報:
systeminfo | findstr /B /C:"OS名" /C:"OSバージョン" /C:"システムの種類" /C:"システムロケール"

echo.
echo ディレクトリ内容:
dir

echo.
echo フォルダ作成テスト:
echo テスト用フォルダを作成します...
mkdir "テストフォルダ_932" 2>nul
if %ERRORLEVEL% neq 0 (
    echo フォルダ作成に失敗しました。
) else (
    echo フォルダが正常に作成されました。
)

echo.
echo UTF-8でのフォルダ作成テスト:
chcp 65001 > nul
mkdir "テストフォルダ_UTF8" 2>nul
if %ERRORLEVEL% neq 0 (
    echo UTF-8でのフォルダ作成に失敗しました。
) else (
    echo UTF-8でフォルダが正常に作成されました。
)

chcp 932 > nul
echo.
echo PowerShell での文字エンコード情報:
powershell -Command "[Console]::OutputEncoding; [Console]::InputEncoding; [System.Text.Encoding]::Default"

echo.
echo =====================================================
echo デバッグ情報の確認が完了しました。
echo この情報をIT担当者に提供してください。
echo =====================================================
echo.
pause
