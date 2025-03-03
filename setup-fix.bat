@echo off
REM 文字コードをUTF-8に設定し、余計な出力を隠す
chcp 65001 > nul
setlocal EnableDelayedExpansion

title OneDrive運用ツールセットアップ

cls
echo =====================================================
echo   OneDrive Management Tool Setup
echo   OneDrive運用ツール セットアップ
echo =====================================================
echo.

REM メッセージを英語のみに制限するか、引用符で囲むことで問題を回避
