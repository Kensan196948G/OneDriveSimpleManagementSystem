@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

:: タイトル設定
title OneDrive検索ツール

:: 画面クリア
cls
echo OneDrive検索ツール
echo =====================================
echo.

:: 検索文字列の入力受付
set /p SEARCH_TERM="検索したい文字列を入力してください: "

:: 検索先フォルダの設定
set ONEDRIVE_PATH=%USERPROFILE%\OneDrive

if not exist "%ONEDRIVE_PATH%" (
    echo OneDriveフォルダが見つかりません: %ONEDRIVE_PATH%
    goto :END
)

echo.
echo "%SEARCH_TERM%"を検索中...お待ちください...
echo.

:: 検索実行
findstr /s /i /n /p "%SEARCH_TERM%" "%ONEDRIVE_PATH%\*.*" > "%TEMP%\onedrive_search_results.txt" 2>nul

:: 結果表示
echo 検索結果:
echo =====================================
type "%TEMP%\onedrive_search_results.txt"
echo =====================================
echo.
echo 検索が完了しました。
echo 結果は %TEMP%\onedrive_search_results.txt にも保存されています。

:END
pause
endlocal
