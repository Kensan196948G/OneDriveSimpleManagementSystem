@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion

title OneDrive運用ツール フォルダ修正

echo =====================================================
echo   OneDrive Management Tool Folder Fix
echo   OneDrive運用ツール フォルダ修正
echo =====================================================
echo.

set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%

echo 現在のディレクトリ内容:
dir

echo.
echo 文字化けしたフォルダを修正します。
echo 注意: この処理は既存のフォルダをリネームするため、必要に応じてバックアップを取ってください。
echo.

set /p CONTINUE="続行しますか？ (Y/N): "
if /i not "%CONTINUE%"=="Y" (
    echo 処理を中止しました。
    goto :end
)

echo.
echo 標準フォルダを作成します...

if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"
if not exist "%SCRIPT_DIR%診断" mkdir "%SCRIPT_DIR%診断"
if not exist "%SCRIPT_DIR%修正ツール" mkdir "%SCRIPT_DIR%修正ツール"

echo.
echo 文字化けしたフォルダ内のファイルを移動し、不要なフォルダを削除します...

for /f "delims=" %%a in ('dir /b /ad') do (
    echo 処理: %%a
    if not "%%a"=="logs" if not "%%a"=="診断" if not "%%a"=="修正ツール" (
        if exist "%%a\*" (
            echo   フォルダ %%a の内容をチェック中...
            dir "%%a" > nul 2>&1
            if !ERRORLEVEL! equ 0 (
                echo   ファイルを適切なフォルダに移動します...
                for /f "delims=" %%f in ('dir /b "%%a\*.*" 2^>nul') do (
                    if "%%~xf"==".log" (
                        move "%%a\%%f" "%SCRIPT_DIR%logs\" > nul
                        echo   %%f を logs フォルダに移動しました。
                    ) else (
                        move "%%a\%%f" "%SCRIPT_DIR%\" > nul
                        echo   %%f をルートフォルダに移動しました。
                    )
                )
            )
            rd "%%a" 2>nul
            if !ERRORLEVEL! equ 0 (
                echo   フォルダ %%a を削除しました。
            ) else (
                echo   フォルダ %%a の削除に失敗しました。手動での確認が必要です。
            )
        )
    )
)

echo.
echo 処理が完了しました。

:end
echo.
echo =====================================================
echo.
pause
