@echo off
chcp 65001 > nul
setlocal

echo =====================================================
echo   PowerShellスクリプト文字化け修正ツール
echo =====================================================
echo.
echo スクリプトのエンコーディングをUTF-8に修正します
echo.

:: 現在のディレクトリを取得
set CURRENT_DIR=%~dp0

:: EncodingFixer.ps1が存在するか確認
if not exist "%CURRENT_DIR%EncodingFixer.ps1" (
    echo エンコーディング修正ツールが見つかりません。
    echo CreateFixerScript.ps1を実行して修正ツールを作成します...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CURRENT_DIR%CreateFixerScript.ps1"
    echo.
    if not exist "%CURRENT_DIR%EncodingFixer.ps1" (
        echo 修正ツールの作成に失敗しました。
        echo 手動で再設定してください。
        pause
        exit /b 1
    )
)

:: PowerShellスクリプトを実行
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CURRENT_DIR%EncodingFixer.ps1"

echo.
echo 処理が完了しました。
echo.

exit /b 0
