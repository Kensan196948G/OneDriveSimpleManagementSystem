@echo off
cmd /c chcp 65001 > nul
setlocal

echo =====================================================
echo   OneDrive運用ツール スクリプト一括変換
echo =====================================================
echo.
echo このツールはフォルダ内の全スクリプトをUTF-8(BOMなし)に変換します
echo.

set TOOL_DIR=%~dp0
set PROJECT_DIR=C:\kitting\OneDrive運用ツール

echo 変換を開始します...

REM PowerShellスクリプトファイル
for /R "%PROJECT_DIR%" %%f in (*.ps1 *.psm1) do (
    echo 変換: %%f
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$content = [System.IO.File]::ReadAllText('%%f', [System.Text.Encoding]::Default); $utf8NoBom = New-Object System.Text.UTF8Encoding($false); [System.IO.File]::WriteAllText('%%f', $content, $utf8NoBom)"
)

REM バッチファイル
for /R "%PROJECT_DIR%" %%f in (*.bat *.cmd) do (
    echo 変換: %%f
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$content = [System.IO.File]::ReadAllText('%%f', [System.Text.Encoding]::Default); $utf8NoBom = New-Object System.Text.UTF8Encoding($false); [System.IO.File]::WriteAllText('%%f', $content, $utf8NoBom)"
)

REM マークダウンファイル
for /R "%PROJECT_DIR%" %%f in (*.md *.markdown) do (
    echo 変換: %%f
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$content = [System.IO.File]::ReadAllText('%%f', [System.Text.Encoding]::Default); $utf8NoBom = New-Object System.Text.UTF8Encoding($false); [System.IO.File]::WriteAllText('%%f', $content, $utf8NoBom)"
)

echo.
echo 変換完了！
echo.

echo スクリプト実行時は以下のコマンドを使用してください:
echo - PowerShell: [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
echo - バッチ: cmd /c chcp 65001 ^> nul
echo.

pause
