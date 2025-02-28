# OneDrive管理ツール メインスクリプト
# 文字コード: UTF-8 with BOM

# スクリプトの実行ポリシーを設定
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# モジュールパスを設定
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path -Path $scriptPath -ChildPath "modules"

# モジュールをインポート
Import-Module "$modulePath\StatusReport.psm1" -ErrorAction SilentlyContinue
Import-Module "$modulePath\Diagnostics.psm1" -ErrorAction SilentlyContinue
Import-Module "$modulePath\EncodingFix.psm1" -ErrorAction SilentlyContinue

# メインメニューを表示
function Show-MainMenu {
    Clear-Host
    Write-Host "=====================================================
  OneDrive Management Tool / OneDrive運用ツール
=====================================================" -ForegroundColor Cyan
    
    Write-Host "1: OneDriveステータスレポート - 使用状況確認" -ForegroundColor Yellow
    Write-Host "2: OneDrive接続診断 - 接続問題のトラブルシューティング" -ForegroundColor Yellow
    Write-Host "3: 文字化け修正 - スクリプト文字化け問題の解決" -ForegroundColor Yellow
    Write-Host "q: 終了" -ForegroundColor Yellow
    Write-Host ""
    $choice = Read-Host "オプションを選択してください"
    
    switch ($choice) {
        "1" { Get-OneDriveStatus }
        "2" { Test-OneDriveConnection }
        "3" { Repair-ScriptEncoding }
        "q" { return "exit" }
        default { 
            Write-Host "無効な選択です。再試行してください。" -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    return $choice
}

# メインループ
$exit = $false
while (-not $exit) {
    $result = Show-MainMenu
    if ($result -eq "exit") {
        $exit = $true
    }
}

Write-Host "OneDrive管理ツールを終了します。" -ForegroundColor Green
