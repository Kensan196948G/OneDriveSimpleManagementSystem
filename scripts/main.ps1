# OneDrive管理ツール メインスクリプト
# 文字コード: UTF-8 with BOM

# スクリプトのエンコーディングを確認
$scriptContent = Get-Content -Path $MyInvocation.MyCommand.Path -Raw
$encoding = [System.Text.Encoding]::UTF8
$preamble = $encoding.GetPreamble()
$hasBOM = $scriptContent.Length -ge $preamble.Length -and 
          ($scriptContent.Substring(0, $preamble.Length) -eq $preamble)

if (-not $hasBOM) {
    Write-Host "警告: このスクリプトはUTF-8 with BOMでエンコードされていません。文字化けが発生する可能性があります。" -ForegroundColor Red
    Write-Host "scripts\modules\EncodingFix.psm1を使用して修正してください。" -ForegroundColor Yellow
    Start-Sleep -Seconds 3
}

# スクリプトの実行ポリシーを設定
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# モジュールパスを設定
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path -Path $scriptPath -ChildPath "modules"

# モジュールをインポート
if (-not (Test-Path "$modulePath\StatusReport.psm1")) {
    Write-Host "エラー: モジュールファイルが見つかりません: $modulePath\StatusReport.psm1" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path "$modulePath\Diagnostics.psm1")) {
    Write-Host "エラー: モジュールファイルが見つかりません: $modulePath\Diagnostics.psm1" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path "$modulePath\EncodingFix.psm1")) {
    Write-Host "エラー: モジュールファイルが見つかりません: $modulePath\EncodingFix.psm1" -ForegroundColor Red
    exit 1
}

Import-Module "$modulePath\StatusReport.psm1" -Force
Import-Module "$modulePath\Diagnostics.psm1" -Force
Import-Module "$modulePath\EncodingFix.psm1" -Force

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
