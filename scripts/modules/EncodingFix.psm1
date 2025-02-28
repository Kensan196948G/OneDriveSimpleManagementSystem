# エンコーディング修正モジュール
# 文字コード: UTF-8 with BOM

function Repair-ScriptEncoding {
    Clear-Host
    Write-Host "==========================================
  文字化け修正ツール
==========================================" -ForegroundColor Cyan
    
    # スクリプトの実際のルートパスを明示的に指定
    $appRootPath = "C:\kitting\OneDrive運用ツール\OneDriveSimpleManagementSystem"
    if (-not (Test-Path $appRootPath)) {
        Write-Host "エラー: アプリケーションのルートパスが見つかりません: $appRootPath" -ForegroundColor Red
        Write-Host "何かキーを押すとメインメニューに戻ります..."
        [void][System.Console]::ReadKey($true)
        return
    }
    
    Write-Host "スクリプトフォルダをスキャンしています: $appRootPath" -ForegroundColor Yellow
    
    # アプリケーション内のPowerShellファイルだけを対象とする
    $files = Get-ChildItem -Path $appRootPath -Recurse -Include *.ps1,*.psm1
    $fixedCount = 0
    
    foreach ($file in $files) {
        try {
            $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8 -ErrorAction Stop
            $encoding = [System.Text.Encoding]::UTF8
            $preamble = $encoding.GetPreamble()
            $hasBOM = $content.Length -ge $preamble.Length -and 
                    ($content.Substring(0, $preamble.Length) -eq $preamble)
            
            if (-not $hasBOM) {
                Write-Host "修正中: $($file.FullName)" -ForegroundColor Yellow
                try {
                    $content | Out-File -FilePath $file.FullName -Encoding utf8 -NoNewline -ErrorAction Stop
                    $fixedCount++
                }
                catch {
                    Write-Host "  → 修正失敗: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "OK: $($file.FullName)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "エラー: $($file.FullName) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`n修正完了: $fixedCount ファイルを修正しました。" -ForegroundColor Green
    Write-Host "再起動後に日本語が正しく表示されるようになります。" -ForegroundColor Cyan
    Write-Host "何かキーを押すとメインメニューに戻ります..."
    [void][System.Console]::ReadKey($true)
}

Export-ModuleMember -Function Repair-ScriptEncoding
