# スクリプト文字化け修正モジュール
# 文字コード: UTF-8 with BOM

function Repair-ScriptEncoding {
    Clear-Host
    Write-Host "=====================================================
  スクリプト文字化け修正
=====================================================" -ForegroundColor Cyan
    
    try {
        $scriptPath = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
        Write-Host "スクリプトファイルの文字エンコーディングを修正しています..." -ForegroundColor Yellow
        
        # スクリプトファイルのリスト
        $scriptFiles = Get-ChildItem -Path $scriptPath -Recurse -Include "*.ps1", "*.psm1"
        
        foreach ($file in $scriptFiles) {
            Write-Host "ファイル確認中: $($file.FullName)" -NoNewline
            
            # ファイルのエンコーディングを確認
            $content = Get-Content -Path $file.FullName -Raw
            $hasBOM = $content.StartsWith([char]0xFEFF)
            
            if ($hasBOM) {
                Write-Host " - OK (UTF-8 with BOM)" -ForegroundColor Green
            } else {
                Write-Host " - 修正が必要 (BOMなし)" -ForegroundColor Yellow
                
                # エンコードを修正してファイルを上書き
                $content | Out-File -FilePath $file.FullName -Encoding utf8
                Write-Host "  → UTF-8 with BOMに変換しました" -ForegroundColor Cyan
            }
        }
        
        Write-Host "`nスクリプトファイルのエンコーディング修正が完了しました。" -ForegroundColor Green
        
    } catch {
        Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    }
    
    Write-Host "`n任意のキーを押して続行してください..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Export-ModuleMember -Function Repair-ScriptEncoding
