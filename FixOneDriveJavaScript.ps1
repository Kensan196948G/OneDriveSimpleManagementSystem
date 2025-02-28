# PowerShellスクリプトは通常UTF-8 with BOMでエンコードします
$ErrorActionPreference = "Stop"

# OneDriveのJavaScriptファイルを修正するスクリプト
Write-Host "OneDriveのJavaScriptファイルを修正しています..." -ForegroundColor Cyan

# OneDriveのJavaScriptファイルを検索
$jsFilePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe.WebView2\EBWebView\Default\File System\001\t\Paths\LOG.old"

if (Test-Path $jsFilePath) {
    # バックアップの作成
    $backup = "$jsFilePath.backup"
    if (-not (Test-Path $backup)) {
        Write-Host "バックアップを作成しています: $backup" -ForegroundColor Yellow
        Copy-Item -Path $jsFilePath -Destination $backup -Force
    }
    
    try {
        # ファイルを読み込む
        $content = Get-Content -Path $jsFilePath -Encoding UTF8
        
        # 修正処理
        # 具体的な修正内容に応じてreplace処理を追加
        # $content = $content -replace "問題のある部分", "修正後の部分"
        
        # 修正したコンテンツを保存
        $content | Set-Content -Path $jsFilePath -Encoding UTF8
        Write-Host "JavaScriptファイルの修正が完了しました: $jsFilePath" -ForegroundColor Green
    }
    catch {
        Write-Host "JavaScriptファイルの修正中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $backup) {
            Write-Host "バックアップから復元しています..." -ForegroundColor Yellow
            Copy-Item -Path $backup -Destination $jsFilePath -Force
        }
    }
} else {
    Write-Host "JavaScriptファイルが見つかりません: $jsFilePath" -ForegroundColor Red
}
