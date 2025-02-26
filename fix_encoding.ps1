# OneDriveStatusCheck.ps1のエンコーディングを修正するスクリプト
$filePath = Join-Path $PSScriptRoot "OneDriveStatusCheck.ps1"

# ファイルの内容を読み込む（エンコーディングを指定せずに読み込む）
$content = [System.IO.File]::ReadAllText($filePath)

# 先頭のBOMと余分な文字を削除
$content = $content -replace "^[\ufeff\u200b\s]*#", "#"

# BOMなしのUTF-8でファイルを書き込む
[System.IO.File]::WriteAllText($filePath, $content, [System.Text.Encoding]::UTF8)

Write-Host "OneDriveStatusCheck.ps1のエンコーディングを修正しました。" -ForegroundColor Green