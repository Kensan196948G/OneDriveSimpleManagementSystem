# OneDriveStatusCheck.ps1を完全にクリーンアップするスクリプト
$filePath = Join-Path $PSScriptRoot "OneDriveStatusCheck.ps1"

# ファイルの内容を読み込む
$content = Get-Content -Path $filePath -Raw

# 先頭の不正な文字を削除（正規表現を使用）
$cleanContent = $content -replace "^[\ufeff\u200b\s]*#", "#"

# 一時ファイルに書き込む
$tempFilePath = Join-Path $PSScriptRoot "OneDriveStatusCheck_clean.ps1"
Set-Content -Path $tempFilePath -Value $cleanContent -Encoding UTF8 -NoNewline

# 元のファイルをバックアップ
$backupFilePath = Join-Path $PSScriptRoot "OneDriveStatusCheck_backup.ps1"
Copy-Item -Path $filePath -Destination $backupFilePath -Force

# 一時ファイルを元のファイルに移動
Move-Item -Path $tempFilePath -Destination $filePath -Force

Write-Host "OneDriveStatusCheck.ps1を完全にクリーンアップしました。" -ForegroundColor Green
Write-Host "バックアップファイルは $backupFilePath に保存されています。" -ForegroundColor Yellow