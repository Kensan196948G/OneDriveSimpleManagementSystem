# OneDriveStatusCheck.ps1のBOMを完全に削除するスクリプト
$filePath = Join-Path $PSScriptRoot "OneDriveStatusCheck.ps1"

# ファイルの内容をバイナリとして読み込む
$bytes = [System.IO.File]::ReadAllBytes($filePath)

# BOMを検出して削除
$bomLength = 0
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    # UTF-8 BOM
    $bomLength = 3
} elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    # UTF-16 BE BOM
    $bomLength = 2
} elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    # UTF-16 LE BOM
    $bomLength = 2
} elseif ($bytes.Length -ge 4 -and $bytes[0] -eq 0x00 -and $bytes[1] -eq 0x00 -and $bytes[2] -eq 0xFE -and $bytes[3] -eq 0xFF) {
    # UTF-32 BE BOM
    $bomLength = 4
} elseif ($bytes.Length -ge 4 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE -and $bytes[2] -eq 0x00 -and $bytes[3] -eq 0x00) {
    # UTF-32 LE BOM
    $bomLength = 4
}

# BOMを削除した新しいバイト配列を作成
if ($bomLength -gt 0) {
    $newBytes = New-Object byte[] ($bytes.Length - $bomLength)
    [System.Array]::Copy($bytes, $bomLength, $newBytes, 0, $newBytes.Length)
} else {
    $newBytes = $bytes
}

# ファイルの内容をテキストに変換
$content = [System.Text.Encoding]::UTF8.GetString($newBytes)

# 先頭の不要な文字を削除
$content = $content -replace "^[\ufeff\u200b\s]*#", "#"

# 新しいファイルを作成
$tempFilePath = Join-Path $PSScriptRoot "OneDriveStatusCheck_new.ps1"
[System.IO.File]::WriteAllText($tempFilePath, $content, [System.Text.Encoding]::UTF8)

# 元のファイルを置き換え
Remove-Item $filePath -Force
Rename-Item $tempFilePath $filePath

Write-Host "OneDriveStatusCheck.ps1のBOMを完全に削除しました。" -ForegroundColor Green