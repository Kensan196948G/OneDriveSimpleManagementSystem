#
# 文字コード変換ツール
# PowerShell スクリプトの文字コード問題を解決するためのユーティリティ
#
# 使用法:
#   .\CharacterEncodingTool.ps1 -InputFile "変換したいファイル.ps1" -SourceEncoding "shift-jis" -TargetEncoding "utf-8"
#
# パラメーター:
#   -InputFile      : 変換するファイルのパス
#   -SourceEncoding : 元の文字コード (省略時は自動検出)
#   -TargetEncoding : 変換後の文字コード (デフォルト: utf-8)
#   -BOM            : BOMの有無 (デフォルト: $true)

param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $false)]
    [string]$SourceEncoding = "auto",
    
    [Parameter(Mandatory = $false)]
    [string]$TargetEncoding = "utf-8",
    
    [Parameter(Mandatory = $false)]
    [bool]$BOM = $true
)

# ファイルの存在確認
if (-not (Test-Path $InputFile)) {
    Write-Host "エラー: ファイル '$InputFile' が見つかりません。" -ForegroundColor Red
    exit 1
}

function Detect-FileEncoding {
    param (
        [string]$FilePath
    )
    
    # ファイルの先頭部分を読み込む
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    
    # BOM チェック
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @{Encoding = "utf-8"; BOM = $true; Confidence = 100}
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return @{Encoding = "utf-16BE"; BOM = $true; Confidence = 100}
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @{Encoding = "utf-16"; BOM = $true; Confidence = 100}
    }
    
    # 文字コードの推測
    $content = [System.IO.File]::ReadAllBytes($FilePath)
    
    # ASCII チェック
    $isASCII = $true
    foreach ($byte in $content) {
        if ($byte -gt 127) {
            $isASCII = $false
            break
        }
    }
    
    if ($isASCII) {
        return @{Encoding = "ascii"; BOM = $false; Confidence = 90}
    }
    
    # UTF-8 チェック (BOMなし)
    $isUTF8 = $true
    $i = 0
    while ($i -lt $content.Length) {
        if ($content[$i] -lt 128) {
            $i++
            continue
        }
        
        # UTF-8の複数バイト文字を検出
        if (($content[$i] -ge 0xC2 -and $content[$i] -le 0xDF) -and
            ($i + 1 -lt $content.Length) -and
            ($content[$i + 1] -ge 0x80 -and $content[$i + 1] -le 0xBF)) {
            $i += 2
        }
        elseif (($content[$i] -ge 0xE0 -and $content[$i] -le 0xEF) -and
                ($i + 2 -lt $content.Length) -and
                ($content[$i + 1] -ge 0x80 -and $content[$i + 1] -le 0xBF) -and
                ($content[$i + 2] -ge 0x80 -and $content[$i + 2] -le 0xBF)) {
            $i += 3
        }
        elseif (($content[$i] -ge 0xF0 -and $content[$i] -le 0xF7) -and
                ($i + 3 -lt $content.Length) -and
                ($content[$i + 1] -ge 0x80 -and $content[$i + 1] -le 0xBF) -and
                ($content[$i + 2] -ge 0x80 -and $content[$i + 2] -le 0xBF) -and
                ($content[$i + 3] -ge 0x80 -and $content[$i + 3] -le 0xBF)) {
            $i += 4
        }
        else {
            $isUTF8 = $false
            break
        }
    }
    
    if ($isUTF8) {
        return @{Encoding = "utf-8"; BOM = $false; Confidence = 85}
    }
    
    # Shift-JIS または EUC-JP の可能性をチェック
    $sjisCount = 0
    $eucjpCount = 0
    
    $i = 0
    while ($i -lt $content.Length - 1) {
        # Shift-JIS の特徴をチェック
        if (($content[$i] -ge 0x81 -and $content[$i] -le 0x9F) -or
            ($content[$i] -ge 0xE0 -and $content[$i] -le 0xFC)) {
            if (($content[$i + 1] -ge 0x40 -and $content[$i + 1] -le 0xFC) -and
                $content[$i + 1] -ne 0x7F) {
                $sjisCount++
                $i += 2
                continue
            }
        }
        
        # EUC-JP の特徴をチェック
        if ($content[$i] -ge 0x8E -and $content[$i] -le 0xFE) {
            if ($content[$i + 1] -ge 0xA1 -and $content[$i + 1] -le 0xFE) {
                $eucjpCount++
                $i += 2
                continue
            }
        }
        
        $i++
    }
    
    if ($sjisCount -gt $eucjpCount) {
        return @{Encoding = "shift-jis"; BOM = $false; Confidence = 75}
    }
    elseif ($eucjpCount -gt $sjisCount) {
        return @{Encoding = "euc-jp"; BOM = $false; Confidence = 75}
    }
    
    # デフォルトは Shift-JIS (日本環境で最も一般的)
    return @{Encoding = "shift-jis"; BOM = $false; Confidence = 60}
}

# メイン処理
try {
    # 文字コードの自動検出
    if ($SourceEncoding -eq "auto") {
        $encodingInfo = Detect-FileEncoding -FilePath $InputFile
        $SourceEncoding = $encodingInfo.Encoding
        Write-Host "ファイルを分析中: $InputFile" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "検出された文字コード: $($SourceEncoding) (信頼度: $($encodingInfo.Confidence)%)" -ForegroundColor Yellow
        Write-Host ""
    }
    
    # 入力エンコーディングの設定
    $srcEncoding = switch ($SourceEncoding.ToLower()) {
        "utf-8"     { [System.Text.Encoding]::UTF8 }
        "utf-16"    { [System.Text.Encoding]::Unicode }
        "utf-16be"  { [System.Text.Encoding]::BigEndianUnicode }
        "shift-jis" { [System.Text.Encoding]::GetEncoding(932) }
        "euc-jp"    { [System.Text.Encoding]::GetEncoding(51932) }
        "ascii"     { [System.Text.Encoding]::ASCII }
        default     { [System.Text.Encoding]::UTF8 }
    }
    
    # 出力エンコーディングの設定
    $dstEncoding = switch ($TargetEncoding.ToLower()) {
        "utf-8"     { New-Object System.Text.UTF8Encoding($BOM) }
        "utf-16"    { [System.Text.Encoding]::Unicode }
        "utf-16be"  { [System.Text.Encoding]::BigEndianUnicode }
        "shift-jis" { [System.Text.Encoding]::GetEncoding(932) }
        "euc-jp"    { [System.Text.Encoding]::GetEncoding(51932) }
        "ascii"     { [System.Text.Encoding]::ASCII }
        default     { New-Object System.Text.UTF8Encoding($BOM) }
    }
    
    # ファイル内容を読み込む
    $content = [System.IO.File]::ReadAllText($InputFile, $srcEncoding)
    
    # ファイル内容のプレビュー表示
    Write-Host "ファイル内容プレビュー (最初の50行):" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------"
    $contentLines = $content -split "`r?`n"
    $previewLines = [Math]::Min(50, $contentLines.Length)
    for ($i = 0; $i -lt $previewLines; $i++) {
        Write-Host $contentLines[$i]
    }
    Write-Host "--------------------------------------------------"
    
    # バックアップファイル名
    $backupFile = "$InputFile.bak"
    $counter = 1
    while (Test-Path $backupFile) {
        $backupFile = "$InputFile.bak$counter"
        $counter++
    }
    
    # バックアップを作成
    Copy-Item -Path $InputFile -Destination $backupFile
    
    # 変換して書き込み
    [System.IO.File]::WriteAllText($InputFile, $content, $dstEncoding)
    
    Write-Host "文字コード変換が完了しました:" -ForegroundColor Green
    Write-Host "  変換前: $SourceEncoding" -ForegroundColor Cyan
    Write-Host "  変換後: $TargetEncoding (BOM: $BOM)" -ForegroundColor Cyan
    Write-Host "  バックアップ: $backupFile" -ForegroundColor Yellow
}
catch {
    Write-Host "ファイル内容の読み込み中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ファイル分析中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}
