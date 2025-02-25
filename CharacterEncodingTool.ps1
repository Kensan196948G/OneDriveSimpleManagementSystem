#
# 文字コード変換ツール
# 文字化けの解決をサポートするためのPowerShellスクリプト
#

# パラメータ定義
param (
    [Parameter(Mandatory=$false)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$SourceEncoding = "utf-8",
    
    [Parameter(Mandatory=$false)]
    [string]$TargetEncoding = "utf-8",
    
    [Parameter(Mandatory=$false)]
    [switch]$BOM = $false,
    
    [switch]$Detect = $false,
    
    [switch]$ShowHelp = $false
)

# ヘルプを表示
function Show-Help {
    Write-Host "文字コード変換ツール - 使用方法:"
    Write-Host "CharacterEncodingTool.ps1 -InputFile <ファイルパス> [-OutputFile <出力パス>] [-SourceEncoding <元の文字コード>] [-TargetEncoding <変換後の文字コード>] [-BOM] [-Detect]"
    Write-Host ""
    Write-Host "パラメータ:"
    Write-Host "  -InputFile       : 変換するファイル"
    Write-Host "  -OutputFile      : 出力ファイル (省略時は元のファイルを上書き)"
    Write-Host "  -SourceEncoding  : 元の文字コード (省略時はUTF-8)"
    Write-Host "  -TargetEncoding  : 変換後の文字コード (省略時はUTF-8)"
    Write-Host "  -BOM             : UTF-8の場合、BOMを付与する"
    Write-Host "  -Detect          : ファイルの文字コードを検出して表示"
    Write-Host "  -ShowHelp        : このヘルプを表示"
    Write-Host ""
    Write-Host "使用可能な文字コード: utf-8, shift-jis, euc-jp, iso-2022-jp, ascii, unicode"
    Write-Host ""
    Write-Host "使用例:"
    Write-Host "  1. ファイルの文字コードを検出:"
    Write-Host "     .\CharacterEncodingTool.ps1 -InputFile file.txt -Detect"
    Write-Host ""
    Write-Host "  2. Shift-JISからUTF-8(BOMなし)に変換:"
    Write-Host "     .\CharacterEncodingTool.ps1 -InputFile file.txt -SourceEncoding shift-jis -TargetEncoding utf-8"
    Write-Host ""
    Write-Host "  3. Shift-JISからUTF-8(BOMあり)に変換:"
    Write-Host "     .\CharacterEncodingTool.ps1 -InputFile file.txt -SourceEncoding shift-jis -TargetEncoding utf-8 -BOM"
}

# エンコーディングを正規化する関数
function Convert-EncodingName {
    param(
        [string]$EncodingName
    )
    
    switch ($EncodingName.ToLower()) {
        "utf-8" { return "utf-8" }
        "utf8" { return "utf-8" }
        "shift_jis" { return "shift-jis" }
        "sjis" { return "shift-jis" }
        "ascii" { return "ascii" }
        "unicode" { return "unicode" }
        default { return $EncodingName }
    }
}

# 文字コード検出の試行
function Detect-Encoding {
    param (
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath -PathType Leaf)) {
        Write-Host "ファイルが見つかりません: $FilePath" -ForegroundColor Red
        return
    }
    
    $encodings = @(
        [System.Text.Encoding]::UTF8,
        [System.Text.Encoding]::GetEncoding("shift-jis"),
        [System.Text.Encoding]::GetEncoding("euc-jp"),
        [System.Text.Encoding]::GetEncoding("iso-2022-jp"),
        [System.Text.Encoding]::ASCII,
        [System.Text.Encoding]::Unicode
    )
    
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    
    # BOM（Byte Order Mark）の確認
    $hasBOM = $false
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $hasBOM = $true
        Write-Host "UTF-8 BOMが検出されました。" -ForegroundColor Cyan
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        Write-Host "Unicode (UTF-16 LE) BOMが検出されました。" -ForegroundColor Cyan
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        Write-Host "Unicode (UTF-16 BE) BOMが検出されました。" -ForegroundColor Cyan
    }
    
    $results = @()
    
    foreach ($encoding in $encodings) {
        try {
            $decodedString = $encoding.GetString($bytes)
            $reEncodedBytes = $encoding.GetBytes($decodedString)
            
            # BOMを考慮して比較
            $matchCount = 0
            $compareLength = [Math]::Min($bytes.Length, $reEncodedBytes.Length)
            
            for ($i = 0; $i -lt $compareLength; $i++) {
                if ($bytes[$i] -eq $reEncodedBytes[$i]) {
                    $matchCount++
                }
            }
            
            $confidencePercent = [Math]::Round(($matchCount / $compareLength) * 100, 2)
            
            $encodingName = $encoding.WebName
            if ($encodingName -eq "utf-8" -and $hasBOM) {
                $encodingName = "utf-8 (BOMあり)"
            }
            
            $results += [PSCustomObject]@{
                Encoding = $encodingName
                Confidence = $confidencePercent
            }
        }
        catch {
            # エラーが発生した場合は無視
        }
    }
    
    # 信頼度でソート
    $results = $results | Sort-Object -Property Confidence -Descending
    
    Write-Host "検出された可能性のある文字コード:" -ForegroundColor Cyan
    $results | Format-Table -AutoSize
    
    return $results[0].Encoding
}

# メイン処理
if ($ShowHelp) {
    Show-Help
    exit 0
}

if (-not $InputFile) {
    Write-Host "入力ファイルを指定してください。" -ForegroundColor Yellow
    Show-Help
    exit 1
}

if (-not (Test-Path $InputFile -PathType Leaf)) {
    Write-Host "ファイルが見つかりません: $InputFile" -ForegroundColor Red
    exit 1
}

if ($Detect) {
    Detect-Encoding -FilePath $InputFile
    exit 0
}

if (-not $OutputFile) {
    $OutputFile = $InputFile
}

# ソースエンコーディング名の正規化
$normalizedSourceEncoding = Convert-EncodingName -EncodingName $SourceEncoding

# ターゲットエンコーディング名の正規化
$normalizedTargetEncoding = Convert-EncodingName -EncodingName $TargetEncoding

try {
    Write-Host "ファイル変換開始: $InputFile" -ForegroundColor Cyan
    Write-Host "元の文字コード: $normalizedSourceEncoding" -ForegroundColor Cyan
    Write-Host "変換後の文字コード: $normalizedTargetEncoding" -ForegroundColor Cyan
    
    if ($normalizedTargetEncoding -eq "utf-8" -and $BOM) {
        Write-Host "UTF-8 BOMを付与します" -ForegroundColor Cyan
    }
    
    # ファイル内容を読み込む
    $content = $null
    
    # ソースエンコーディング
    switch ($normalizedSourceEncoding) {
        "utf-8" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::UTF8) }
        "shift-jis" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::GetEncoding("shift-jis")) }
        "euc-jp" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::GetEncoding("euc-jp")) }
        "iso-2022-jp" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::GetEncoding("iso-2022-jp")) }
        "ascii" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::ASCII) }
        "unicode" { $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::Unicode) }
        default { 
            Write-Host "未対応の文字コードです: $normalizedSourceEncoding" -ForegroundColor Red
            exit 1
        }
    }
    
    # ターゲットエンコーディング
    switch ($normalizedTargetEncoding) {
        "utf-8" { 
            if ($BOM) {
                $utf8Bom = New-Object System.Text.UTF8Encoding($true)
                [System.IO.File]::WriteAllText($OutputFile, $content, $utf8Bom)
            } else {
                [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::UTF8)
            }
        }
        "shift-jis" { [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::GetEncoding("shift-jis")) }
        "euc-jp" { [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::GetEncoding("euc-jp")) }
        "iso-2022-jp" { [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::GetEncoding("iso-2022-jp")) }
        "ascii" { [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::ASCII) }
        "unicode" { [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::Unicode) }
        default { 
            Write-Host "未対応の文字コードです: $normalizedTargetEncoding" -ForegroundColor Red
            exit 1
        }
    }
    
    Write-Host "ファイルを変換しました: $InputFile -> $OutputFile ($normalizedSourceEncoding -> $normalizedTargetEncoding)" -ForegroundColor Green
    
    if ($normalizedTargetEncoding -eq "utf-8" -and $BOM) {
        Write-Host "UTF-8 BOMが付与されました" -ForegroundColor Green
    }
}
catch {
    Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    exit 1
}
