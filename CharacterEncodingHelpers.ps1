# 文字コード関連の共通ヘルパー関数
# このファイルは他のツールから参照されることを想定

# BOMの検出と表示
function Test-HasBom {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath -PathType Leaf)) {
        Write-Error "ファイルが見つかりません: $FilePath"
        return $false
    }
    
    $bytes = [byte[]](Get-Content $FilePath -Encoding Byte -ReadCount 3 -TotalCount 3)
    
    # UTF-8 BOM (EF BB BF)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return $true
    }
    
    return $false
}

# 文字コードエンコーディングの変換名前を正規化する
function Get-NormalizedEncodingName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EncodingName
    )
    
    switch ($EncodingName.ToLower()) {
        { $_ -in "utf8", "utf-8" } { return "UTF8" }
        { $_ -in "utf8bom", "utf-8-bom", "utf8-bom", "utf-8bom" } { return "UTF8BOM" }
        { $_ -in "sjis", "shift_jis", "shift-jis", "shiftjis" } { return "SJIS" }
        { $_ -in "eucjp", "euc-jp", "euc_jp" } { return "EUCJP" }
        { $_ -in "jis", "iso-2022-jp" } { return "JIS" }
        "ascii" { return "ASCII" }
        { $_ -in "unicode", "utf-16", "utf16" } { return "Unicode" }
        default { return $EncodingName }
    }
}

# 文字コード名からエンコーディングオブジェクトを取得
function Get-EncodingObject {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EncodingName
    )
    
    $normalizedName = Get-NormalizedEncodingName -EncodingName $EncodingName
    
    switch ($normalizedName) {
        "UTF8" { return New-Object System.Text.UTF8Encoding($false) } # BOMなしUTF-8
        "UTF8BOM" { return New-Object System.Text.UTF8Encoding($true) } # BOMありUTF-8
        "SJIS" { return [System.Text.Encoding]::GetEncoding("shift-jis") }
        "EUCJP" { return [System.Text.Encoding]::GetEncoding("euc-jp") }
        "JIS" { return [System.Text.Encoding]::GetEncoding("iso-2022-jp") }
        "ASCII" { return [System.Text.Encoding]::ASCII }
        "Unicode" { return [System.Text.Encoding]::Unicode }
        default { 
            Write-Warning "未知の文字コード: $EncodingName。UTF-8(BOMなし)を使用します。"
            return New-Object System.Text.UTF8Encoding($false)
        }
    }
}

# ファイルの文字コード自動検出
function Get-FileEncoding {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath -PathType Leaf)) {
        Write-Error "ファイルが見つかりません: $FilePath"
        return $null
    }
    
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    
    # BOMチェック
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @{
            Name = "UTF8BOM"
            DisplayName = "UTF-8 (BOMあり)"
            Encoding = (New-Object System.Text.UTF8Encoding($true))
            HasBOM = $true
        }
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @{
            Name = "Unicode"
            DisplayName = "Unicode (UTF-16 LE)"
            Encoding = [System.Text.Encoding]::Unicode
            HasBOM = $true
        }
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return @{
            Name = "BigEndianUnicode"
            DisplayName = "Unicode (UTF-16 BE)"
            Encoding = [System.Text.Encoding]::BigEndianUnicode
            HasBOM = $true
        }
    }
    
    # 各エンコーディングで一致率を確認
    $encodings = @(
        @{ Name = "UTF8"; DisplayName = "UTF-8"; Encoding = [System.Text.Encoding]::UTF8; HasBOM = $false },
        @{ Name = "SJIS"; DisplayName = "Shift-JIS"; Encoding = [System.Text.Encoding]::GetEncoding("shift-jis"); HasBOM = $false },
        @{ Name = "EUCJP"; DisplayName = "EUC-JP"; Encoding = [System.Text.Encoding]::GetEncoding("euc-jp"); HasBOM = $false },
        @{ Name = "JIS"; DisplayName = "ISO-2022-JP"; Encoding = [System.Text.Encoding]::GetEncoding("iso-2022-jp"); HasBOM = $false },
        @{ Name = "ASCII"; DisplayName = "ASCII"; Encoding = [System.Text.Encoding]::ASCII; HasBOM = $false }
    )
    
    $results = @()
    
    foreach ($enc in $encodings) {
        try {
            $encoding = $enc.Encoding
            $decodedString = $encoding.GetString($bytes)
            $reEncodedBytes = $encoding.GetBytes($decodedString)
            
            $matchCount = 0
            $compareLength = [Math]::Min($bytes.Length, $reEncodedBytes.Length)
            
            for ($i = 0; $i -lt $compareLength; $i++) {
                if ($bytes[$i] -eq $reEncodedBytes[$i]) {
                    $matchCount++
                }
            }
            
            $confidence = [Math]::Round(($matchCount / $compareLength) * 100, 2)
            
            $results += [PSCustomObject]@{
                Name = $enc.Name
                DisplayName = $enc.DisplayName
                Encoding = $encoding
                Confidence = $confidence
                HasBOM = $enc.HasBOM
            }
        }
        catch {
            # エラー時は無視
        }
    }
    
    # 信頼度でソート
    $sortedResults = $results | Sort-Object -Property Confidence -Descending
    
    return $sortedResults[0]
}
