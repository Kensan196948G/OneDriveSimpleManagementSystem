#
# ファイルエンコーディング検出スクリプト
# フォルダ内のファイルのエンコーディングを検出して表示します
#

# UTF-8エンコーディングの設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function Get-FileEncoding {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        
        # BOM検出
        if ($bytes.Length -ge 4 -and $bytes[0] -eq 0x00 -and $bytes[1] -eq 0x00 -and $bytes[2] -eq 0xFE -and $bytes[3] -eq 0xFF) {
            return "UTF-32 BE BOM"
        } elseif ($bytes.Length -ge 4 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE -and $bytes[2] -eq 0x00 -and $bytes[3] -eq 0x00) {
            return "UTF-32 LE BOM"
        } elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            return "UTF-8 BOM"
        } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            return "UTF-16 BE BOM"
        } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            return "UTF-16 LE BOM"
        }
        
        # BOMなし場合のエンコーディング推測
        
        # ASCIIかどうかチェック
        $isAscii = $true
        foreach ($byte in $bytes) {
            if ($byte -gt 127) {
                $isAscii = $false
                break
            }
        }
        
        if ($isAscii) {
            return "ASCII"
        }
        
        # UTF-8 (BOMなし) の可能性を検討
        $isUtf8 = $true
        $i = 0
        while ($i -lt $bytes.Length) {
            # 先頭バイトから1バイト文字、2バイト文字、3バイト文字、4バイト文字を判別
            if ($bytes[$i] -lt 0x80) {
                # 1バイト文字
                $i++
            } elseif (($bytes[$i] -ge 0xC2) -and ($bytes[$i] -le 0xDF) -and ($i + 1 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0x80) -and ($bytes[$i + 1] -le 0xBF))) {
                # 2バイト文字
                $i += 2
            } elseif (($bytes[$i] -eq 0xE0) -and ($i + 2 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0xA0) -and ($bytes[$i + 1] -le 0xBF)) -and (($bytes[$i + 2] -ge 0x80) -and ($bytes[$i + 2] -le 0xBF))) {
                # 3バイト文字 (E0)
                $i += 3
            } elseif ((($bytes[$i] -ge 0xE1) -and ($bytes[$i] -le 0xEF)) -and ($i + 2 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0x80) -and ($bytes[$i + 1] -le 0xBF)) -and (($bytes[$i + 2] -ge 0x80) -and ($bytes[$i + 2] -le 0xBF))) {
                # 3バイト文字 (E1-EF)
                $i += 3
            } elseif (($bytes[$i] -eq 0xF0) -and ($i + 3 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0x90) -and ($bytes[$i + 1] -le 0xBF)) -and (($bytes[$i + 2] -ge 0x80) -and ($bytes[$i + 2] -le 0xBF)) -and (($bytes[$i + 3] -ge 0x80) -and ($bytes[$i + 3] -le 0xBF))) {
                # 4バイト文字 (F0)
                $i += 4
            } elseif ((($bytes[$i] -ge 0xF1) -and ($bytes[$i] -le 0xF3)) -and ($i + 3 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0x80) -and ($bytes[$i + 1] -le 0xBF)) -and (($bytes[$i + 2] -ge 0x80) -and ($bytes[$i + 2] -le 0xBF)) -and (($bytes[$i + 3] -ge 0x80) -and ($bytes[$i + 3] -le 0xBF))) {
                # 4バイト文字 (F1-F3)
                $i += 4
            } elseif (($bytes[$i] -eq 0xF4) -and ($i + 3 -lt $bytes.Length) -and (($bytes[$i + 1] -ge 0x80) -and ($bytes[$i + 1] -le 0x8F)) -and (($bytes[$i + 2] -ge 0x80) -and ($bytes[$i + 2] -le 0xBF)) -and (($bytes[$i + 3] -ge 0x80) -and ($bytes[$i + 3] -le 0xBF))) {
                # 4バイト文字 (F4)
                $i += 4
            } else {
                $isUtf8 = $false
                break
            }
        }
        
        if ($isUtf8) {
            return "UTF-8 (BOM無し)"
        }
        
        # Shift-JISの特徴的なバイトシーケンスをチェック
        $shiftJisLikeliHood = 0
        $i = 0
        while ($i -lt $bytes.Length - 1) {
            # Shift-JIS 第一バイト範囲をチェック
            if ((($bytes[$i] -ge 0x81) -and ($bytes[$i] -le 0x9F)) -or (($bytes[$i] -ge 0xE0) -and ($bytes[$i] -le 0xEF))) {
                # 第二バイトチェック
                if ((($bytes[$i + 1] -ge 0x40) -and ($bytes[$i + 1] -le 0x7E)) -or (($bytes[$i + 1] -ge 0x80) -and ($bytes[$i + 1] -le 0xFC))) {
                    $shiftJisLikeliHood++
                }
                $i += 2
            } else {
                $i++
            }
        }
        
        # 日本語環境ではShift-JISの可能性が高い
        if ($shiftJisLikeliHood > 5) {
            return "Shift-JIS"
        }
        
        # それ以外はUTF-8かISO-8859-1（Latin-1）と推測
        try {
            $testContent = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
            return "UTF-8 (BOM無し) [推測]"
        } catch {
            return "不明 (おそらくShift-JISまたはCP932)"
        }
    } catch {
        return "エラー: $_" 
    }
}

function Test-Utf8EncodingCompatible {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [Parameter(Mandatory=$false)]
        [string]$Encoding = "不明"
    )
    
    try {
        # UTF-8でファイルを読み込んで内容を確認
        $utf8Content = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
        
        # 文字化けを検出するための一般的なパターン
        $garbledPattern = '�|▒|■|□|△|▲|▼|◆|◇|○|●|★|☆|\?\?\?'
        
        if ($utf8Content -match $garbledPattern) {
            # 文字化けの可能性が高い
            return $false
        } else {
            # 日本語文字が含まれていることを確認（ひらがな、カタカナ、漢字の範囲）
            if ($utf8Content -match '[ぁ-んァ-ン一-龯]') {
                return $true
            }
            
            # エンコーディングからチェック
            if ($Encoding -like "*UTF-8*" -or $Encoding -eq "ASCII") {
                return $true
            }
            
            # 特に判断できない場合は「要確認」
            return "要確認"
        }
    } catch {
        return "エラー: $_"
    }
}

# メイン処理
$targetPath = ""
$targetPath = Read-Host "フォルダまたはファイルのパスを入力してください（空白の場合は現在のスクリプトフォルダを使用）"

if ([string]::IsNullOrEmpty($targetPath)) {
    $targetPath = $PSScriptRoot
}

if (-not (Test-Path $targetPath)) {
    Write-Host "エラー: 指定されたパス '$targetPath' が存在しません。" -ForegroundColor Red
    exit 1
}

$isFolder = (Get-Item $targetPath) -is [System.IO.DirectoryInfo]

# ヘッダー表示
Write-Host "`nファイルエンコーディング検査ツール" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "対象: $targetPath`n" -ForegroundColor Cyan

# 結果を保存する配列
$results = @()

if ($isFolder) {
    # 対象の拡張子
    $extensions = @("*.ps1", "*.psm1", "*.bat", "*.cmd", "*.md", "*.txt", "*.json")
    
    $files = Get-ChildItem -Path $targetPath -Include $extensions -Recurse -File
    
    foreach ($file in $files) {
        $encoding = Get-FileEncoding -FilePath $file.FullName
        $isUtf8Compatible = Test-Utf8EncodingCompatible -FilePath $file.FullName -Encoding $encoding
        
        $status = if ($isUtf8Compatible -eq $true) {
            "正常"
        } elseif ($isUtf8Compatible -eq $false) {
            "文字化けの可能性"
        } elseif ($isUtf8Compatible -eq "要確認") {
            "要確認"
        } else {
            "エラー"
        }
        
        $statusColor = switch ($status) {
            "正常" { "Green" }
            "文字化けの可能性" { "Red" }
            "要確認" { "Yellow" }
            "エラー" { "Magenta" }
            default { "White" }
        }
        
        $relPath = $file.FullName.Replace("$targetPath\", "")
        
        Write-Host "ファイル: " -NoNewline
        Write-Host $relPath -ForegroundColor Cyan -NoNewline
        Write-Host " - エンコーディング: " -NoNewline
        Write-Host $encoding -ForegroundColor Yellow -NoNewline
        Write-Host " - 状態: " -NoNewline
        Write-Host $status -ForegroundColor $statusColor
        
        # 結果を配列に追加
        $results += [PSCustomObject]@{
            ファイル = $relPath
            エンコーディング = $encoding
            状態 = $status
            パス = $file.FullName
        }
    }
} else {
    $file = Get-Item $targetPath
    $encoding = Get-FileEncoding -FilePath $file.FullName
    $isUtf8Compatible = Test-Utf8EncodingCompatible -FilePath $file.FullName -Encoding $encoding
    
    $status = if ($isUtf8Compatible -eq $true) {
        "正常"
    } elseif ($isUtf8Compatible -eq $false) {
        "文字化けの可能性"
    } elseif ($isUtf8Compatible -eq "要確認") {
        "要確認"
    } else {
        "エラー"
    }
    
    $statusColor = switch ($status) {
        "正常" { "Green" }
        "文字化けの可能性" { "Red" }
        "要確認" { "Yellow" }
        "エラー" { "Magenta" }
        default { "White" }
    }
    
    Write-Host "ファイル: " -NoNewline
    Write-Host $file.Name -ForegroundColor Cyan -NoNewline
    Write-Host " - エンコーディング: " -NoNewline
    Write-Host $encoding -ForegroundColor Yellow -NoNewline
    Write-Host " - 状態: " -NoNewline
    Write-Host $status -ForegroundColor $statusColor
    
    # 結果を配列に追加
    $results += [PSCustomObject]@{
        ファイル = $file.Name
        エンコーディング = $encoding
        状態 = $status
        パス = $file.FullName
    }
}

# 結果のサマリーを表示
$problemFiles = $results | Where-Object { $_.状態 -ne "正常" }
$normalFiles = $results | Where-Object { $_.状態 -eq "正常" }

Write-Host "`n====== 結果サマリー ======" -ForegroundColor Cyan
Write-Host "検査したファイル数: $($results.Count)" -ForegroundColor White
Write-Host "正常なファイル数: $($normalFiles.Count)" -ForegroundColor Green
Write-Host "問題のあるファイル数: $($problemFiles.Count)" -ForegroundColor $(if ($problemFiles.Count -gt 0) { "Red" } else { "Green" })

if ($problemFiles.Count -gt 0) {
    Write-Host "`n問題のあるファイル一覧:" -ForegroundColor Yellow
    foreach ($file in $problemFiles) {
        $statusColor = switch ($file.状態) {
            "文字化けの可能性" { "Red" }
            "要確認" { "Yellow" }
            "エラー" { "Magenta" }
            default { "White" }
        }
        Write-Host "$($file.ファイル) - $($file.エンコーディング)" -ForegroundColor $statusColor
    }
    
    # 修正方法のアドバイス
    Write-Host "`n修正方法:" -ForegroundColor Cyan
    Write-Host "1. EncodingFixer.bat を実行してGUIツールでファイルを選択・変換" -ForegroundColor White
    Write-Host "2. ConvertAllScripts.bat を実行して一括変換" -ForegroundColor White
    Write-Host "3. 個別に変換するには以下のコマンドをPowerShellで実行:" -ForegroundColor White
    Write-Host '   $content = [System.IO.File]::ReadAllText("ファイルパス", [System.Text.Encoding]::Default); [System.IO.File]::WriteAllText("ファイルパス", $content, [System.Text.Encoding]::UTF8)' -ForegroundColor Gray
}

Write-Host "`n処理が完了しました。Enterキーを押すと終了します..." -ForegroundColor Cyan
Read-Host
