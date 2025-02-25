#
# 文字化け診断・修正ツール - 修正版
# 文字化けしたファイルの診断と修復を行うスクリプト
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 共通ヘルパーのインポート
$helperPath = Join-Path $PSScriptRoot "CharacterEncodingHelpers.ps1"
if (Test-Path $helperPath) {
    . $helperPath
} else {
    Write-Warning "ヘルパースクリプトが見つかりません: $helperPath"
}

# フォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "文字化け診断・修正ツール"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# フォームアイコンの設定（アプリケーション共通アイコンがあれば）
$iconPath = Join-Path $PSScriptRoot "AppIcon.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
}

# ファイル選択部分
$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Location = New-Object System.Drawing.Point(10, 20)
$fileLabel.Size = New-Object System.Drawing.Size(100, 20)
$fileLabel.Text = "ファイルパス:"
$form.Controls.Add($fileLabel)

$fileTextBox = New-Object System.Windows.Forms.TextBox
$fileTextBox.Location = New-Object System.Drawing.Point(120, 20)
$fileTextBox.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($fileTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(680, 18)
$browseButton.Size = New-Object System.Drawing.Size(80, 23)
$browseButton.Text = "参照..."
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "すべてのファイル (*.*)|*.*|テキストファイル (*.txt)|*.txt|CSVファイル (*.csv)|*.csv|HTMLファイル (*.html;*.htm)|*.html;*.htm"
    $openFileDialog.FilterIndex = 1
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileTextBox.Text = $openFileDialog.FileName
        # ファイルが選択されたらデフォルトで検出を実行
        AnalyzeFile
    }
})
$form.Controls.Add($browseButton)

# 文字コード選択部分
$sourceEncodingLabel = New-Object System.Windows.Forms.Label
$sourceEncodingLabel.Location = New-Object System.Drawing.Point(10, 60)
$sourceEncodingLabel.Size = New-Object System.Drawing.Size(100, 20)
$sourceEncodingLabel.Text = "元の文字コード:"
$form.Controls.Add($sourceEncodingLabel)

$sourceEncodingComboBox = New-Object System.Windows.Forms.ComboBox
$sourceEncodingComboBox.Location = New-Object System.Drawing.Point(120, 60)
$sourceEncodingComboBox.Size = New-Object System.Drawing.Size(200, 20)
$sourceEncodingComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$sourceEncodingComboBox.Items.Add("自動検出")
[void]$sourceEncodingComboBox.Items.Add("UTF8")
[void]$sourceEncodingComboBox.Items.Add("UTF8BOM")
[void]$sourceEncodingComboBox.Items.Add("SJIS")
[void]$sourceEncodingComboBox.Items.Add("EUCJP")
[void]$sourceEncodingComboBox.Items.Add("JIS")
[void]$sourceEncodingComboBox.Items.Add("ASCII")
[void]$sourceEncodingComboBox.Items.Add("Unicode")
$sourceEncodingComboBox.SelectedIndex = 0
$form.Controls.Add($sourceEncodingComboBox)

$targetEncodingLabel = New-Object System.Windows.Forms.Label
$targetEncodingLabel.Location = New-Object System.Drawing.Point(340, 60)
$targetEncodingLabel.Size = New-Object System.Drawing.Size(120, 20)
$targetEncodingLabel.Text = "変換後の文字コード:"
$form.Controls.Add($targetEncodingLabel)

$targetEncodingComboBox = New-Object System.Windows.Forms.ComboBox
$targetEncodingComboBox.Location = New-Object System.Drawing.Point(470, 60)
$targetEncodingComboBox.Size = New-Object System.Drawing.Size(200, 20)
$targetEncodingComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$targetEncodingComboBox.Items.Add("UTF8")
[void]$targetEncodingComboBox.Items.Add("UTF8BOM")
[void]$targetEncodingComboBox.Items.Add("SJIS")
[void]$targetEncodingComboBox.Items.Add("EUCJP")
[void]$targetEncodingComboBox.Items.Add("JIS")
[void]$targetEncodingComboBox.Items.Add("ASCII")
[void]$targetEncodingComboBox.Items.Add("Unicode")
$targetEncodingComboBox.SelectedIndex = 0
$form.Controls.Add($targetEncodingComboBox)

# 実行ボタン
$analyzeButton = New-Object System.Windows.Forms.Button
$analyzeButton.Location = New-Object System.Drawing.Point(10, 100)
$analyzeButton.Size = New-Object System.Drawing.Size(150, 30)
$analyzeButton.Text = "文字コードを分析"
$analyzeButton.Add_Click({
    AnalyzeFile
})
$form.Controls.Add($analyzeButton)

$fixButton = New-Object System.Windows.Forms.Button
$fixButton.Location = New-Object System.Drawing.Point(180, 100)
$fixButton.Size = New-Object System.Drawing.Size(150, 30)
$fixButton.Text = "文字コードを変換"
$fixButton.Add_Click({
    FixEncoding
})
$form.Controls.Add($fixButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(350, 100)
$clearButton.Size = New-Object System.Drawing.Size(150, 30)
$clearButton.Text = "クリア"
$clearButton.Add_Click({
    $previewTextBox.Clear()
})
$form.Controls.Add($clearButton)

# ステータス表示用ラベル
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(520, 105)
$statusLabel.Size = New-Object System.Drawing.Size(250, 20)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$statusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($statusLabel)

# プレビュー部分
$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Location = New-Object System.Drawing.Point(10, 150)
$previewLabel.Size = New-Object System.Drawing.Size(100, 20)
$previewLabel.Text = "プレビュー:"
$form.Controls.Add($previewLabel)

$previewTextBox = New-Object System.Windows.Forms.TextBox
$previewTextBox.Location = New-Object System.Drawing.Point(10, 180)
$previewTextBox.Size = New-Object System.Drawing.Size(760, 350)
$previewTextBox.Multiline = $true
$previewTextBox.ScrollBars = "Both"
$previewTextBox.ReadOnly = $true
$previewTextBox.Font = New-Object System.Drawing.Font("MS Gothic", 10)
$form.Controls.Add($previewTextBox)

# 検出した文字コードを保存する変数
$script:detectedEncoding = $null

# ステータス更新関数
function Update-Status {
    param (
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::DarkBlue
    )
    
    $statusLabel.Text = $Message
    $statusLabel.ForeColor = $Color
    $form.Refresh()
}

# ファイルを分析する関数
function AnalyzeFile {
    $filePath = $fileTextBox.Text
    
    if (-not (Test-Path $filePath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません: $filePath", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $previewTextBox.Clear()
    Update-Status "ファイルを分析中..." [System.Drawing.Color]::DarkBlue
    $previewTextBox.AppendText("ファイルを分析中: $filePath`r`n`r`n")
    
    try {
        # ファイルのバイト配列を取得
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        
        # BOMチェック
        $hasBOM = $false
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $hasBOM = $true
            $previewTextBox.AppendText("UTF-8 BOMが検出されました。`r`n`r`n")
        }
        elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            $previewTextBox.AppendText("Unicode (UTF-16 LE) BOMが検出されました。`r`n`r`n")
        }
        elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            $previewTextBox.AppendText("Unicode (UTF-16 BE) BOMが検出されました。`r`n`r`n")
        }
        
        # 共通関数が利用可能な場合はそちらを使用
        if (Get-Command -Name Get-FileEncoding -ErrorAction SilentlyContinue) {
            $encodingInfo = Get-FileEncoding -FilePath $filePath
            $previewTextBox.AppendText("検出された文字コード: $($encodingInfo.DisplayName) (信頼度: $($encodingInfo.Confidence)%)`r`n`r`n")
            $script:detectedEncoding = $encodingInfo.Name
        }
        else {
            # 手動での文字コード検出
            $encodings = @(
                [System.Text.Encoding]::UTF8,
                [System.Text.Encoding]::GetEncoding("shift-jis"),
                [System.Text.Encoding]::GetEncoding("euc-jp"),
                [System.Text.Encoding]::GetEncoding("iso-2022-jp"),
                [System.Text.Encoding]::ASCII,
                [System.Text.Encoding]::Unicode
            )
            
            $results = @()
            
            foreach ($encoding in $encodings) {
                try {
                    $decodedString = $encoding.GetString($bytes)
                    $reEncodedBytes = $encoding.GetBytes($decodedString)
                    
                    $matchCount = 0
                    $compareLength = [Math]::Min($bytes.Length, $reEncodedBytes.Length)
                    
                    for ($i = 0; $i -lt $compareLength; $i++) {
                        if ($bytes[$i] -eq $reEncodedBytes[$i]) {
                            $matchCount++
                        }
                    }
                    
                    $confidencePercent = [Math]::Round(($matchCount / $compareLength) * 100, 2)
                    
                    $encodingName = switch ($encoding.WebName) {
                        "utf-8" { if ($hasBOM) { "UTF8BOM" } else { "UTF8" } }
                        "shift_jis" { "SJIS" }
                        "euc-jp" { "EUCJP" }
                        "iso-2022-jp" { "JIS" }
                        "us-ascii" { "ASCII" }
                        "unicode" { "Unicode" }
                        default { $encoding.WebName }
                    }
                    
                    $displayName = switch ($encoding.WebName) {
                        "utf-8" { if ($hasBOM) { "UTF-8 (BOMあり)" } else { "UTF-8" } }
                        "shift_jis" { "Shift-JIS" }
                        "euc-jp" { "EUC-JP" }
                        "iso-2022-jp" { "ISO-2022-JP" }
                        "us-ascii" { "ASCII" }
                        "unicode" { "Unicode (UTF-16)" }
                        default { $encoding.WebName }
                    }
                    
                    $results += [PSCustomObject]@{
                        Name = $encodingName
                        DisplayName = $displayName
                        Encoding = $encoding
                        Confidence = $confidencePercent
                    }
                }
                catch {
                    # エラーが発生した場合は無視
                }
            }
            
            # 信頼度でソート
            $sortedResults = $results | Sort-Object -Property Confidence -Descending
            
            $previewTextBox.AppendText("検出された可能性のある文字コード:`r`n")
            foreach ($result in $sortedResults) {
                $previewTextBox.AppendText("- $($result.DisplayName): $($result.Confidence)%`r`n")
            }
            
            # 最も可能性の高い文字コードを特定
            $bestMatch = $sortedResults[0]
            $previewTextBox.AppendText("`r`n最も可能性の高い文字コード: $($bestMatch.DisplayName)`r`n`r`n")
            $script:detectedEncoding = $bestMatch.Name
        }
        
        # ファイル内容のプレビュー表示
        try {
            # 検出されたエンコーディングでファイルを読み込む
            $encodingObj = if (Get-Command -Name Get-EncodingObject -ErrorAction SilentlyContinue) {
                Get-EncodingObject -EncodingName $script:detectedEncoding
            } else {
                # エンコーディング名から適切なオブジェクトを作成
                switch ($script:detectedEncoding) {
                    "UTF8" { [System.Text.Encoding]::UTF8 }
                    "UTF8BOM" { New-Object System.Text.UTF8Encoding($true) }
                    "SJIS" { [System.Text.Encoding]::GetEncoding("shift-jis") }
                    "EUCJP" { [System.Text.Encoding]::GetEncoding("euc-jp") }
                    "JIS" { [System.Text.Encoding]::GetEncoding("iso-2022-jp") }
                    "ASCII" { [System.Text.Encoding]::ASCII }
                    "Unicode" { [System.Text.Encoding]::Unicode }
                    default { [System.Text.Encoding]::UTF8 }
                }
            }
            
            $content = [System.IO.File]::ReadAllText($filePath, $encodingObj)
            $previewLines = $content -split "`r`n" | Select-Object -First 50
            $preview = [string]::Join("`r`n", $previewLines)
            
            $previewTextBox.AppendText("ファイル内容プレビュー (最初の50行):`r`n")
            $previewTextBox.AppendText("--------------------------------------------------`r`n")
            $previewTextBox.AppendText($preview)
            
            if (($content -split "`r`n").Length -gt 50) {
                $previewTextBox.AppendText("`r`n--------------------------------------------------`r`n")
                $previewTextBox.AppendText("(以下省略)")
            }
            
            Update-Status "文字コード分析が完了しました" [System.Drawing.Color]::Green
        }
        catch {
            $previewTextBox.AppendText("ファイル内容の読み込み中にエラーが発生しました: $_`r`n")
            Update-Status "エラーが発生しました" [System.Drawing.Color]::Red
        }
    }
    catch {
        $previewTextBox.AppendText("ファイル分析中にエラーが発生しました: $_`r`n")
        Update-Status "エラーが発生しました" [System.Drawing.Color]::Red
    }
}

# 文字コードを変換する関数
function FixEncoding {
    $filePath = $fileTextBox.Text
    
    if (-not (Test-Path $filePath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません: $filePath", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $sourceEncoding = $sourceEncodingComboBox.SelectedItem
    if ($sourceEncoding -eq "自動検出") {
        if ($script:detectedEncoding) {
            $sourceEncoding = $script:detectedEncoding
        }
        else {
            # 自動検出が選択されているが、まだ検出していない場合
            AnalyzeFile
            $sourceEncoding = $script:detectedEncoding
        }
    }
    
    $targetEncoding = $targetEncodingComboBox.SelectedItem
    
    $outputPath = "$filePath.converted"
    
    $previewTextBox.Clear()
    Update-Status "文字コード変換中..." [System.Drawing.Color]::DarkBlue
    $previewTextBox.AppendText("文字コード変換中: $filePath`r`n")
    $previewTextBox.AppendText("元の文字コード: $sourceEncoding`r`n")
    $previewTextBox.AppendText("変換後の文字コード: $targetEncoding`r`n`r`n")
    
    try {
        # ソースエンコーディングとターゲットエンコーディングのオブジェクトを取得
        $sourceEncodingObj = if (Get-Command -Name Get-EncodingObject -ErrorAction SilentlyContinue) {
            Get-EncodingObject -EncodingName $sourceEncoding
        } else {
            # 手動でエンコーディングオブジェクトを作成
            switch ($sourceEncoding) {
                "UTF8" { [System.Text.Encoding]::UTF8 }
                "UTF8BOM" { New-Object System.Text.UTF8Encoding($true) }
                "SJIS" { [System.Text.Encoding]::GetEncoding("shift-jis") }
                "EUCJP" { [System.Text.Encoding]::GetEncoding("euc-jp") }
                "JIS" { [System.Text.Encoding]::GetEncoding("iso-2022-jp") }
                "ASCII" { [System.Text.Encoding]::ASCII }
                "Unicode" { [System.Text.Encoding]::Unicode }
                default { [System.Text.Encoding]::UTF8 }
            }
        }
        
        $targetEncodingObj = if (Get-Command -Name Get-EncodingObject -ErrorAction SilentlyContinue) {
            Get-EncodingObject -EncodingName $targetEncoding
        } else {
            # 手動でエンコーディングオブジェクトを作成
            switch ($targetEncoding) {
                "UTF8" { [System.Text.Encoding]::UTF8 }
                "UTF8BOM" { New-Object System.Text.UTF8Encoding($true) }
                "SJIS" { [System.Text.Encoding]::GetEncoding("shift-jis") }
                "EUCJP" { [System.Text.Encoding]::GetEncoding("euc-jp") }
                "JIS" { [System.Text.Encoding]::GetEncoding("iso-2022-jp") }
                "ASCII" { [System.Text.Encoding]::ASCII }
                "Unicode" { [System.Text.Encoding]::Unicode }
                default { [System.Text.Encoding]::UTF8 }
            }
        }
        
        # ファイルを読み込み、内容を取得
        $content = [System.IO.File]::ReadAllText($filePath, $sourceEncodingObj)
        
        # 元のファイルのバックアップを作成
        $backupPath = "$filePath.backup"
        Copy-Item -Path $filePath -Destination $backupPath -Force
        
        # 新しいエンコーディングで内容を書き込み
        [System.IO.File]::WriteAllText($filePath, $content, $targetEncodingObj)
        
        $previewTextBox.AppendText("変換が完了しました。`r`n")
        $previewTextBox.AppendText("元のファイルは $backupPath にバックアップされました。`r`n`r`n")
        
        # 変換後のファイルを表示
        $convertedContent = [System.IO.File]::ReadAllText($filePath, $targetEncodingObj)
        $convertedLines = $convertedContent -split "`r`n" | Select-Object -First 50
        $convertedPreview = [string]::Join("`r`n", $convertedLines)
        
        $previewTextBox.AppendText("変換後のファイル内容プレビュー (最初の50行):`r`n")
        $previewTextBox.AppendText("--------------------------------------------------`r`n")
        $previewTextBox.AppendText($convertedPreview)
        
        if (($convertedContent -split "`r`n").Length -gt 50) {
            $previewTextBox.AppendText("`r`n--------------------------------------------------`r`n")
            $previewTextBox.AppendText("(以下省略)")
        }
    }
    catch {
        $previewTextBox.AppendText("エラーが発生しました: $_`r`n")
        $previewTextBox.AppendText("詳細: $($_.Exception.Message)`r`n")
        $previewTextBox.AppendText("スタックトレース: $($_.ScriptStackTrace)`r`n")
    }
}

# フォームを表示
[void]$form.ShowDialog()
