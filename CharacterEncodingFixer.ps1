#
# 文字化け診断・修正GUIツール
# 
# PowerShellスクリプトの文字コード問題をGUIで診断・修正するツール

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# フォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "文字コード変換ツール"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)

# ファイル選択ラベル
$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Location = New-Object System.Drawing.Point(20, 20)
$fileLabel.Size = New-Object System.Drawing.Size(150, 20)
$fileLabel.Text = "変換するファイル:"
$form.Controls.Add($fileLabel)

# ファイルパステキストボックス
$fileTextBox = New-Object System.Windows.Forms.TextBox
$fileTextBox.Location = New-Object System.Drawing.Point(20, 45)
$fileTextBox.Size = New-Object System.Drawing.Size(500, 20)
$fileTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($fileTextBox)

# 参照ボタン
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(530, 43)
$browseButton.Size = New-Object System.Drawing.Size(100, 25)
$browseButton.Text = "参照..."
$browseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PowerShellスクリプト (*.ps1)|*.ps1|すべてのファイル (*.*)|*.*"
    $openFileDialog.Title = "変換するファイルを選択"

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileTextBox.Text = $openFileDialog.FileName
        AnalyzeFile
    }
})
$form.Controls.Add($browseButton)

# 元のエンコーディングラベル
$sourceEncodingLabel = New-Object System.Windows.Forms.Label
$sourceEncodingLabel.Location = New-Object System.Drawing.Point(20, 80)
$sourceEncodingLabel.Size = New-Object System.Drawing.Size(150, 20)
$sourceEncodingLabel.Text = "元の文字コード:"
$form.Controls.Add($sourceEncodingLabel)

# 元のエンコーディングコンボボックス
$sourceEncodingComboBox = New-Object System.Windows.Forms.ComboBox
$sourceEncodingComboBox.Location = New-Object System.Drawing.Point(170, 80)
$sourceEncodingComboBox.Size = New-Object System.Drawing.Size(150, 20)
$sourceEncodingComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$sourceEncodingComboBox.Items.AddRange(@("auto", "utf-8", "utf-16", "shift-jis", "euc-jp", "ascii"))
$sourceEncodingComboBox.SelectedIndex = 0
$form.Controls.Add($sourceEncodingComboBox)

# 変換後のエンコーディングラベル
$targetEncodingLabel = New-Object System.Windows.Forms.Label
$targetEncodingLabel.Location = New-Object System.Drawing.Point(350, 80)
$targetEncodingLabel.Size = New-Object System.Drawing.Size(150, 20)
$targetEncodingLabel.Text = "変換後の文字コード:"
$form.Controls.Add($targetEncodingLabel)

# 変換後のエンコーディングコンボボックス
$targetEncodingComboBox = New-Object System.Windows.Forms.ComboBox
$targetEncodingComboBox.Location = New-Object System.Drawing.Point(500, 80)
$targetEncodingComboBox.Size = New-Object System.Drawing.Size(150, 20)
$targetEncodingComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$targetEncodingComboBox.Items.AddRange(@("utf-8", "utf-16", "shift-jis", "euc-jp", "ascii"))
$targetEncodingComboBox.SelectedIndex = 0
$form.Controls.Add($targetEncodingComboBox)

# BOMチェックボックス
$bomCheckBox = New-Object System.Windows.Forms.CheckBox
$bomCheckBox.Location = New-Object System.Drawing.Point(500, 110)
$bomCheckBox.Size = New-Object System.Drawing.Size(150, 20)
$bomCheckBox.Text = "BOMを含める"
$bomCheckBox.Checked = $true
$form.Controls.Add($bomCheckBox)

# 分析ボタン
$analyzeButton = New-Object System.Windows.Forms.Button
$analyzeButton.Location = New-Object System.Drawing.Point(20, 110)
$analyzeButton.Size = New-Object System.Drawing.Size(120, 30)
$analyzeButton.Text = "ファイル分析"
$analyzeButton.Add_Click({ AnalyzeFile })
$form.Controls.Add($analyzeButton)

# 変換ボタン
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(150, 110)
$convertButton.Size = New-Object System.Drawing.Size(120, 30)
$convertButton.Text = "変換実行"
$convertButton.Add_Click({ ConvertFile })
$form.Controls.Add($convertButton)

# 結果表示用リッチテキストボックス
$resultTextBox = New-Object System.Windows.Forms.RichTextBox
$resultTextBox.Location = New-Object System.Drawing.Point(20, 150)
$resultTextBox.Size = New-Object System.Drawing.Size(640, 350)
$resultTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$resultTextBox.ReadOnly = $true
$resultTextBox.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Controls.Add($resultTextBox)

# ステータスラベル
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 510)
$statusLabel.Size = New-Object System.Drawing.Size(640, 40)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($statusLabel)

# ファイル分析関数
function AnalyzeFile {
    $resultTextBox.Clear()
    $statusLabel.Text = "ファイル分析中..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    
    $filePath = $fileTextBox.Text
    
    if (-not (Test-Path $filePath)) {
        $statusLabel.Text = "エラー: ファイルが見つかりません。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    # 文字コード自動検出
    $encodingInfo = Detect-FileEncoding -FilePath $filePath
    
    # 結果表示
    $resultTextBox.AppendText("ファイル: " + $filePath + "`r`n")
    $resultTextBox.AppendText("検出された文字コード: " + $encodingInfo.Encoding + " (信頼度: " + $encodingInfo.Confidence + "%)`r`n")
    $resultTextBox.AppendText("BOM: " + $encodingInfo.BOM + "`r`n`r`n")
    
    # 文字コードをセット
    if ($sourceEncodingComboBox.SelectedItem -eq "auto") {
        # 検出された文字コードを自動選択
        for ($i = 0; $i -lt $sourceEncodingComboBox.Items.Count; $i++) {
            if ($sourceEncodingComboBox.Items[$i] -eq $encodingInfo.Encoding) {
                $sourceEncodingComboBox.SelectedIndex = $i
                break
            }
        }
    }
    
    # ファイルの内容をプレビュー表示
    try {
        $srcEncoding = switch ($encodingInfo.Encoding.ToLower()) {
            "utf-8"     { [System.Text.Encoding]::UTF8 }
            "utf-16"    { [System.Text.Encoding]::Unicode }
            "utf-16be"  { [System.Text.Encoding]::BigEndianUnicode }
            "shift-jis" { [System.Text.Encoding]::GetEncoding(932) }
            "euc-jp"    { [System.Text.Encoding]::GetEncoding(51932) }
            "ascii"     { [System.Text.Encoding]::ASCII }
            default     { [System.Text.Encoding]::UTF8 }
        }
        
        $content = [System.IO.File]::ReadAllText($filePath, $srcEncoding)
        $contentLines = $content -split "`r?`n"
        
        $resultTextBox.AppendText("ファイル内容プレビュー (最初の20行):" + "`r`n")
        $resultTextBox.AppendText("--------------------------------------------------" + "`r`n")
        
        $previewLines = [Math]::Min(20, $contentLines.Length)
        for ($i = 0; $i -lt $previewLines; $i++) {
            $resultTextBox.AppendText($contentLines[$i] + "`r`n")
        }
        
        $resultTextBox.AppendText("--------------------------------------------------" + "`r`n")
        $statusLabel.Text = "ファイル分析が完了しました。変換を実行するには [変換実行] ボタンをクリックしてください。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    }
    catch {
        $resultTextBox.AppendText("エラー: ファイル内容の読み取り中にエラーが発生しました。" + "`r`n")
        $resultTextBox.AppendText($_.Exception.Message + "`r`n")
        $statusLabel.Text = "エラーが発生しました。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
}

# ファイル変換関数
function ConvertFile {
    $filePath = $fileTextBox.Text
    $sourceEncoding = $sourceEncodingComboBox.SelectedItem
    $targetEncoding = $targetEncodingComboBox.SelectedItem
    $useBom = $bomCheckBox.Checked
    
    if (-not (Test-Path $filePath)) {
        $statusLabel.Text = "エラー: ファイルが見つかりません。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    try {
        # 入力エンコーディング
        $srcEncoding = switch ($sourceEncoding.ToLower()) {
            "utf-8"     { [System.Text.Encoding]::UTF8 }
            "utf-16"    { [System.Text.Encoding]::Unicode }
            "utf-16be"  { [System.Text.Encoding]::BigEndianUnicode }
            "shift-jis" { [System.Text.Encoding]::GetEncoding(932) }
            "euc-jp"    { [System.Text.Encoding]::GetEncoding(51932) }
            "ascii"     { [System.Text.Encoding]::ASCII }
            default     { [System.Text.Encoding]::UTF8 }
        }
        
        # 出力エンコーディング
        $dstEncoding = switch ($targetEncoding.ToLower()) {
            "utf-8"     { New-Object System.Text.UTF8Encoding($useBom) }
            "utf-16"    { [System.Text.Encoding]::Unicode }
            "utf-16be"  { [System.Text.Encoding]::BigEndianUnicode }
            "shift-jis" { [System.Text.Encoding]::GetEncoding(932) }
            "euc-jp"    { [System.Text.Encoding]::GetEncoding(51932) }
            "ascii"     { [System.Text.Encoding]::ASCII }
            default     { New-Object System.Text.UTF8Encoding($useBom) }
        }
        
        # ファイル内容を読み込む
        $content = [System.IO.File]::ReadAllText($filePath, $srcEncoding)
        
        # バックアップファイル作成
        $backupFile = "$filePath.bak"
        $counter = 1
        while (Test-Path $backupFile) {
            $backupFile = "$filePath.bak$counter"
            $counter++
        }
        
        Copy-Item -Path $filePath -Destination $backupFile
        
        # 変換して書き込み
        [System.IO.File]::WriteAllText($filePath, $content, $dstEncoding)
        
        $resultTextBox.AppendText("`r`n変換が完了しました:" + "`r`n")
        $resultTextBox.AppendText("  変換前: $sourceEncoding`r`n")
        $resultTextBox.AppendText("  変換後: $targetEncoding (BOM: $useBom)`r`n")
        $resultTextBox.AppendText("  バックアップ: $backupFile`r`n")
        
        $statusLabel.Text = "変換が正常に完了しました。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    }
    catch {
        $resultTextBox.AppendText("エラー: 変換処理中にエラーが発生しました。" + "`r`n")
        $resultTextBox.AppendText($_.Exception.Message + "`r`n")
        $statusLabel.Text = "変換エラーが発生しました。"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
}

# 文字コード検出関数
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
    
    # 文字コードの推測（簡易版）
    $content = $bytes
    
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
    
    # UTF-8 チェック
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
    
    # 日本語環境ではShift-JISが最も一般的
    return @{Encoding = "shift-jis"; BOM = $false; Confidence = 60}
}

# フォームの表示
$form.ShowDialog()
