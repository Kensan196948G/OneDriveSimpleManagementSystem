#
# PowerShellスクリプト文字化け修正ツール
# 指定されたフォルダ内のスクリプトファイルをUTF-8に変換します
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUIフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShellスクリプト文字化け修正ツール"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("メイリオ", 9)

# フォームアイコンの設定
try {
    $iconPath = Join-Path $PSScriptRoot "icon.ico"
    if (Test-Path $iconPath) {
        $form.Icon = New-Object System.Drawing.Icon($iconPath)
    }
} catch {
    # アイコン設定エラーは無視
}

# フォーム上部の説明ラベル
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 20)
$descriptionLabel.Size = New-Object System.Drawing.Size(660, 60)
$descriptionLabel.Text = "このツールは、PowerShellスクリプトの文字化けを修正します。`r`n指定されたフォルダ内のスクリプトファイルをUTF-8エンコーディングに変換します。"
$form.Controls.Add($descriptionLabel)

# フォルダ選択グループボックス
$folderGroupBox = New-Object System.Windows.Forms.GroupBox
$folderGroupBox.Location = New-Object System.Drawing.Point(20, 90)
$folderGroupBox.Size = New-Object System.Drawing.Size(660, 80)
$folderGroupBox.Text = "対象フォルダ"
$form.Controls.Add($folderGroupBox)

# フォルダパステキストボックス
$folderTextBox = New-Object System.Windows.Forms.TextBox
$folderTextBox.Location = New-Object System.Drawing.Point(20, 30)
$folderTextBox.Size = New-Object System.Drawing.Size(500, 25)
$folderTextBox.Text = "C:\kitting\OneDrive運用ツール"
$folderGroupBox.Controls.Add($folderTextBox)

# フォルダ参照ボタン
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(530, 28)
$browseButton.Size = New-Object System.Drawing.Size(110, 30)
$browseButton.Text = "参照..."
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "スクリプトファイルがあるフォルダを選択してください"
    $folderBrowser.SelectedPath = $folderTextBox.Text
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $folderTextBox.Text = $folderBrowser.SelectedPath
    }
})
$folderGroupBox.Controls.Add($browseButton)

# ファイルパターングループボックス
$patternGroupBox = New-Object System.Windows.Forms.GroupBox
$patternGroupBox.Location = New-Object System.Drawing.Point(20, 180)
$patternGroupBox.Size = New-Object System.Drawing.Size(660, 80)
$patternGroupBox.Text = "ファイルパターン（カンマ区切りで複数指定可）"
$form.Controls.Add($patternGroupBox)

# パターンテキストボックス
$patternTextBox = New-Object System.Windows.Forms.TextBox
$patternTextBox.Location = New-Object System.Drawing.Point(20, 30)
$patternTextBox.Size = New-Object System.Drawing.Size(620, 25)
$patternTextBox.Text = "*.ps1,*.psm1,*.bat,*.cmd"
$patternGroupBox.Controls.Add($patternTextBox)

# ファイルリスト表示用リストボックス
$fileListBox = New-Object System.Windows.Forms.ListBox
$fileListBox.Location = New-Object System.Drawing.Point(20, 320)
$fileListBox.Size = New-Object System.Drawing.Size(660, 180)
$fileListBox.SelectionMode = "MultiExtended"
$form.Controls.Add($fileListBox)

# ファイルリストラベル
$fileListLabel = New-Object System.Windows.Forms.Label
$fileListLabel.Location = New-Object System.Drawing.Point(20, 300)
$fileListLabel.Size = New-Object System.Drawing.Size(660, 20)
$fileListLabel.Text = "検出されたファイル (変換するファイルを選択):"
$form.Controls.Add($fileListLabel)

# スキャンボタン
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(20, 270)
$scanButton.Size = New-Object System.Drawing.Size(200, 30)
$scanButton.Text = "ファイルをスキャン"
$scanButton.BackColor = [System.Drawing.Color]::LightBlue
$scanButton.Add_Click({
    $fileListBox.Items.Clear()
    
    try {
        $folder = $folderTextBox.Text
        $patterns = $patternTextBox.Text -split ','
        
        if (-not (Test-Path $folder -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show("指定されたフォルダが存在しません: $folder", "エラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $statusLabel.Text = "ファイルをスキャンしています..."
        $form.Refresh()
        
        $files = @()
        foreach ($pattern in $patterns) {
            $patternFiles = Get-ChildItem -Path $folder -Filter $pattern.Trim() -Recurse -File
            $files += $patternFiles
        }
        
        if ($files.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("指定されたパターンに一致するファイルが見つかりませんでした。", 
                "情報", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $statusLabel.Text = "ファイルが見つかりませんでした。"
            return
        }
        
        foreach ($file in $files) {
            $fileListBox.Items.Add($file.FullName)
        }
        
        $statusLabel.Text = "$($files.Count) 件のファイルが見つかりました。"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("スキャン中にエラーが発生しました: $_", "エラー", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "エラーが発生しました。"
    }
})
$form.Controls.Add($scanButton)

# 選択解除ボタン
$clearSelectionButton = New-Object System.Windows.Forms.Button
$clearSelectionButton.Location = New-Object System.Drawing.Point(230, 270)
$clearSelectionButton.Size = New-Object System.Drawing.Size(150, 30)
$clearSelectionButton.Text = "選択解除"
$clearSelectionButton.Add_Click({
    $fileListBox.ClearSelected()
})
$form.Controls.Add($clearSelectionButton)

# 全選択ボタン
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(390, 270)
$selectAllButton.Size = New-Object System.Drawing.Size(150, 30)
$selectAllButton.Text = "全選択"
$selectAllButton.Add_Click({
    for ($i = 0; $i -lt $fileListBox.Items.Count; $i++) {
        $fileListBox.SetSelected($i, $true)
    }
})
$form.Controls.Add($selectAllButton)

# 変換ボタン
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(20, 510)
$convertButton.Size = New-Object System.Drawing.Size(200, 40)
$convertButton.Text = "エンコーディング変換"
$convertButton.BackColor = [System.Drawing.Color]::LightGreen
$convertButton.Add_Click({
    $selectedFiles = $fileListBox.SelectedItems
    
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("変換するファイルを選択してください。", 
            "警告", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "$($selectedFiles.Count) 個のファイルをUTF-8に変換します。よろしいですか？",
        "確認",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $successCount = 0
            $errorCount = 0
            
            foreach ($filePath in $selectedFiles) {
                try {
                    $encoding = Get-FileEncoding $filePath
                    
                    # ファイルの内容を読み取り
                    $content = [System.IO.File]::ReadAllText($filePath, $encoding)
                    
                    # BOMなしUTF-8で書き出し
                    $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
                    [System.IO.File]::WriteAllText($filePath, $content, $utf8NoBomEncoding)
                    
                    $successCount++
                }
                catch {
                    $errorCount++
                    Write-Error "ファイル '$filePath' の変換に失敗しました: $_"
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "変換完了！ $successCount 個のファイルを変換しました。$(if ($errorCount -gt 0) { "`r`n$errorCount 個のファイルの変換に失敗しました。" })",
                "完了",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            $statusLabel.Text = "変換完了: 成功=$successCount, 失敗=$errorCount"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "変換中にエラーが発生しました: $_",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $statusLabel.Text = "エラーが発生しました。"
        }
    }
})
$form.Controls.Add($convertButton)

# ステータスラベル
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(230, 520)
$statusLabel.Size = New-Object System.Drawing.Size(450, 20)
$statusLabel.Text = "準備完了"
$form.Controls.Add($statusLabel)

# ファイルエンコーディング検出関数
function Get-FileEncoding {
    param (
        [Parameter(Mandatory=$true)]
        [string] $FilePath
    )
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        
        # BOM検出
        if ($bytes.Length -ge 2) {
            # UTF-16 LE
            if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
                return [System.Text.Encoding]::Unicode
            }
            
            # UTF-16 BE
            if ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
                return [System.Text.Encoding]::BigEndianUnicode
            }
            
            # UTF-8 with BOM
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
                return [System.Text.Encoding]::UTF8
            }
        }
        
        # BOMが無い場合
        # エンコーディング推測（単純化）
        $encoding = [System.Text.Encoding]::GetEncoding('shift-jis')
        
        # ASCII/UTF-8/SHIFT-JIS推測
        $content = $encoding.GetString($bytes)
        if ($content.Contains('�')) {
            # Shift-JISで文字化けしたらUTF-8の可能性
            return [System.Text.Encoding]::UTF8
        }
        
        # デフォルトはShift-JIS
        return $encoding
    }
    catch {
        Write-Error "エンコーディングの検出に失敗しました: $_"
        # デフォルトはShift-JIS
        return [System.Text.Encoding]::GetEncoding('shift-jis')
    }
}

# フォーム表示
[void]$form.ShowDialog()
