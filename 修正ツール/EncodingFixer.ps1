# PowerShellスクリプト文字化け修正ツール
# This tool converts script files in the specified folder to UTF-8 encoding

# Add assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'PowerShellスクリプト文字化け修正ツール'
$form.Size = New-Object System.Drawing.Size(700, 680)  # Increased height for new controls
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font('MS Gothic', 9)

# Set form icon
try {
    $iconPath = Join-Path $PSScriptRoot 'icon.ico'
    if (Test-Path $iconPath) {
        $form.Icon = New-Object System.Drawing.Icon($iconPath)
    }
} catch {
    # Ignore icon setting errors
}

# Description label at the top of the form
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 20)
$descriptionLabel.Size = New-Object System.Drawing.Size(660, 60)
$descriptionLabel.Text = 'このツールは、PowerShellスクリプトの文字化けを修正します。
指定されたフォルダ内のスクリプトファイルをUTF-8エンコーディングに変換します。'
$form.Controls.Add($descriptionLabel)

# Folder selection group box
$folderGroupBox = New-Object System.Windows.Forms.GroupBox
$folderGroupBox.Location = New-Object System.Drawing.Point(20, 90)
$folderGroupBox.Size = New-Object System.Drawing.Size(660, 80)
$folderGroupBox.Text = '対象フォルダ'
$form.Controls.Add($folderGroupBox)

# Folder path text box
$folderTextBox = New-Object System.Windows.Forms.TextBox
$folderTextBox.Location = New-Object System.Drawing.Point(20, 30)
$folderTextBox.Size = New-Object System.Drawing.Size(500, 25)
$folderTextBox.Text = 'C:\kitting\OneDrive運用ツール'
$folderGroupBox.Controls.Add($folderTextBox)

# Browse button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(530, 28)
$browseButton.Size = New-Object System.Drawing.Size(110, 30)
$browseButton.Text = '参照...'
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = 'スクリプトファイルがあるフォルダを選択してください'
    $folderBrowser.SelectedPath = $folderTextBox.Text
    
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $folderTextBox.Text = $folderBrowser.SelectedPath
    }
})
$folderGroupBox.Controls.Add($browseButton)

# File pattern group box
$patternGroupBox = New-Object System.Windows.Forms.GroupBox
$patternGroupBox.Location = New-Object System.Drawing.Point(20, 180)
$patternGroupBox.Size = New-Object System.Drawing.Size(660, 80)
$patternGroupBox.Text = 'ファイルパターン（カンマ区切りで複数指定可）'
$form.Controls.Add($patternGroupBox)

# Pattern text box
$patternTextBox = New-Object System.Windows.Forms.TextBox
$patternTextBox.Location = New-Object System.Drawing.Point(20, 30)
$patternTextBox.Size = New-Object System.Drawing.Size(620, 25)
$patternTextBox.Text = '*.ps1,*.psm1,*.bat,*.cmd'
$patternGroupBox.Controls.Add($patternTextBox)

# Encoding options group box
$encodingGroupBox = New-Object System.Windows.Forms.GroupBox
$encodingGroupBox.Location = New-Object System.Drawing.Point(20, 270)
$encodingGroupBox.Size = New-Object System.Drawing.Size(660, 60)
$encodingGroupBox.Text = 'エンコーディングオプション'
$form.Controls.Add($encodingGroupBox)

# BOM option radio buttons
$bomRadioButton = New-Object System.Windows.Forms.RadioButton
$bomRadioButton.Location = New-Object System.Drawing.Point(20, 25)
$bomRadioButton.Size = New-Object System.Drawing.Size(300, 20)
$bomRadioButton.Text = 'UTF-8 BOM付き（PowerShellスクリプト推奨）'
$bomRadioButton.Checked = $true
$encodingGroupBox.Controls.Add($bomRadioButton)

$noBomRadioButton = New-Object System.Windows.Forms.RadioButton
$noBomRadioButton.Location = New-Object System.Drawing.Point(330, 25)
$noBomRadioButton.Size = New-Object System.Drawing.Size(300, 20)
$noBomRadioButton.Text = 'UTF-8 BOMなし（バッチファイル推奨）'
$encodingGroupBox.Controls.Add($noBomRadioButton)

# File list box
$fileListBox = New-Object System.Windows.Forms.ListBox
$fileListBox.Location = New-Object System.Drawing.Point(20, 390)
$fileListBox.Size = New-Object System.Drawing.Size(660, 180)
$fileListBox.SelectionMode = 'MultiExtended'
$form.Controls.Add($fileListBox)

# File list label
$fileListLabel = New-Object System.Windows.Forms.Label
$fileListLabel.Location = New-Object System.Drawing.Point(20, 370)
$fileListLabel.Size = New-Object System.Drawing.Size(660, 20)
$fileListLabel.Text = '検出されたファイル（変換するファイルを選択）:'
$form.Controls.Add($fileListLabel)

# Scan button
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(20, 340)
$scanButton.Size = New-Object System.Drawing.Size(200, 30)
$scanButton.Text = 'ファイルをスキャン'
$scanButton.BackColor = [System.Drawing.Color]::LightBlue
$scanButton.Add_Click({
    $fileListBox.Items.Clear()
    
    try {
        $folder = $folderTextBox.Text
        $patterns = $patternTextBox.Text -split ','
        
        if (-not (Test-Path $folder -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show('指定されたフォルダが存在しません: ' + $folder, 'エラー', 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $statusLabel.Text = 'ファイルをスキャンしています...'
        $form.Refresh()
        
        $files = @()
        foreach ($pattern in $patterns) {
            $patternFiles = Get-ChildItem -Path $folder -Filter $pattern.Trim() -Recurse -File
            $files += $patternFiles
        }
        
        if ($files.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('指定されたパターンに一致するファイルが見つかりませんでした。', 
                '情報', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $statusLabel.Text = 'ファイルが見つかりませんでした。'
            return
        }
        
        foreach ($file in $files) {
            $fileListBox.Items.Add($file.FullName)
        }
        
        $statusLabel.Text = $files.Count.ToString() + ' 件のファイルが見つかりました。'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show('スキャン中にエラーが発生しました: ' + $_, 'エラー', 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = 'エラーが発生しました。'
    }
})
$form.Controls.Add($scanButton)

# Clear selection button
$clearSelectionButton = New-Object System.Windows.Forms.Button
$clearSelectionButton.Location = New-Object System.Drawing.Point(230, 340)
$clearSelectionButton.Size = New-Object System.Drawing.Size(150, 30)
$clearSelectionButton.Text = '選択解除'
$clearSelectionButton.Add_Click({
    $fileListBox.ClearSelected()
})
$form.Controls.Add($clearSelectionButton)

# Select all button
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(390, 340)
$selectAllButton.Size = New-Object System.Drawing.Size(150, 30)
$selectAllButton.Text = '全選択'
$selectAllButton.Add_Click({
    for ($i = 0; $i -lt $fileListBox.Items.Count; $i++) {
        $fileListBox.SetSelected($i, $true)
    }
})
$form.Controls.Add($selectAllButton)

# Convert button
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(20, 590)  # Adjusted position
$convertButton.Size = New-Object System.Drawing.Size(200, 40)
$convertButton.Text = 'エンコーディング変換'
$convertButton.BackColor = [System.Drawing.Color]::LightGreen
$convertButton.Add_Click({
    $selectedFiles = $fileListBox.SelectedItems
    
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('変換するファイルを選択してください。', 
            '警告', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Determine encoding type based on radio button selection
    $useBom = $bomRadioButton.Checked
    $encodingType = if ($useBom) { "UTF-8 BOM付き" } else { "UTF-8 BOMなし" }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $selectedFiles.Count.ToString() + ' 個のファイルを' + $encodingType + 'に変換します。この操作は元に戻せません。よろしいですか？',
        '確認',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $successCount = 0
            $errorCount = 0
            
            foreach ($filePath in $selectedFiles) {
                try {
                    # Read file content (auto-detect encoding)
                    $content = Get-Content -Path $filePath -Encoding Default -Raw -ErrorAction Stop
                    
                    if ($useBom) {
                        # Write with UTF-8 with BOM
                        $utf8Encoding = New-Object System.Text.UTF8Encoding $true
                        [System.IO.File]::WriteAllText($filePath, $content, $utf8Encoding)
                    } else {
                        # Write with UTF-8 no BOM
                        $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
                        [System.IO.File]::WriteAllText($filePath, $content, $utf8NoBomEncoding)
                    }
                    
                    $successCount++
                }
                catch {
                    $errorCount++
                    Write-Error ('ファイル ''' + $filePath + ''' の変換に失敗しました: ' + $_)
                }
            }
            
            $message = '変換完了！ ' + $successCount + ' 個のファイルを' + $encodingType + 'に変換しました。'
            if ($errorCount -gt 0) {
                $message = $message + "`r`n" + $errorCount + ' 個のファイルの変換に失敗しました。'
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                $message,
                '完了',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            $statusLabel.Text = '変換完了: 成功=' + $successCount + ', 失敗=' + $errorCount
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                '変換中にエラーが発生しました: ' + $_,
                'エラー',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $statusLabel.Text = 'エラーが発生しました。'
        }
    }
})
$form.Controls.Add($convertButton)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(230, 600)  # Adjusted position
$statusLabel.Size = New-Object System.Drawing.Size(450, 20)
$statusLabel.Text = '準備完了'
$form.Controls.Add($statusLabel)

# Display form
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
