﻿# PowerShellスクリプト文字化け修正ツール
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

# エンコーディングオプショングループボックス
$encodingGroupBox = New-Object System.Windows.Forms.GroupBox
$encodingGroupBox.Location = New-Object System.Drawing.Point(20, 270)
$encodingGroupBox.Size = New-Object System.Drawing.Size(660, 60)
$encodingGroupBox.Text = "エンコーディングオプション"
$form.Controls.Add($encodingGroupBox)

# BOMオプションラジオボタン
$bomRadioButton = New-Object System.Windows.Forms.RadioButton
$bomRadioButton.Location = New-Object System.Drawing.Point(20, 25)
$bomRadioButton.Size = New-Object System.Drawing.Size(300, 20)
$bomRadioButton.Text = "UTF-8 BOM付き（PowerShellスクリプト推奨）"
$bomRadioButton.Checked = $true
$encodingGroupBox.Controls.Add($bomRadioButton)

$noBomRadioButton = New-Object System.Windows.Forms.RadioButton
$noBomRadioButton.Location = New-Object System.Drawing.Point(330, 25)
$noBomRadioButton.Size = New-Object System.Drawing.Size(300, 20)
$noBomRadioButton.Text = "UTF-8 BOMなし（バッチファイル推奨）"
$encodingGroupBox.Controls.Add($noBomRadioButton)

# ファイルリストボックス
$fileListBox = New-Object System.Windows.Forms.ListBox
$fileListBox.Location = New-Object System.Drawing.Point(20, 390)
$fileListBox.Size = New-Object System.Drawing.Size(660, 180)
$fileListBox.SelectionMode = "MultiExtended"
$form.Controls.Add($fileListBox)

# ファイルリストラベル
$fileListLabel = New-Object System.Windows.Forms.Label
$fileListLabel.Location = New-Object System.Drawing.Point(20, 370)
$fileListLabel.Size = New-Object System.Drawing.Size(660, 20)
$fileListLabel.Text = "検出されたファイル（変換するファイルを選択）:"
$form.Controls.Add($fileListLabel)

# スキャンボタン
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(20, 340)
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
$clearSelectionButton.Location = New-Object System.Drawing.Point(230, 340)
$clearSelectionButton.Size = New-Object System.Drawing.Size(150, 30)
$clearSelectionButton.Text = "選択解除"
$clearSelectionButton.Add_Click({
    $fileListBox.ClearSelected()
})
$form.Controls.Add($clearSelectionButton)

# 全選択ボタン
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(390, 340)
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
$convertButton.Location = New-Object System.Drawing.Point(20, 580)
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
    
    # バッチファイルが含まれているかチェック
    $hasBatchFiles = $false
    foreach ($filePath in $selectedFiles) {
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        if ($extension -eq '.bat' -or $extension -eq '.cmd') {
            $hasBatchFiles = $true
            break
        }
    }
    
    $message = "$($selectedFiles.Count) 個のファイルのエンコーディングを変換します。"
    
    if ($hasBatchFiles) {
        $message = $message + "`r`n注意: バッチファイル(.bat/.cmd)は変換せず、元のエンコーディングを維持します。"
    } else {
        $message = $message + "`r`nファイルはUTF-8(BOMなし)で保存されます。"
    }
    
    $message = $message + "`r`nよろしいですか？"
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "確認",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $successCount = 0
            $errorCount = 0
            $skippedCount = 0
            
            foreach ($filePath in $selectedFiles) {
                try {
                    # ファイルの拡張子を取得
                    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
                    
                    # バッチファイルの場合はスキップ
                    if ($extension -eq '.bat' -or $extension -eq '.cmd') {
                        $skippedCount++
                        continue
                    }
                    
                    # ファイルのエンコーディングを検出
                    $bytes = [System.IO.File]::ReadAllBytes($filePath)
                    $encoding = $null
                    
                    # BOMの検出
                    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
                        # UTF-8 with BOM
                        $encoding = [System.Text.Encoding]::UTF8
                    } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
                        # UTF-16 LE
                        $encoding = [System.Text.Encoding]::Unicode
                    } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
                        # UTF-16 BE
                        $encoding = [System.Text.Encoding]::BigEndianUnicode
                    } else {
                        # BOMなしの場合、基本的にはShift-JISとして扱う（日本語環境の場合）
                        $encoding = [System.Text.Encoding]::GetEncoding(932) # Shift-JIS
                    }
                    
                    # ファイルの内容を読み取り
                    $content = [System.IO.File]::ReadAllText($filePath, $encoding)
                    
                    # BOMの有無を選択
                    if ($bomRadioButton.Checked) {
                        # BOM付きUTF-8で書き出し
                        $utf8Encoding = New-Object System.Text.UTF8Encoding $true
                        [System.IO.File]::WriteAllText($filePath, $content, $utf8Encoding)
                    } else {
                        # BOMなしUTF-8で書き出し
                        $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
                        [System.IO.File]::WriteAllText($filePath, $content, $utf8NoBomEncoding)
                    }
                    
                    $successCount++
                }
                catch {
                    $errorCount++
                    Write-Error "ファイル '$filePath' の変換に失敗しました: $_"
                }
            }
            
            # 完了メッセージの作成
            $completeMessage = "変換完了！ $successCount 個のファイルを変換しました。"
            
            if ($skippedCount -gt 0) {
                $completeMessage = $completeMessage + "`r`n$skippedCount 個のバッチファイルはスキップされました。"
            }
            
            if ($errorCount -gt 0) {
                $completeMessage = $completeMessage + "`r`n$errorCount 個のファイルの変換に失敗しました。"
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                $completeMessage,
                "完了",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            $statusLabel.Text = "変換完了: 成功=$successCount, スキップ=$skippedCount, 失敗=$errorCount"
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
$statusLabel.Location = New-Object System.Drawing.Point(230, 590)
$statusLabel.Size = New-Object System.Drawing.Size(450, 20)
$statusLabel.Text = "準備完了"
$form.Controls.Add($statusLabel)

# フォーム表示
[void]$form.ShowDialog()