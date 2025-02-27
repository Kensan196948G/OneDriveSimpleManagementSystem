# CharacterEncodingFixer.ps1
# PowerShellスクリプトの文字エンコーディングを修正するGUIツール
# 
# 更新履歴:
# 2025/03/20 - 管理者権限での文字化け問題対応
# 2025/03/15 - 初期バージョン作成

# 文字エンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# メインフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShellスクリプト文字エンコーディング修正ツール"
$form.Size = New-Object System.Drawing.Size(700, 550) # 少し高さ追加
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("メイリオ", 10)

# フォームアイコンの設定
try {
    $iconPath = Join-Path $PSScriptRoot "icon.ico"
    if (Test-Path $iconPath) {
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    }
} catch {
    # アイコン設定エラーは無視
}

# 実行環境情報の表示
$labelRuntime = New-Object System.Windows.Forms.Label
$labelRuntime.Location = New-Object System.Drawing.Point(10, 10)
$labelRuntime.Size = New-Object System.Drawing.Size(660, 20)

# 管理者権限で実行されているかチェック
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$adminStatus = if ($isAdmin) { "管理者権限で実行中" } else { "通常権限で実行中" }

# 現在のコンソールエンコーディング情報
$encodingInfo = [Console]::OutputEncoding.EncodingName
$codePageInfo = [Console]::OutputEncoding.CodePage

$labelRuntime.Text = "$adminStatus - エンコーディング: $encodingInfo (CodePage: $codePageInfo)"
if ($isAdmin) {
    $labelRuntime.ForeColor = [System.Drawing.Color]::Red
}
$form.Controls.Add($labelRuntime)

# ラベルの作成
$labelDescription = New-Object System.Windows.Forms.Label
$labelDescription.Location = New-Object System.Drawing.Point(10, 40)
$labelDescription.Size = New-Object System.Drawing.Size(660, 50)
$labelDescription.Text = "このツールは、PowerShellスクリプトファイルの文字エンコーディングを修正して、文字化け問題を解決します。対象のフォルダまたはファイルを選択してください。"
$form.Controls.Add($labelDescription)

# パス入力テキストボックスの作成
$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Location = New-Object System.Drawing.Point(10, 100)
$textBoxPath.Size = New-Object System.Drawing.Size(580, 25)
$textBoxPath.Text = $PSScriptRoot
$form.Controls.Add($textBoxPath)

# フォルダ選択ボタンの作成
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(600, 100)
$buttonBrowse.Size = New-Object System.Drawing.Size(70, 25)
$buttonBrowse.Text = "参照..."
$buttonBrowse.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.SelectedPath = $textBoxPath.Text
    $folderDialog.Description = "スキャンするフォルダを選択してください"
    $folderDialog.ShowNewFolderButton = $false
    
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxPath.Text = $folderDialog.SelectedPath
    }
})
$form.Controls.Add($buttonBrowse)

# ファイルパターンラベル
$labelPattern = New-Object System.Windows.Forms.Label
$labelPattern.Location = New-Object System.Drawing.Point(10, 140)
$labelPattern.Size = New-Object System.Drawing.Size(200, 25)
$labelPattern.Text = "ファイルパターン: "
$form.Controls.Add($labelPattern)

# ファイルパターンテキストボックス
$textBoxPattern = New-Object System.Windows.Forms.TextBox
$textBoxPattern.Location = New-Object System.Drawing.Point(150, 140)
$textBoxPattern.Size = New-Object System.Drawing.Size(520, 25)
$textBoxPattern.Text = "*.ps1"
$form.Controls.Add($textBoxPattern)

# エンコーディングオプションのグループボックス
$groupBoxEncoding = New-Object System.Windows.Forms.GroupBox
$groupBoxEncoding.Location = New-Object System.Drawing.Point(10, 180)
$groupBoxEncoding.Size = New-Object System.Drawing.Size(660, 100)
$groupBoxEncoding.Text = "エンコーディングオプション"
$form.Controls.Add($groupBoxEncoding)

# BOMありUTF-8ラジオボタン
$radioUTF8BOM = New-Object System.Windows.Forms.RadioButton
$radioUTF8BOM.Location = New-Object System.Drawing.Point(20, 30)
$radioUTF8BOM.Size = New-Object System.Drawing.Size(300, 25)
$radioUTF8BOM.Text = "UTF-8 with BOM (PowerShell推奨)"
$radioUTF8BOM.Checked = $true
$groupBoxEncoding.Controls.Add($radioUTF8BOM)

# BOMなしUTF-8ラジオボタン
$radioUTF8NoBOM = New-Object System.Windows.Forms.RadioButton
$radioUTF8NoBOM.Location = New-Object System.Drawing.Point(20, 60)
$radioUTF8NoBOM.Size = New-Object System.Drawing.Size(300, 25)
$radioUTF8NoBOM.Text = "UTF-8 without BOM"
$groupBoxEncoding.Controls.Add($radioUTF8NoBOM)

# リストボックスの作成
$listBoxFiles = New-Object System.Windows.Forms.ListBox
$listBoxFiles.Location = New-Object System.Drawing.Point(10, 290)
$listBoxFiles.Size = New-Object System.Drawing.Size(660, 150)
$form.Controls.Add($listBoxFiles)

# ステータスラベル
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10, 450)
$labelStatus.Size = New-Object System.Drawing.Size(660, 25)
$labelStatus.Text = "準備完了"
$form.Controls.Add($labelStatus)

# スキャンボタンの作成
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Location = New-Object System.Drawing.Point(10, 480)
$buttonScan.Size = New-Object System.Drawing.Size(150, 30)
$buttonScan.Text = "ファイルをスキャン"
$buttonScan.Add_Click({
    $listBoxFiles.Items.Clear()
    $path = $textBoxPath.Text
    $pattern = $textBoxPattern.Text
    
    if (-not (Test-Path $path)) {
        [System.Windows.Forms.MessageBox]::Show("指定されたパスが存在しません。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $labelStatus.Text = "スキャン中..."
    $files = Get-ChildItem -Path $path -Filter $pattern -Recurse
    
    if ($files.Count -eq 0) {
        $labelStatus.Text = "該当するファイルが見つかりませんでした。"
    } else {
        $labelStatus.Text = "$($files.Count) 個のファイルが見つかりました。"
        foreach ($file in $files) {
            $listBoxFiles.Items.Add($file.FullName)
        }
    }
})
$form.Controls.Add($buttonScan)

# 変換ボタンの作成
$buttonConvert = New-Object System.Windows.Forms.Button
$buttonConvert.Location = New-Object System.Drawing.Point(170, 480)
$buttonConvert.Size = New-Object System.Drawing.Size(150, 30)
$buttonConvert.Text = "エンコーディング変換"
$buttonConvert.Add_Click({
    if ($listBoxFiles.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("変換するファイルがありません。先にスキャンを実行してください。", "警告", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $encoding = if ($radioUTF8BOM.Checked) { New-Object System.Text.UTF8Encoding($true) } else { New-Object System.Text.UTF8Encoding($false) }
    $converted = 0
    $errors = 0
    
    foreach ($filePath in $listBoxFiles.Items) {
        try {
            $labelStatus.Text = "処理中: $filePath"
            $labelStatus.Refresh()
            
            # テキストを読み込み
            $content = [System.IO.File]::ReadAllText($filePath)
            # UTF-8で書き直し
            [System.IO.File]::WriteAllText($filePath, $content, $encoding)
            $converted++
        }
        catch {
            $errors++
            [System.Windows.Forms.MessageBox]::Show("ファイルの処理中にエラーが発生しました: $filePath`n`n$($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    
    $labelStatus.Text = "処理完了: $converted 個のファイルを変換しました。 $errors 個のエラーが発生しました。"
    
    # 処理結果に応じてメッセージを表示
    $message = "処理が完了しました。`n`n変換されたファイル数: $converted`nエラー数: $errors"
    $icon = if ($errors -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information }
    [System.Windows.Forms.MessageBox]::Show($message, "処理完了", [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
})
$form.Controls.Add($buttonConvert)

# 終了ボタンの作成
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(530, 480)
$buttonExit.Size = New-Object System.Drawing.Size(120, 30)
$buttonExit.Text = "終了"
$buttonExit.Add_Click({
    $form.Close()
})
$form.Controls.Add($buttonExit)

# エンコーディング情報ボタンの作成
$buttonInfo = New-Object System.Windows.Forms.Button
$buttonInfo.Location = New-Object System.Drawing.Point(330, 480)
$buttonInfo.Size = New-Object System.Drawing.Size(190, 30)
$buttonInfo.Text = "エンコーディング情報"
$buttonInfo.Add_Click({
    # 選択されているファイルのエンコーディング情報を表示
    if ($listBoxFiles.SelectedItem) {
        $selectedFile = $listBoxFiles.SelectedItem.ToString()
        try {
            # ファイルを読み込み
            $bytes = [System.IO.File]::ReadAllBytes($selectedFile)
            
            # BOMをチェック
            $hasBOM = $false
            $encoding = "不明"
            
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
                $encoding = "UTF-8 with BOM"
                $hasBOM = $true
            } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
                $encoding = "UTF-16 LE (Little Endian)"
            } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
                $encoding = "UTF-16 BE (Big Endian)"
            } else {
                # BOMなしの場合は内容から推測
                $content = [System.IO.File]::ReadAllText($selectedFile)
                if ([System.Text.Encoding]::UTF8.GetByteCount($content) -eq $bytes.Length) {
                    $encoding = "UTF-8 without BOM（推測）"
                } else {
                    $encoding = "Shift-JIS または他のエンコーディング（推測）"
                }
            }
            
            # ファイルサイズも取得
            $fileSize = (Get-Item $selectedFile).Length
            
            [System.Windows.Forms.MessageBox]::Show(
                "ファイル名: $([System.IO.Path]::GetFileName($selectedFile))`n" +
                "パス: $selectedFile`n" +
                "サイズ: $fileSize バイト`n" +
                "エンコーディング: $encoding`n",
                "ファイルエンコーディング情報",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show("ファイルの読み取り中にエラーが発生しました: $($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("ファイルが選択されていません。情報を表示するファイルをリストから選択してください。", "選択エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$form.Controls.Add($buttonInfo)

# フォームの表示
$form.ShowDialog()
