# OneDriveステータスレポートモジュール
# 文字コード: UTF-8 with BOM

function Get-OneDriveStatus {
    Clear-Host
    Write-Host "=====================================================
  OneDriveステータスレポート
=====================================================" -ForegroundColor Cyan
    
    try {
        # OneDriveプロセスの確認
        $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
        if ($oneDriveProcess) {
            Write-Host "OneDriveステータス: " -NoNewline
            Write-Host "実行中" -ForegroundColor Green
        } else {
            Write-Host "OneDriveステータス: " -NoNewline
            Write-Host "停止中" -ForegroundColor Red
        }
        
        # OneDriveフォルダの確認
        $oneDrivePath = [System.Environment]::GetFolderPath("UserProfile") + "\OneDrive"
        if (Test-Path -Path $oneDrivePath) {
            Write-Host "OneDriveフォルダ: " -NoNewline
            Write-Host "存在しています ($oneDrivePath)" -ForegroundColor Green
            
            # 空き容量の確認
            $drive = Get-PSDrive -Name ($oneDrivePath.Substring(0, 1))
            $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
            $usedSpaceGB = [math]::Round(($drive.Used) / 1GB, 2)
            $totalSpaceGB = [math]::Round(($drive.Free + $drive.Used) / 1GB, 2)
            
            Write-Host "ドライブ容量: 合計 ${totalSpaceGB}GB (使用中: ${usedSpaceGB}GB, 空き: ${freeSpaceGB}GB)"
            
            # OneDriveフォルダサイズの取得
            Write-Host "OneDriveフォルダサイズを計算しています..." -ForegroundColor Yellow
            $folderSize = Get-ChildItem -Path $oneDrivePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
            $folderSizeGB = [math]::Round($folderSize.Sum / 1GB, 2)
            Write-Host "OneDriveフォルダサイズ: ${folderSizeGB}GB" -ForegroundColor Cyan
        } else {
            Write-Host "OneDriveフォルダ: " -NoNewline
            Write-Host "見つかりません" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    }
    
    Write-Host "`n任意のキーを押して続行してください..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Export-ModuleMember -Function Get-OneDriveStatus
