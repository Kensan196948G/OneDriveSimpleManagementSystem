# OneDrive接続診断モジュール
# 文字コード: UTF-8 with BOM

function Test-OneDriveConnection {
    Clear-Host
    Write-Host "=====================================================
  OneDrive接続診断
=====================================================" -ForegroundColor Cyan
    
    try {
        # インターネット接続の確認
        Write-Host "インターネット接続を確認しています..." -NoNewline
        if (Test-Connection -ComputerName www.microsoft.com -Count 2 -Quiet) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "失敗" -ForegroundColor Red
            Write-Host "インターネット接続に問題があります。ネットワーク設定を確認してください。" -ForegroundColor Yellow
        }
        
        # OneDriveサービス接続の確認
        Write-Host "OneDriveサービスへの接続を確認しています..." -NoNewline
        if (Test-Connection -ComputerName onedrive.live.com -Count 2 -Quiet) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "失敗" -ForegroundColor Red
            Write-Host "OneDriveサービスに接続できません。" -ForegroundColor Yellow
        }
        
        # OneDriveプロセスの確認
        Write-Host "OneDriveプロセスの状態を確認しています..." -NoNewline
        $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
        if ($oneDriveProcess) {
            Write-Host "実行中" -ForegroundColor Green
        } else {
            Write-Host "停止中" -ForegroundColor Red
            Write-Host "OneDriveが実行されていません。手動で起動してください。" -ForegroundColor Yellow
        }
        
        # OneDrive設定の確認
        Write-Host "`nOneDrive診断情報:" -ForegroundColor Cyan
        Write-Host "----------------------------------------"
        $oneDrivePath = [System.Environment]::GetFolderPath("UserProfile") + "\OneDrive"
        Write-Host "OneDriveフォルダパス: $oneDrivePath"
        
        # 現在のユーザー名を表示
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Write-Host "現在のユーザー: $currentUser"
        
    } catch {
        Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    }
    
    Write-Host "`n任意のキーを押して続行してください..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Export-ModuleMember -Function Test-OneDriveConnection
