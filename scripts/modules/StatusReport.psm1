# OneDriveステータスレポートモジュール
# 文字コード: UTF-8 with BOM

function Get-OneDriveStatus {
    Clear-Host
    Write-Host "==========================================
  OneDriveステータスレポート
==========================================" -ForegroundColor Cyan
    
    Write-Host "OneDriveの状態を確認しています..." -ForegroundColor Yellow
    
    # OneDriveプロセスの確認
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        Write-Host "OneDriveプロセス: 実行中" -ForegroundColor Green
    } else {
        Write-Host "OneDriveプロセス: 停止中" -ForegroundColor Red
    }
    
    # OneDriveフォルダの確認
    $oneDrivePath = [System.Environment]::GetEnvironmentVariable("OneDrive", "User")
    if (Test-Path $oneDrivePath) {
        Write-Host "OneDriveフォルダ: 存在します ($oneDrivePath)" -ForegroundColor Green
        
        # 容量情報
        try {
            # ドライブ文字を正しく抽出する方法に修正
            $driveLetter = $oneDrivePath.Substring(0, 1) + ":"
            $drive = Get-PSDrive -Name $driveLetter.Substring(0, 1)
            
            if ($drive) {
                $totalGB = [math]::Round($drive.Used / 1GB + $drive.Free / 1GB, 2)
                $usedGB = [math]::Round($drive.Used / 1GB, 2)
                $freeGB = [math]::Round($drive.Free / 1GB, 2)
                $usedPercent = [math]::Round(($drive.Used / ($drive.Used + $drive.Free)) * 100, 1)
                
                Write-Host "ドライブ容量: $usedGB GB 使用中 / $totalGB GB 合計 ($usedPercent%)" -ForegroundColor Cyan
                Write-Host "空き容量: $freeGB GB" -ForegroundColor Cyan
            } else {
                throw "ドライブ情報を取得できませんでした: $driveLetter"
            }
        } catch {
            Write-Host "容量情報の取得に失敗しました。エラー: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "OneDriveフォルダ: 見つかりません" -ForegroundColor Red
    }
    
    Write-Host "`nOneDrive同期状態の詳細を確認するには、タスクバーのOneDriveアイコンを右クリックしてください。" -ForegroundColor Yellow
    Write-Host "何かキーを押すとメインメニューに戻ります..."
    [void][System.Console]::ReadKey($true)
}

Export-ModuleMember -Function Get-OneDriveStatus
