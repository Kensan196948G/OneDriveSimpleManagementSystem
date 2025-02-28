# OneDrive診断モジュール
# 文字コード: UTF-8 with BOM

function Test-OneDriveConnection {
    Clear-Host
    Write-Host "==========================================
  OneDrive接続診断
==========================================" -ForegroundColor Cyan
    
    Write-Host "OneDriveの接続状態を診断しています..." -ForegroundColor Yellow
    
    # インターネット接続の確認
    Write-Host "`n[1/4] インターネット接続を確認しています..." -ForegroundColor Yellow
    $internetTest = Test-Connection -ComputerName "www.microsoft.com" -Count 2 -Quiet
    if ($internetTest) {
        Write-Host "インターネット接続: 正常" -ForegroundColor Green
    } else {
        Write-Host "インターネット接続: 問題あり - ネットワーク設定を確認してください" -ForegroundColor Red
    }
    
    # OneDriveプロセスの確認
    Write-Host "`n[2/4] OneDriveプロセスを確認しています..." -ForegroundColor Yellow
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        Write-Host "OneDriveプロセス: 実行中" -ForegroundColor Green
    } else {
        Write-Host "OneDriveプロセス: 停止中 - OneDriveを起動してください" -ForegroundColor Red
    }
    
    # OneDriveサービスエンドポイントの確認
    Write-Host "`n[3/4] OneDriveサービスへの接続を確認しています..." -ForegroundColor Yellow
    $endpoints = @(
        "onedrive.live.com",
        "login.live.com",
        "oneclient.sfx.ms"
    )
    
    $allEndpointsAccessible = $true
    foreach ($endpoint in $endpoints) {
        try {
            $result = Test-NetConnection -ComputerName $endpoint -Port 443 -WarningAction SilentlyContinue
            if ($result.TcpTestSucceeded) {
                Write-Host "$endpoint: 接続成功" -ForegroundColor Green
            } else {
                Write-Host "$endpoint: 接続失敗" -ForegroundColor Red
                $allEndpointsAccessible = $false
            }
        } catch {
            Write-Host "$endpoint: 接続テスト中にエラーが発生しました" -ForegroundColor Red
            $allEndpointsAccessible = $false
        }
    }
    
    # OneDrive認証の確認
    Write-Host "`n[4/4] OneDrive認証状態を確認しています..." -ForegroundColor Yellow
    $oneDrivePath = [System.Environment]::GetEnvironmentVariable("OneDrive", "User")
    if (Test-Path $oneDrivePath) {
        Write-Host "OneDriveフォルダ: アクセス可能" -ForegroundColor Green
    } else {
        Write-Host "OneDriveフォルダ: アクセスできません - サインインが必要な可能性があります" -ForegroundColor Red
    }
    
    # 診断結果のサマリー
    Write-Host "`n診断結果サマリー:" -ForegroundColor Cyan
    if ($internetTest -and $oneDriveProcess -and $allEndpointsAccessible -and (Test-Path $oneDrivePath)) {
        Write-Host "OneDriveは正常に動作しているようです。" -ForegroundColor Green
    } else {
        Write-Host "OneDriveに問題が見つかりました。以下の対処を試してください:" -ForegroundColor Yellow
        Write-Host "1. OneDriveアプリケーションを再起動する" -ForegroundColor Yellow
        Write-Host "2. Microsoftアカウントに再サインインする" -ForegroundColor Yellow
        Write-Host "3. インターネット接続を確認する" -ForegroundColor Yellow
        Write-Host "4. OneDriveアプリケーションを再インストールする" -ForegroundColor Yellow
    }
    
    Write-Host "`n何かキーを押すとメインメニューに戻ります..."
    [void][System.Console]::ReadKey($true)
}

Export-ModuleMember -Function Test-OneDriveConnection
