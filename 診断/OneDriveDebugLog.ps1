#
# OneDrive接続診断ツール
# Microsoft Graph APIとの接続とOneDriveアクセス権限を診断します
#

# 文字エンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function Write-ColorOutput {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    
    $originalColor = $host.UI.RawUI.ForegroundColor
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    Write-Output $Message
    $host.UI.RawUI.ForegroundColor = $originalColor
}

function Write-Log {
    param(
        [string]$Message,
        [string]$LogFile,
        [string]$Type = "INFO",
        [string]$ForegroundColor = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    
    # ログファイルへの書き込み
    if ($LogFile) {
        $logEntry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
    
    # コンソール出力
    Write-ColorOutput -Message $logEntry -ForegroundColor $ForegroundColor
}

# 診断開始
Clear-Host
Write-ColorOutput "===========================================" -ForegroundColor Cyan
Write-ColorOutput "     OneDrive接続診断ツール" -ForegroundColor Cyan
Write-ColorOutput "===========================================" -ForegroundColor Cyan
Write-ColorOutput ""

# ログファイルの設定
$logFolderPath = Join-Path $PSScriptRoot "..\logs"
if (-not (Test-Path -Path $logFolderPath)) {
    New-Item -Path $logFolderPath -ItemType Directory | Out-Null
}

$logFile = Join-Path $logFolderPath ("OneDriveDebug_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
Write-ColorOutput "ログファイル: $logFile" -ForegroundColor Yellow
Write-ColorOutput ""

# システム情報の収集
Write-Log "システム情報の収集を開始します" -LogFile $logFile -Type "INFO" -ForegroundColor Green

try {
    $os = Get-CimInstance Win32_OperatingSystem
    $computerSystem = Get-CimInstance Win32_ComputerSystem
    
    Write-Log "OS: $($os.Caption) $($os.Version)" -LogFile $logFile -Type "INFO"
    Write-Log "コンピュータ名: $($computerSystem.Name)" -LogFile $logFile -Type "INFO"
    Write-Log "PowerShellバージョン: $($PSVersionTable.PSVersion)" -LogFile $logFile -Type "INFO"
}
catch {
    Write-Log "システム情報の収集中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
}

Write-ColorOutput ""

# Microsoft.Graphモジュールの確認
Write-Log "Microsoft.Graphモジュールの確認" -LogFile $logFile -Type "INFO" -ForegroundColor Green

$requiredModules = @(
    "Microsoft.Graph.Authentication", 
    "Microsoft.Graph.Users", 
    "Microsoft.Graph.Files"
)

foreach ($module in $requiredModules) {
    try {
        if (Get-Module -ListAvailable -Name $module) {
            $moduleInfo = Get-Module -ListAvailable -Name $module | Select-Object -First 1
            Write-Log "モジュール $module はインストール済み (v$($moduleInfo.Version))" -LogFile $logFile -Type "INFO" -ForegroundColor Green
            
            # モジュールのインポート
            Import-Module $module -ErrorAction Stop
            Write-Log "モジュール $module をインポートしました" -LogFile $logFile -Type "INFO"
        } else {
            Write-Log "モジュール $module はインストールされていません" -LogFile $logFile -Type "WARNING" -ForegroundColor Yellow
            
            $installChoice = Read-Host "モジュール $module をインストールしますか？ (y/n)"
            if ($installChoice -eq "y") {
                Write-Log "モジュール $module のインストールを開始します" -LogFile $logFile -Type "INFO"
                Install-Module $module -Scope CurrentUser -Force -ErrorAction Stop
                Import-Module $module -ErrorAction Stop
                Write-Log "モジュール $module のインストールとインポートに成功しました" -LogFile $logFile -Type "INFO" -ForegroundColor Green
            } else {
                Write-Log "モジュール $module はインストールされていないため、一部の診断が実行できません" -LogFile $logFile -Type "WARNING" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Log "モジュール $module の確認中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
    }
}

Write-ColorOutput ""

# Microsoft Graph APIへの接続確認
Write-Log "Microsoft Graph APIへの接続確認" -LogFile $logFile -Type "INFO" -ForegroundColor Green

try {
    # 既存の接続確認
    $existingConnection = $null
    try {
        $existingConnection = Get-MgContext
    } catch {
        Write-Log "既存のGraph接続はありません（正常）" -LogFile $logFile -Type "INFO"
    }
    
    if ($existingConnection) {
        Write-Log "既存のGraph接続があります: $($existingConnection.Account)" -LogFile $logFile -Type "INFO" -ForegroundColor Green
        $reconnect = Read-Host "再接続しますか？ (y/n)"
        
        if ($reconnect -eq "y") {
            Disconnect-MgGraph | Out-Null
            Write-Log "既存の接続を切断しました" -LogFile $logFile -Type "INFO"
            $existingConnection = $null
        }
    }
    
    if (-not $existingConnection) {
        Write-Log "Microsoft Graph APIへの接続を開始します" -LogFile $logFile -Type "INFO"
        
        # 接続パラメータの設定
        $params = @{
            Scopes = @(
                "User.Read",
                "Files.Read",
                "Files.Read.All",
                "User.Read.All"
            )
        }
        
        Connect-MgGraph @params -ErrorAction Stop
        $connection = Get-MgContext
        
        if ($connection) {
            Write-Log "Microsoft Graph APIへの接続に成功しました" -LogFile $logFile -Type "INFO" -ForegroundColor Green
            Write-Log "アカウント: $($connection.Account)" -LogFile $logFile -Type "INFO"
            Write-Log "テナント: $($connection.TenantId)" -LogFile $logFile -Type "INFO"
            Write-Log "スコープ: $($connection.Scopes -join ', ')" -LogFile $logFile -Type "INFO"
        } else {
            Write-Log "Microsoft Graph APIへの接続に失敗しました" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
        }
    }
    
    # 現在のユーザー情報を取得
    try {
        Write-Log "現在のユーザー情報を取得します" -LogFile $logFile -Type "INFO"
        
        # プロファイルを取得
        $currentContext = Get-MgContext
        if ($currentContext) {
            $filter = "userPrincipalName eq '$($currentContext.Account)'"
            $currentUser = Get-MgUser -Filter $filter -ErrorAction Stop
            
            if ($currentUser) {
                Write-Log "ユーザー情報: $($currentUser.DisplayName) ($($currentUser.UserPrincipalName))" -LogFile $logFile -Type "INFO" -ForegroundColor Green
                
                # 管理者かどうかを確認
                try {
                    $roles = Get-MgUserMemberOf -UserId $currentUser.Id -All
                    $adminRoles = @()
                    
                    foreach ($role in $roles) {
                        $roleType = $role.AdditionalProperties.'@odata.type'
                        if ($roleType -eq "#microsoft.graph.directoryRole") {
                            $roleName = $role.AdditionalProperties.displayName
                            $adminRoles += $roleName
                        }
                    }
                    
                    if ($adminRoles.Count -gt 0) {
                        Write-Log "管理者ロール: $($adminRoles -join ', ')" -LogFile $logFile -Type "INFO" -ForegroundColor Green
                    } else {
                        Write-Log "このアカウントに管理者ロールはありません" -LogFile $logFile -Type "WARNING" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Log "ロール情報の取得中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
                }
                
                # OneDriveの取得を試行
                try {
                    Write-Log "OneDrive情報の取得を試みます" -LogFile $logFile -Type "INFO"
                    $drive = Get-MgUserDrive -UserId $currentUser.Id
                    
                    if ($drive) {
                        Write-Log "OneDrive情報の取得に成功しました" -LogFile $logFile -Type "INFO" -ForegroundColor Green
                        Write-Log "OneDriveタイプ: $($drive.DriveType)" -LogFile $logFile -Type "INFO"
                        
                        # クォータ情報があれば表示
                        if ($drive.Quota) {
                            $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
                            $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
                            $remainingGB = [math]::Round($drive.Quota.Remaining / 1GB, 2)
                            
                            Write-Log "クォータ状態: $($drive.Quota.State)" -LogFile $logFile -Type "INFO"
                            Write-Log "総容量: $totalGB GB" -LogFile $logFile -Type "INFO"
                            Write-Log "使用中: $usedGB GB" -LogFile $logFile -Type "INFO"
                            Write-Log "残り: $remainingGB GB" -LogFile $logFile -Type "INFO"
                        } else {
                            Write-Log "クォータ情報がありません" -LogFile $logFile -Type "WARNING" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log "OneDrive情報を取得できませんでした" -LogFile $logFile -Type "WARNING" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Log "OneDrive情報の取得中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
                }
            } else {
                Write-Log "現在のユーザー情報を取得できませんでした" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
            }
        } else {
            Write-Log "Microsoft Graph APIに接続されていません" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
        }
    }
    catch {
        Write-Log "ユーザー情報の取得中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
    }
}
catch {
    Write-Log "Microsoft Graph APIへの接続確認中にエラーが発生しました: $_" -LogFile $logFile -Type "ERROR" -ForegroundColor Red
}

Write-ColorOutput ""
Write-Log "診断が完了しました" -LogFile $logFile -Type "INFO