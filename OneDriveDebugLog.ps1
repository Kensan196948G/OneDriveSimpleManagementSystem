# OneDriveDebugLog.ps1
# OneDriveの接続テストと詳細ログ出力を行うスクリプト
# 
# 更新履歴:
# 2025/03/10 - 初期バージョン作成
# 2025/03/15 - 古いPowerShellバージョン対応

# PowerShellのエンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ログファイルの設定
$logFolder = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}
$logFile = Join-Path $logFolder "OneDriveDebug_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ログ出力関数
function Write-DebugLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # ログファイルに出力
    $logMessage | Out-File -FilePath $logFile -Encoding UTF8 -Append
    
    # コンソールにも出力
    $color = switch ($Level) {
        "INFO"    { "White" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        "SUCCESS" { "Green" }
        default   { "Gray" }
    }
    
    Write-Host $logMessage -ForegroundColor $color
}

# Microsoft Graph モジュールの確認
function Test-GraphModules {
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Files"
    )
    
    $allModulesPresent = $true
    
    foreach ($module in $requiredModules) {
        if (Get-Module -ListAvailable -Name $module) {
            $installedVersion = (Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-DebugLog "モジュール $module はインストール済み (バージョン: $installedVersion)" -Level "SUCCESS"
            
            # モジュールを読み込む
            Import-Module $module -Force
        } else {
            Write-DebugLog "モジュール $module がインストールされていません" -Level "ERROR"
            $allModulesPresent = $false
        }
    }
    
    return $allModulesPresent
}

# Graph APIの接続テスト
function Test-GraphConnection {
    try {
        Write-DebugLog "Microsoft Graph APIへの接続を試行します..." -Level "INFO"
        
        # 権限スコープ
        $scopes = @(
            "User.Read",
            "Files.Read"
        )
        
        # 接続を試行
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        
        # 接続情報を取得
        $context = Get-MgContext
        if ($context) {
            Write-DebugLog "接続成功: アカウント $($context.Account)" -Level "SUCCESS"
            Write-DebugLog "アプリケーション: $($context.AppName)" -Level "INFO"
            Write-DebugLog "テナントID: $($context.TenantId)" -Level "INFO"
            Write-DebugLog "スコープ: $($context.Scopes -join ', ')" -Level "INFO"
            
            return $true
        } else {
            Write-DebugLog "接続に成功しましたが、コンテキストが取得できません" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-DebugLog "接続エラー: $($_.Exception.Message)" -Level "ERROR"
        if ($_.Exception.InnerException) {
            Write-DebugLog "詳細エラー: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
        return $false
    }
}

# OneDrive情報の取得テスト
function Test-OneDriveAccess {
    try {
        Write-DebugLog "自分のOneDrive情報へのアクセスをテストしています..." -Level "INFO"
        
        # 自分のOneDriveを取得
        $drive = Get-MgDrive -ErrorAction Stop
        
        if ($drive) {
            Write-DebugLog "OneDriveへのアクセス成功 (ID: $($drive.Id))" -Level "SUCCESS"
            Write-DebugLog "ドライブタイプ: $($drive.DriveType)" -Level "INFO"
            Write-DebugLog "所有者: $($drive.Owner.User.DisplayName)" -Level "INFO"
            
            # クォータ情報を表示
            if ($drive.Quota) {
                Write-DebugLog "クォータ状態: $($drive.Quota.State)" -Level "INFO"
                
                if ($drive.Quota.Total -gt 0) {
                    $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
                    $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
                    $remainingGB = [math]::Round($drive.Quota.Remaining / 1GB, 2)
                    $usagePercentage = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 1)
                    
                    Write-DebugLog "総容量: $totalGB GB" -Level "INFO"
                    Write-DebugLog "使用容量: $usedGB GB" -Level "INFO"
                    Write-DebugLog "残り容量: $remainingGB GB" -Level "INFO"
                    Write-DebugLog "使用率: $usagePercentage%" -Level "INFO"
                } else {
                    Write-DebugLog "総容量: 不明/無制限" -Level "WARNING"
                    $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
                    Write-DebugLog "使用容量: $usedGB GB" -Level "INFO"
                }
            } else {
                Write-DebugLog "クォータ情報が取得できません" -Level "WARNING"
            }
            
            # ルートフォルダの項目を取得
            try {
                Write-DebugLog "ルートフォルダの内容を取得しています..." -Level "INFO"
                $rootItems = Get-MgDriveRootChild -DriveId $drive.Id -Top 5
                
                if ($rootItems -and $rootItems.Count -gt 0) {
                    Write-DebugLog "ルートフォルダ内の項目数: $($rootItems.Count)" -Level "SUCCESS"
                    foreach ($item in $rootItems) {
                        # 三項演算子を使わない方法に修正 (下位互換性対応)
                        $itemType = if ($null -ne $item.File) { "ファイル" } else { "フォルダ" }
                        Write-DebugLog "  - $($item.Name) (タイプ: $itemType, 更新: $($item.LastModifiedDateTime))" -Level "INFO"
                    }
                } else {
                    Write-DebugLog "ルートフォルダは空か、アクセスできません" -Level "WARNING"
                }
            }
            catch {
                Write-DebugLog "ルートフォルダの内容取得中にエラー: $($_.Exception.Message)" -Level "ERROR"
            }
            
            return $true
        } else {
            Write-DebugLog "OneDriveが見つかりませんでした" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-DebugLog "OneDriveアクセス中にエラーが発生: $($_.Exception.Message)" -Level "ERROR"
        if ($_.Exception.InnerException) {
            Write-DebugLog "詳細エラー: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
        return $false
    }
}

# アクセス権限の確認
function Test-AdminAccess {
    try {
        Write-DebugLog "管理者権限の確認を行っています..." -Level "INFO"
        
        # 現在のユーザー情報を取得
        $context = Get-MgContext
        if (-not $context) {
            Write-DebugLog "認証コンテキストがありません" -Level "ERROR"
            return $false
        }
        
        # ユーザープリンシパル名でユーザーを検索
        $filter = "userPrincipalName eq '$($context.Account)'"
        $user = Get-MgUser -Filter $filter
        
        if (-not $user) {
            Write-DebugLog "現在のユーザー情報が取得できません" -Level "ERROR"
            return $false
        }
        
        Write-DebugLog "現在のユーザー: $($user.DisplayName) ($($user.UserPrincipalName))" -Level "INFO"
        
        # ユーザーのロールを確認
        $roles = Get-MgUserMemberOf -UserId $user.Id
        
        $adminRoles = @(
            'Global Administrator',
            'Company Administrator',
            'SharePoint Administrator',
            'Global Reader',
            'Teams Administrator',
            'OneDrive管理者',
            'SharePoint管理者',
            'グローバル管理者'
        )
        
        $isAdmin = $false
        foreach ($role in $roles) {
            $roleType = $role.AdditionalProperties.'@odata.type'
            $displayName = $role.AdditionalProperties.displayName
            
            Write-DebugLog "ロール/グループ: $displayName (タイプ: $roleType)" -Level "INFO"
            
            if (($roleType -eq "#microsoft.graph.directoryRole" -or $roleType -eq "#microsoft.graph.group") -and ($adminRoles -contains $displayName)) {
                Write-DebugLog "管理者ロールを検出: $displayName" -Level "SUCCESS"
                $isAdmin = $true
            }
        }
        
        if ($isAdmin) {
            Write-DebugLog "ユーザーは管理者権限を持っています" -Level "SUCCESS"
        } else {
            Write-DebugLog "ユーザーは管理者権限を持っていません" -Level "WARNING"
        }
        
        return $isAdmin
    }
    catch {
        Write-DebugLog "管理者権限確認中にエラー: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# メイン処理
Write-DebugLog "OneDrive診断ツールを開始します" -Level "INFO"
Write-DebugLog "ログファイル: $logFile" -Level "INFO"

# 実行中のPowerShellバージョンを確認
$psVersion = $PSVersionTable.PSVersion
Write-DebugLog "PowerShellバージョン: $($psVersion.Major).$($psVersion.Minor).$($psVersion.Build)" -Level "INFO"

# モジュールの確認
$modulesOk = Test-GraphModules
if (-not $modulesOk) {
    Write-DebugLog "必要なモジュールがインストールされていません。インストールしますか？" -Level "WARNING"
    $installModules = Read-Host "モジュールをインストールしますか？ (Y/N)"
    
    if ($installModules -eq 'Y' -or $installModules -eq 'y') {
        try {
            foreach ($module in @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Files")) {
                Write-DebugLog "モジュール $module をインストールしています..." -Level "INFO"
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Import-Module $module -Force
                Write-DebugLog "モジュール $module をインストールしました" -Level "SUCCESS"
            }
            $modulesOk = $true
        }
        catch {
            Write-DebugLog "モジュールのインストール中にエラー: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

if ($modulesOk) {
    try {
        # GraphAPIへの接続
        $connected = Test-GraphConnection
        
        if ($connected) {
            # OneDriveへのアクセスをテスト
            $oneDriveAccess = Test-OneDriveAccess
            
            if ($oneDriveAccess) {
                Write-DebugLog "OneDriveへのアクセステストは成功しました" -Level "SUCCESS"
                
                # 管理者権限の確認
                $isAdmin = Test-AdminAccess
                if ($isAdmin) {
                    Write-DebugLog "管理者として他のユーザーのOneDrive情報にアクセスできます" -Level "SUCCESS"
                } else {
                    Write-DebugLog "一般ユーザー権限のため、自分のOneDrive情報のみアクセスできます" -Level "INFO"
                }
            } else {
                Write-DebugLog "OneDriveへのアクセステストが失敗しました" -Level "ERROR"
            }
        } else {
            Write-DebugLog "GraphAPIへの接続が確立できませんでした" -Level "ERROR"
        }
    }
    catch {
        Write-DebugLog "処理中にエラーが発生しました: $($_.Exception.Message)" -Level "ERROR"
        if ($_.Exception.InnerException) {
            Write-DebugLog "詳細エラー: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
    }
    finally {
        # 接続を切断
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-DebugLog "GraphAPIから切断しました" -Level "INFO"
        }
        catch {
            Write-DebugLog "切断処理中にエラーが発生: $($_.Exception.Message)" -Level "WARNING"
        }
    }
}

Write-DebugLog "診断が完了しました。ログファイルは $logFile に保存されています" -Level "INFO"
Write-Host "エンターキーを押すと終了します..." -ForegroundColor Yellow
Read-Host | Out-Null
