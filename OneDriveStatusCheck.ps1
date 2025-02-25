# OneDriveStatusCheck.ps1
# OneDriveのステータスを確認し、HTML/CSV/JSファイルとして出力する
# 
# 更新履歴:
# 2025/03/05 - 認証エラー修正、"me"エンドポイント問題解決、出力処理実装

# PowerShell セッション内の全関数を削除（再定義防止）
Get-ChildItem Function:\ | Remove-Item -Force -ErrorAction SilentlyContinue

# ログ出力関数を最初に定義
function Write-DetailLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # ログファイルが定義されている場合のみファイル出力
    if ($script:logFilePath) {
        $logMessage | Out-File -FilePath $script:logFilePath -Encoding UTF8 -Append
    }

    $consoleColor = switch ($Level) {
        'INFO'    { 'White' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
    }
    Write-Host $logMessage -ForegroundColor $consoleColor
}

# エラーハンドリング関数の定義
function Write-ErrorLog {
    param(
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$CustomMessage
    )
    
    Write-DetailLog "エラーが発生しました: $CustomMessage" -Level ERROR
    Write-DetailLog "エラーの種類: $($ErrorRecord.Exception.GetType().Name)" -Level ERROR
    Write-DetailLog "エラーメッセージ: $($ErrorRecord.Exception.Message)" -Level ERROR
    Write-DetailLog "発生場所: $($ErrorRecord.InvocationInfo.PositionMessage)" -Level ERROR
    
    if ($ErrorRecord.Exception.InnerException) {
        Write-DetailLog "内部エラー: $($ErrorRecord.Exception.InnerException.Message)" -Level ERROR
    }
    
    if ($ErrorRecord.ScriptStackTrace) {
        Write-DetailLog "スタックトレース:`n$($ErrorRecord.ScriptStackTrace)" -Level ERROR
    }
}

function Set-EncodingEnvironment {
    # 現在のエンコーディング設定を保存
    $script:originalOutputEncoding = [Console]::OutputEncoding
    $script:originalInputEncoding = [Console]::InputEncoding
    
    # 現在のエンコーディング情報を表示
    $currentEncoding = [Console]::OutputEncoding
    Write-Host "現在の出力エンコーディング: $($currentEncoding.EncodingName) (CodePage: $($currentEncoding.CodePage))" -ForegroundColor Yellow
    
    # PowerShellのエンコーディングをUTF-8に設定
    if ($currentEncoding.CodePage -ne 65001) {
        Write-Host "出力エンコーディングをUTF-8に変更しています..." -ForegroundColor Yellow
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'
        
        $newEncoding = [Console]::OutputEncoding
        Write-Host "変更後の出力エンコーディング: $($newEncoding.EncodingName) (CodePage: $($newEncoding.CodePage))" -ForegroundColor Green
    } else {
        Write-Host "出力エンコーディングはすでにUTF-8です。変更は不要です。" -ForegroundColor Green
    }
    
    # 環境変数も設定
    $env:PYTHONIOENCODING = "utf-8"
    [System.Environment]::SetEnvironmentVariable('PYTHONIOENCODING', 'utf-8', 'Process')
}

function Restore-OriginalEncoding {
    if ($script:originalOutputEncoding) {
        Write-Host "元のエンコーディングに戻しています..." -ForegroundColor Yellow
        [Console]::OutputEncoding = $script:originalOutputEncoding
        [Console]::InputEncoding = $script:originalInputEncoding
        $OutputEncoding = $script:originalOutputEncoding
        
        $restoredEncoding = [Console]::OutputEncoding
        Write-Host "復元後の出力エンコーディング: $($restoredEncoding.EncodingName) (CodePage: $($restoredEncoding.CodePage))" -ForegroundColor Green
    }
}

# Microsoft Azure ADポータルでアプリケーションの登録を確認・修正する手順を表示する関数
function Show-AppRegistrationInstructions {
    Write-Host "`n=================================================" -ForegroundColor Yellow
    Write-Host "  Microsoft Graph API 権限設定の確認方法" -ForegroundColor Yellow
    Write-Host "=================================================" -ForegroundColor Yellow
    Write-Host "1. Azure ADポータル (https://portal.azure.com) にグローバル管理者でサインイン"
    Write-Host "2. 「Azure Active Directory」→「アプリの登録」を選択"
    Write-Host "3. 「Microsoft Graph Command Line Tools」または類似名のアプリを探して選択"
    Write-Host "4. 「API のアクセス許可」を選択"
    Write-Host "5. 以下の権限が付与されているか確認:"
    Write-Host "   - User.Read.All"
    Write-Host "   - Files.Read.All"
    Write-Host "6. 「(テナント名) に管理者の同意を与える」ボタンをクリック"
    Write-Host "7. 不足している権限がある場合は「アクセス許可の追加」で追加"
    Write-Host "8. 追加後、再度「管理者の同意を与える」をクリック"
    Write-Host "=================================================" -ForegroundColor Yellow
    
    $checkPermissions = Read-Host "アプリケーション権限を確認しますか？ (y/n)"
    if ($checkPermissions -eq "y") {
        Start-Process "https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"
    }
}

# アクセストークンの管理のためのグローバル変数
$script:AccessToken = $null
$script:TokenExpiration = [DateTime]::MinValue
$script:IsAdmin = $false
$script:CurrentUserId = $null

# 管理者権限の確認関数（/me エンドポイントを回避）
function Test-IsGlobalAdmin {
    try {
        # ユーザー自身のIDを取得
        if (-not $script:CurrentUserId) {
            $context = Get-MgContext
            if (-not $context) {
                Write-DetailLog "認証コンテキストを取得できません。" -Level ERROR
                return $false
            }
            
            # ユーザープリンシパル名でユーザーを検索
            try {
                $filter = "userPrincipalName eq '$($context.Account)'"
                $user = Get-MgUser -Filter $filter -ErrorAction Stop
                if ($user) {
                    $script:CurrentUserId = $user.Id
                    Write-DetailLog "現在のユーザーID: $($script:CurrentUserId)" -Level INFO
                }
                else {
                    Write-DetailLog "現在のユーザーを取得できませんでした。" -Level WARNING
                    return $false
                }
            }
            catch {
                Write-DetailLog "現在のユーザーID取得中にエラー: $($_.Exception.Message)" -Level WARNING
                return $false
            }
        }
        
        # ロールの確認
        try {
            $roles = Get-MgUserMemberOf -UserId $script:CurrentUserId -All -ErrorAction Stop
            
            # ロールとグループのチェック
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
            
            foreach ($role in $roles) {
                $roleType = $role.AdditionalProperties.'@odata.type'
                $displayName = $role.AdditionalProperties.displayName
                
                if (($roleType -eq "#microsoft.graph.directoryRole" -or 
                     $roleType -eq "#microsoft.graph.group") -and 
                    ($adminRoles -contains $displayName)) {
                    Write-DetailLog "管理者ロールを発見: $displayName" -Level INFO
                    return $true
                }
            }
            
            Write-DetailLog "ユーザーは管理者ロールを持っていません。" -Level INFO
            return $false
        }
        catch {
            Write-DetailLog "ロール確認中にエラー: $($_.Exception.Message)" -Level WARNING
            return $false
        }
    }
    catch {
        Write-DetailLog "管理者権限の確認中にエラーが発生しました: $($_.Exception.Message)" -Level WARNING
        return $false
    }
}

# アクセストークンを取得する関数
function Get-ValidAccessToken {
    try {
        # トークンが有効期限切れか確認
        $now = [DateTime]::UtcNow
        if (($script:AccessToken -eq $null) -or ($now -ge $script:TokenExpiration)) {
            Write-DetailLog "新しいアクセストークンを取得しています..." -Level INFO
            
            # 現在のコンテキストを確認
            $context = Get-MgContext
            if (-not $context) {
                Write-DetailLog "Microsoft Graph APIに接続されていません。再接続を試みます。" -Level WARNING
                
                # 基本的なスコープで再接続を試みる
                $minimalScopes = @("User.Read", "Files.Read")
                Connect-MgGraphSafely -RequiredScopes $minimalScopes
                $context = Get-MgContext
                
                if (-not $context) {
                    Write-DetailLog "Microsoft Graph APIへの再接続に失敗しました。" -Level ERROR
                    return $null
                }
            }
            
            # アクセストークンを取得
            try {
                # トークンの文字列を取得するためのリフレクション
                $contextClassType = [Microsoft.Graph.PowerShell.Authentication.GraphSession].Assembly.GetType('Microsoft.Graph.PowerShell.Authentication.GraphSession')
                $tokenProperty = $contextClassType.GetProperty('AuthContext', [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance)
                $authContext = $tokenProperty.GetValue($context)
                
                # アクセストークンとその有効期限を取得
                $tokenField = $authContext.GetType().GetProperty('AccessToken')
                $script:AccessToken = $tokenField.GetValue($authContext)
                
                # 有効期限の取得
                $expiryField = $authContext.GetType().GetProperty('AccessTokenExpiration')
                if ($expiryField) {
                    $script:TokenExpiration = $expiryField.GetValue($authContext)
                    $timeUntilExpiry = ($script:TokenExpiration - $now).TotalMinutes
                    Write-DetailLog "アクセストークンを取得しました。有効期限まで約 $([Math]::Round($timeUntilExpiry)) 分です。" -Level INFO
                } else {
                    # 有効期限が取得できない場合は、60分と仮定
                    $script:TokenExpiration = $now.AddMinutes(60)
                    Write-DetailLog "アクセストークンを取得しました。有効期限は不明ですが、60分間有効と仮定します。" -Level INFO
                }
            }
            catch {
                Write-DetailLog "トークン取得中にエラーが発生しました: $($_.Exception.Message)" -Level ERROR
                
                # 詳細なエラー診断
                if ($_.Exception.Message -match "権限") {
                    Write-DetailLog "アクセス権限が不足している可能性があります。管理者権限が必要かもしれません。" -Level WARNING
                    Show-AppRegistrationInstructions
                }
                
                # 最後の手段としてコンテキストから直接アクセストークン取得を試みる
                try {
                    $script:AccessToken = $context.AccessToken
                    Write-DetailLog "コンテキストから直接アクセストークンを取得しました。有効期限は不明です。" -Level WARNING
                }
                catch {
                    Write-DetailLog "コンテキストからのトークン取得も失敗しました。" -Level ERROR
                    return $null
                }
            }
        }
        
        # 一般ユーザー向けのメッセージ
        if (-not $script:IsAdmin -and $script:AccessToken) {
            Write-DetailLog "一般ユーザー権限でアクセストークンを使用します。一部の情報は取得できない場合があります。" -Level INFO
        }
        
        return $script:AccessToken
    }
    catch {
        Write-DetailLog "アクセストークンの取得に失敗しました: $($_.Exception.Message)" -Level ERROR
        Write-ErrorLog $_ "アクセストークン処理中のエラー"
        return $null
    }
}

# Microsoft Graph モジュールの確認とインストール関数
function Install-RequiredModules {
    $modules = @("Microsoft.Graph.Users", "Microsoft.Graph.Files", "Microsoft.Graph.Authentication")
    
    # セッションクリーンアップ
    Get-Module | Where-Object { $_.Name -like "Microsoft.Graph*" } | Remove-Module -Force
    Get-Item Function:\Get-Mg* -ErrorAction SilentlyContinue | Remove-Item -Force
    
    foreach ($module in $modules) {
        Write-Host "モジュール $module の設定を開始します..." -ForegroundColor Yellow
        
        try {
            # モジュールが既にインポートされているか確認
            if (-not (Get-Module -Name $module)) {
                Write-DetailLog "モジュール $module をインストールします..." -Level INFO
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-DetailLog "モジュール $module をインポートします..." -Level INFO
                Import-Module -Name $module -Force -DisableNameChecking -ErrorAction Stop
                Write-DetailLog "モジュール $module を読み込みました。" -Level INFO
            } else {
                Write-DetailLog "モジュール $module は既に読み込まれています。" -Level INFO
            }
        }
        catch {
            Write-DetailLog "モジュール $module の読み込みに失敗: $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }
    return $true
}

# 既存のパーミッションチェックと安全な接続処理（改善版）
function Connect-MgGraphSafely {
    param (
        [array]$RequiredScopes,
        [int]$RetryCount = 2
    )

    try {
        Write-DetailLog "認証プロセスを開始します..." -Level INFO
        
        # 最小限の権限（自分のユーザー情報とファイル）
        $minimalScopes = @(
            "User.Read",          # 自分のユーザー情報の読み取り
            "Files.Read"          # 自分のファイルの読み取り
        )
        
        # 既存の接続を確認
        $existingContext = Get-MgContext
        if ($existingContext) {
            Write-DetailLog "既存の接続を検出: $($existingContext.Account)" -Level INFO
            
            # トークンの取得と確認
            $token = Get-ValidAccessToken
            if ($token) {
                Write-DetailLog "有効なアクセストークンがあります。" -Level INFO
                
                # ユーザー自身のIDを取得（後で使用）
                try {
                    $filter = "userPrincipalName eq '$($existingContext.Account)'"
                    $currentUser = Get-MgUser -Filter $filter -ErrorAction Stop
                    if ($currentUser) {
                        $script:CurrentUserId = $currentUser.Id
                        Write-DetailLog "現在のユーザーID: $($script:CurrentUserId)" -Level INFO
                    }
                }
                catch {
                    Write-DetailLog "現在のユーザーID取得中にエラー: $($_.Exception.Message)" -Level WARNING
                }
                
                # 権限チェック（/me エンドポイントを使わない）
                $script:IsAdmin = Test-IsGlobalAdmin
                if ($script:IsAdmin) {
                    Write-DetailLog "管理者権限を持つユーザーとして接続しています。" -Level INFO
                } else {
                    Write-DetailLog "一般ユーザー権限で接続しています。一部の情報は取得できません。" -Level WARNING
                }
                
                return $existingContext
            }
            
            Write-DetailLog "アクセストークンの更新が必要です。" -Level INFO
        }

        # 認証方法の選択
        Write-Host "`n認証方法を選択してください：" -ForegroundColor Yellow
        Write-Host "1: 通常の認証（推奨）- Microsoftアカウントとパスワードを入力" -ForegroundColor Green
        Write-Host "2: デバイスコード認証 - コードを使用して別デバイスで認証" -ForegroundColor Green
        $authChoice = Read-Host "`n選択 (1 or 2)"

        # 認証パラメータの設定
        $connectParams = @{
            Scopes = $RequiredScopes
            ErrorAction = "Stop"
        }

        if ($authChoice -eq "2") {
            $connectParams["UseDeviceAuthentication"] = $true
        }

        # 明示的に同意を求める
        if ($RetryCount -eq 2) {  # 最初の試行の場合のみ
            Write-DetailLog "アプリケーション権限への同意を要求します..." -Level INFO
        }

        try {
            # スコープを表示
            Write-DetailLog "$($connectParams["Scopes"] -join ", ") スコープでの認証を試行します..." -Level INFO
            
            # 通常の認証を試行
            Connect-MgGraph @connectParams
            
            $context = Get-MgContext
            if (-not $context) {
                throw "認証コンテキストを取得できません。"
            }

            Write-DetailLog "認証成功: $($context.Account)" -Level INFO
            
            # ユーザー自身のIDを取得（後で使用）
            try {
                $filter = "userPrincipalName eq '$($context.Account)'"
                $currentUser = Get-MgUser -Filter $filter -ErrorAction Stop
                if ($currentUser) {
                    $script:CurrentUserId = $currentUser.Id
                    Write-DetailLog "現在のユーザーID: $($script:CurrentUserId)" -Level INFO
                }
            }
            catch {
                Write-DetailLog "現在のユーザーID取得中にエラー: $($_.Exception.Message)" -Level WARNING
            }
            
            # 権限チェック（/me エンドポイントを使わない）
            $script:IsAdmin = Test-IsGlobalAdmin
            
            # アクセストークンを保存
            $script:AccessToken = $context.AccessToken
            if ($context.ExpiresOn) {
                $script:TokenExpiration = $context.ExpiresOn.AddMinutes(-30)
            }
            else {
                $script:TokenExpiration = [DateTime]::UtcNow.AddHours(1)
            }
            
            return $context
        }
        catch {
            if ($RetryCount -gt 0) {
                Write-DetailLog "認証に失敗しました: $($_.Exception.Message)" -Level WARNING
                Write-DetailLog "最小権限で再試行します..." -Level INFO
                
                # 最小権限で再試行
                Start-Sleep -Seconds 2
                return Connect-MgGraphSafely -RequiredScopes $minimalScopes -RetryCount ($RetryCount - 1)
            }
            else {
                throw "複数回の認証試行が失敗しました: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-DetailLog "認証プロセスでエラーが発生しました: $($_.Exception.Message)" -Level ERROR
        
        # エラー詳細を記録
        if ($_.Exception.InnerException) {
            Write-DetailLog "詳細エラー: $($_.Exception.InnerException.Message)" -Level ERROR
        }
        
        throw "Microsoft Graph API への接続に失敗しました。"
    }
}

# 安全にプロパティ値を取得する関数
function Get-SafePropertyValue {
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$InputObject,
        
        [Parameter(Mandatory=$true)]
        [string]$PropertyName,
        
        [Parameter(Mandatory=$false)]
        $DefaultValue = 0
    )
    
    try {
        if ($null -eq $InputObject) { return $DefaultValue }
        if (-not (Get-Member -InputObject $InputObject -Name $PropertyName -ErrorAction SilentlyContinue)) { return $DefaultValue }
        $value = $InputObject.$PropertyName
        if ($null -eq $value) { return $DefaultValue }
        return $value
    }
    catch {
        Write-DetailLog "プロパティ '$PropertyName' の取得に失敗しました: $($_.Exception.Message)" -Level WARNING
        return $DefaultValue
    }
}

# OneDriveの状態を取得する関数
function Get-OneDriveStatus {
    param (
        [string]$userId,
        [switch]$SelfOnly
    )
    
    try {
        # 自分自身のユーザーIDを使用する場合、CurrentUserIdを使用
        if ($userId -eq "me" -and $script:CurrentUserId) {
            $userId = $script:CurrentUserId
            Write-DetailLog "'me'エンドポイントを避け、具体的なユーザーID '$userId' を使用します" -Level INFO
        }
        
        # 現在のユーザーと同じか確認
        $currentContext = Get-MgContext
        $isSelf = ($userId -eq $script:CurrentUserId -or $userId -eq "me")
        
        # 一般ユーザーで他のユーザーのOneDriveにアクセスしようとしている場合
        if (-not $script:IsAdmin -and -not $isSelf -and -not $SelfOnly) {
            Write-DetailLog "管理者権限がないため、他のユーザー '$userId' のOneDrive情報は取得できません" -Level WARNING
            return $null
        }
        
        # ユーザーの存在確認
        try {
            Write-DetailLog "ユーザー $userId のOneDrive情報を取得中..." -Level INFO
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            if (-not $user) {
                Write-DetailLog "ユーザー $userId が見つかりません" -Level ERROR
                return $null
            }
        }
        catch {
            if ($_.Exception.Message -match "Unauthorized" -or $_.Exception.Message -match "権限がありません") {
                Write-DetailLog "ユーザー情報の取得に必要な権限がありません。管理者権限が必要です。" -Level ERROR
                return $null
            }
            Write-DetailLog "ユーザー情報の取得中にエラーが発生: $($_.Exception.Message)" -Level ERROR
            return $null
        }

        # 有効なアクセストークンを確保
        $token = Get-ValidAccessToken
        if (-not $token) {
            Write-DetailLog "有効なアクセストークンがありません。再認証が必要です。" -Level ERROR
            
            # 一般ユーザーの場合の追加情報
            if (-not $script:IsAdmin) {
                Write-DetailLog "一般ユーザー権限では一部の情報にアクセスできません。自分のOneDrive情報のみ取得可能です。" -Level INFO
                
                # 自分自身のOneDriveのみ取得を試みる
                if ($isSelf) {
                    Write-DetailLog "自分のOneDrive情報の取得を続行します..." -Level INFO
                } else {
                    return $null
                }
            } else {
                return $null
            }
        }
        
        try {
            # OneDriveの情報を取得（APIエンドポイントの例、実際の実装に合わせて調整）
            $driveData = Get-MgUserDrive -UserId $userId -ErrorAction Stop
            
            if (-not $driveData) {
                Write-DetailLog "ユーザー $userId のOneDriveが見つからないか、アクセスできません" -Level WARNING
                return $null
            }
            
            # OneDriveの詳細情報を取得
            $quotaInfo = $driveData.Quota
            
            # 返却するデータを構築
            $oneDriveInfo = [PSCustomObject]@{
                ユーザーID = $userId
                氏名 = "$($user.DisplayName)"
                ログオンアカウント名 = "$($user.UserPrincipalName)".Split('@')[0]
                メールアドレス = "$($user.Mail)"
                合計容量GB = [math]::Round((Get-SafePropertyValue -InputObject $quotaInfo -PropertyName 'Total' -DefaultValue 0) / 1GB, 2)
                使用容量GB = [math]::Round((Get-SafePropertyValue -InputObject $quotaInfo -PropertyName 'Used' -DefaultValue 0) / 1GB, 2)
                残り容量GB = [math]::Round((Get-SafePropertyValue -InputObject $quotaInfo -PropertyName 'Remaining' -DefaultValue 0) / 1GB, 2)
                使用率 = [math]::Round((Get-SafePropertyValue -InputObject $quotaInfo -PropertyName 'Used' -DefaultValue 0) / (Get-SafePropertyValue -InputObject $quotaInfo -PropertyName 'Total' -DefaultValue 1) * 100, 1)
                状態 = "$($quotaInfo.State)"
                DriveId = "$($driveData.Id)"
                WebUrl = "$($driveData.WebUrl)"
            }
            
            return $oneDriveInfo
        }
        catch {
            # エラーの種類に応じたメッセージを表示
            if ($_.Exception.Message -match "Unauthorized" -or $_.Exception.Message -match "権限がありません") {
                Write-DetailLog "OneDrive情報へのアクセス権限がありません。必要なスコープ: Files.Read.All" -Level ERROR
            } else {
                Write-DetailLog "OneDrive情報の取得中にエラーが発生: $($_.Exception.Message)" -Level ERROR
            }
            Write-ErrorLog $_ "OneDrive情報取得エラー"
            return $null
        }
    }
    catch {
        Write-DetailLog "OneDrive情報の取得中にエラーが発生: $($_.Exception.Message)" -Level ERROR
        Write-ErrorLog $_ "OneDrive処理中のエラー"
        return $null
    }
}

# HTML、CSS、JavaScriptのテンプレートを作成する関数
function Create-HTMLTemplate {
    param (
        [string]$Title = "OneDrive ステータス",
        [string]$JsFileName
    )
    
    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.2.2/css/buttons.dataTables.min.css">
    <style>
        body {
            font-family: 'メイリオ', 'Meiryo', sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            text-align: center;
            margin-bottom: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #0078d4;
            color: white;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .status-ok {
            color: green;
            font-weight: bold;
        }
        .status-warning {
            color: orange;
            font-weight: bold;
        }
        .status-error {
            color: red;
            font-weight: bold;
        }
        .report-info {
            margin: 20px 0;
            padding: 10px;
            background-color: #e6f7ff;
            border-left: 5px solid #0078d4;
        }
        .chart-container {
            margin: 20px 0;
            border: 1px solid #ddd;
            padding: 10px;
            border-radius: 5px;
        }
        .usage-bar {
            height: 20px;
            background: linear-gradient(to right, #4CAF50, #FFEB3B, #F44336);
            border-radius: 5px;
            margin-top: 5px;
        }
        .footer {
            margin-top: 30px;
            text-align: center;
            color: #666;
            font-size: 0.8em;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>OneDrive ステータスレポート</h1>
        
        <div class="report-info">
            <p><strong>生成日時：</strong> <span id="reportDate"></span></p>
            <p><strong>取得アカウント：</strong> <span id="reportAccount"></span></p>
        </div>
        
        <div class="chart-container">
            <h2>使用状況サマリー</h2>
            <div id="summaryChart"></div>
        </div>
        
        <table id="statusTable" class="display">
            <thead>
                <tr>
                    <th>氏名</th>
                    <th>ログオンアカウント名</th>
                    <th>メールアドレス</th>
                    <th>状態</th>
                    <th>割当容量(GB)</th>
                    <th>使用容量(GB)</th>
                    <th>残容量(GB)</th>
                    <th>使用率</th>
                    <th>最終更新日時</th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <!-- データはJavaScriptで挿入されます -->
            </tbody>
        </table>
        
        <div class="footer">
            <p>このレポートはPowerShellスクリプトによって自動生成されました。</p>
            <p>© 2025 OneDrive運用管理ツール</p>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/dataTables.buttons.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.3/jszip.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.print.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.colVis.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="$JsFileName"></script>
</body>
</html>
"@

    return $htmlTemplate
}

# JavaScriptのテンプレートを作成する関数
function Create-JSTemplate {
    param (
        [array]$Data,
        [string]$Account,
        [datetime]$GeneratedDate
    )
    
    # JavaScriptにデータを渡すためJSON形式に変換
    $jsonData = $Data | ConvertTo-Json -Depth 5
    
    # `$(`など特殊文字を含む部分はバッククォートでエスケープ
    $jsTemplate = @"
// OneDrive Status Report JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // レポート情報を設定
    document.getElementById('reportDate').textContent = '${GeneratedDate.ToString("yyyy/MM/dd HH:mm:ss")}';
    document.getElementById('reportAccount').textContent = '${Account}';
    
    // JSONデータをJavaScriptオブジェクトに変換
    const userData = ${jsonData};
    
    // テーブルデータを構築
    const tableBody = document.getElementById('tableBody');
    let totalUsers = 0;
    let activeUsers = 0;
    let totalStorage = 0;
    let usedStorage = 0;
    
    userData.forEach(user => {
        const row = document.createElement('tr');
        
        // 氏名
        const nameCell = document.createElement('td');
        nameCell.textContent = user.氏名 || 'N/A';
        row.appendChild(nameCell);
        
        // ログオンアカウント名
        const accountCell = document.createElement('td');
        accountCell.textContent = user.ログオンアカウント名 || 'N/A';
        row.appendChild(accountCell);
        
        // メールアドレス
        const emailCell = document.createElement('td');
        emailCell.textContent = user.メールアドレス || 'N/A';
        row.appendChild(emailCell);
        
        // 状態
        const statusCell = document.createElement('td');
        statusCell.textContent = user.状態 || 'N/A';
        if (user.状態 === '有効') {
            statusCell.className = 'status-ok';
            activeUsers++;
        } else {
            statusCell.className = 'status-error';
        }
        row.appendChild(statusCell);
        
        // 割当容量(GB)
        const allocatedCell = document.createElement('td');
        allocatedCell.textContent = user.割当容量GB || 'N/A';
        row.appendChild(allocatedCell);
        
        // 使用容量(GB)
        const usedCell = document.createElement('td');
        usedCell.textContent = user.使用容量GB || 'N/A';
        row.appendChild(usedCell);
        
        // 残容量(GB)
        const remainingCell = document.createElement('td');
        remainingCell.textContent = user.残容量GB || 'N/A';
        row.appendChild(remainingCell);
        
        // 使用率
        const usageCell = document.createElement('td');
        if (typeof user.使用率 === 'number') {
            usageCell.textContent = user.使用率 + '%';
            if (user.使用率 > 90) {
                usageCell.className = 'status-error';
            } else if (user.使用率 > 70) {
                usageCell.className = 'status-warning';
            } else {
                usageCell.className = 'status-ok';
            }
        } else {
            usageCell.textContent = user.使用率 || 'N/A';
        }
        row.appendChild(usageCell);
        
        // 最終更新日時
        const lastModifiedCell = document.createElement('td');
        lastModifiedCell.textContent = user.最終更新日時 || 'N/A';
        row.appendChild(lastModifiedCell);
        
        // データ集計
        totalUsers++;
        if (user.状態 === '有効' && !isNaN(parseFloat(user.使用容量GB))) {
            if (!isNaN(parseFloat(user.割当容量GB))) {
                totalStorage += parseFloat(user.割当容量GB);
            } else {
                totalStorage += 1024; // 1TBをデフォルト値とする
            }
            usedStorage += parseFloat(user.使用容量GB);
        }
        
        tableBody.appendChild(row);
    });
    
    // DataTablesの初期化
    `$(document).ready(function() {
        `$('#statusTable').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.13.1/i18n/ja.json"
            },
            dom: 'Bfrtip',
            buttons: [
                'copy', 'csv', 'excel', 'pdf', 'print', 'colvis'
            ],
            pageLength: 25,
            order: [[3, 'desc'], [7, 'desc']]
        });
    });
    
    // 使用状況サマリーチャート
    if (document.getElementById('summaryChart')) {
        const usagePercentage = totalStorage > 0 ? (usedStorage / totalStorage * 100).toFixed(1) : 0;
        const activePercentage = totalUsers > 0 ? (activeUsers / totalUsers * 100).toFixed(1) : 0;
        
        const summaryHTML = `
            <div style="margin: 20px 0;">
                <div style="display: flex; justify-content: space-between;">
                    <div style="flex: 1; margin-right: 20px;">
                        <h3>ユーザー状態</h3>
                        <p>総ユーザー数: <strong>``${totalUsers}</strong></p>
                        <p>有効ユーザー数: <strong>``${activeUsers}</strong> (``${activePercentage}%)</p>
                        <div class="usage-bar" style="width: ``${activePercentage}%"></div>
                    </div>
                    <div style="flex: 1;">
                        <h3>ストレージ使用状況</h3>
                        <p>総割当容量: <strong>``${totalStorage.toFixed(1)} GB</strong></p>
                        <p>総使用容量: <strong>``${usedStorage.toFixed(1)} GB</strong> (``${usagePercentage}%)</p>
                        <div class="usage-bar" style="width: ``${usagePercentage}%"></div>
                    </div>
                </div>
            </div>
        `;
        
        document.getElementById('summaryChart').innerHTML = summaryHTML;
    }
});
"@

    return $jsTemplate
}

# CSV生成関数
function Export-ToCsv {
    param (
        [array]$Data,
        [string]$Path
    )
    
    try {
        # BOM付きUTF-8でCSV出力
        $Data | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
        Write-DetailLog "CSVファイルを出力しました: $Path" -Level INFO
        return $true
    }
    catch {
        Write-DetailLog "CSVファイル出力中にエラー: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# HTML生成関数
function Export-ToHtml {
    param (
        [string]$HtmlTemplate,
        [string]$Path
    )
    
    try {
        # UTF-8でHTML出力
        $HtmlTemplate | Out-File -FilePath $Path -Encoding UTF8
        Write-DetailLog "HTMLファイルを出力しました: $Path" -Level INFO
        return $true
    }
    catch {
        Write-DetailLog "HTMLファイル出力中にエラー: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# JavaScript生成関数
function Export-ToJS {
    param (
        [string]$JsTemplate,
        [string]$Path
    )
    
    try {
        # UTF-8でJavaScript出力
        $JsTemplate | Out-File -FilePath $Path -Encoding UTF8
        Write-DetailLog "JavaScriptファイルを出力しました: $Path" -Level INFO
        return $true
    }
    catch {
        Write-DetailLog "JavaScriptファイル出力中にエラー: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# スクリプト開始時に表示
Show-AppRegistrationInstructions

# エンコーディングを一時切り替え
Set-EncodingEnvironment

# メイン処理ブロック
try {
    # エラーハンドリング設定
    $ErrorActionPreference = "Continue"
    
    # 実行時刻を使ってベースのファイル名を決定
    $scriptStartTime = Get-Date -Format "yyyyMMddHHmmss"
    $fileNameBase = "OneDriveStatus_${scriptStartTime}"
    
    # 出力フォルダの作成（日付ベース）
    $outputFolderName = "OneDriveStatus." + (Get-Date -Format "yyyyMMdd")
    $outputPath = Join-Path "C:\kitting\OneDrive運用ツール" $outputFolderName
    
    # フォルダが既に存在する場合は削除しない（追記モードに変更）
    if (-not (Test-Path $outputPath)) {
        New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
    }
    
    # エラーログの初期化（スクリプト開始時に必ず実行）
    $script:logFilePath = Join-Path $outputPath ("${fileNameBase}.log")
    
    # 出力ファイルのパスを設定（すべて出力フォルダ内に）
    $htmlFilePath = Join-Path $outputPath ("${fileNameBase}.html")
    $csvFilePath = Join-Path $outputPath ("${fileNameBase}.csv")
    $jsFilePath = Join-Path $outputPath ("${fileNameBase}.js")
    
    # 相対パスのJSファイル名（HTMLに埋め込むため）
    $jsFileNameOnly = "${fileNameBase}.js"

    Write-DetailLog "出力フォルダを作成しました: $outputPath" -Level INFO
    Write-DetailLog "出力ファイルの設定:" -Level INFO
    Write-DetailLog "- ログ: $script:logFilePath" -Level INFO
    Write-DetailLog "- HTML: $htmlFilePath" -Level INFO
    Write-DetailLog "- CSV: $csvFilePath" -Level INFO
    Write-DetailLog "- JS: $jsFilePath" -Level INFO
    
    # グローバルエラーハンドラーの設定
    $Global:ErrorActionPreference = "Continue"
    trap {
        Write-ErrorLog $_ "予期しないエラーが発生しました。"
        continue
    }

    # モジュールのインストールと読み込み
    if (-not (Install-RequiredModules)) {
        throw "必要なモジュールの設定に失敗しました。"
    }

    # Microsoft Graph APIへの接続
    try {
        Write-Host "Microsoft Graph API への接続を開始します..." -ForegroundColor Yellow
        
        # 使用できる機能に合わせてスコープを調整
        $fullScopes = @(
            "User.Read.All",        # すべてのユーザー情報の読み取り
            "Files.Read.All",       # すべてのファイルの読み取り
            "Directory.Read.All",   # ディレクトリの読み取り
            "Sites.Read.All"        # すべてのサイトの読み取り
        )
        
        # 安全な接続を実行
        try {
            $context = Connect-MgGraphSafely -RequiredScopes $fullScopes
        }
        catch {
            Write-ErrorLog $_ "Microsoft Graph API への接続に失敗しました。"
            Write-Host "Microsoft Graph API への接続に失敗しました。処理を中止します。" -ForegroundColor Red
            Exit 1
        }

        # 接続情報を表示
        Write-DetailLog "接続アカウント情報:" -Level INFO
        Write-DetailLog "- アカウント: $($context.Account)" -Level INFO
        Write-DetailLog "- アプリケーション: $($context.AppName)" -Level INFO
        Write-DetailLog "- テナント: $($context.TenantId)" -Level INFO
        Write-DetailLog "- トークン有効期限: $($context.ExpiresOn)" -Level INFO

        # ユーザー情報の取得
        Write-DetailLog "ユーザー情報の取得を開始します..." -Level INFO
        
        $results = @()
        $processedCount = 0
        $successCount = 0
        $failureCount = 0
        
        # 管理者権限があればすべてのユーザー、なければ自分のみ
        if ($script:IsAdmin) {
            Write-DetailLog "管理者権限を使用して全ユーザーのOneDrive情報を取得します。" -Level INFO
            try {
                $users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,OnPremisesSamAccountName -ErrorAction Stop
                Write-DetailLog "合計 $($users.Count) 人のユーザーが見つかりました。" -Level INFO
            }
            catch {
                Write-DetailLog "ユーザー一覧の取得中にエラーが発生: $($_.Exception.Message)" -Level ERROR
                Write-DetailLog "現在のユーザーのみでスキャンを実行します。" -Level WARNING
                $users = @()
                if ($script:CurrentUserId) {
                    $user = Get-MgUser -UserId $script:CurrentUserId -Property Id,UserPrincipalName,DisplayName,OnPremisesSamAccountName
                    $users += $user
                }
                else {
                    Write-DetailLog "現在のユーザーIDを取得できませんでした。" -Level ERROR
                    throw "ユーザー情報を取得できません。"
                }
            }
        }
        else {
            Write-DetailLog "一般ユーザー権限のため、自分のOneDrive情報のみ取得します。" -Level WARNING
            $users = @()
            if ($script:CurrentUserId) {
                $user = Get-MgUser -UserId $script:CurrentUserId -Property Id,UserPrincipalName,DisplayName,OnPremisesSamAccountName
                $users += $user
            }
            else {
                Write-DetailLog "現在のユーザーIDを取得できませんでした。" -Level ERROR
                throw "ユーザー情報を取得できません。"
            }
            Write-DetailLog "自分のユーザーアカウント情報を取得しました。" -Level INFO
        }

        foreach ($user in $users) {
            $processedCount++
            Write-DetailLog "ユーザー $($user.UserPrincipalName) のOneDrive情報を取得中... ($processedCount/$($users.Count))" -Level INFO
            
            # 自分自身の場合は特別処理
            $isSelf = ($user.Id -eq $script:CurrentUserId)
            if ($isSelf) {
                $driveInfo = Get-OneDriveStatus -userId $user.Id -SelfOnly
            } else {
                $driveInfo = Get-OneDriveStatus -userId $user.Id
            }

            if ($driveInfo) {
                $successCount++
                
                # クォータ情報を取得
                $quota = $driveInfo.Quota
                $state = Get-SafePropertyValue -InputObject $driveInfo -PropertyName "State" -DefaultValue "不明"
                $lastModified = Get-SafePropertyValue -InputObject $driveInfo -PropertyName "LastModifiedDateTime" -DefaultValue $null
                
                # 変数を安全に取得
                $quotaTotal = Get-SafePropertyValue -InputObject $quota -PropertyName "Total" -DefaultValue 0
                $quotaUsed = Get-SafePropertyValue -InputObject $quota -PropertyName "Used" -DefaultValue 0
                $quotaRemaining = Get-SafePropertyValue -InputObject $quota -PropertyName "Remaining" -DefaultValue 0
                
                # GB単位に変換
                if ($quotaTotal -eq 0) {
                    $totalGB = "無制限（1TB）"
                    $usedGB = [math]::Round($quotaUsed / 1GB, 2)
                    $remainingGB = "無制限（1TB-使用量）"
                    $usagePercentage = "計測不可"
                } else {
                    $totalGB = [math]::Round($quotaTotal / 1GB, 2)
                    $usedGB = [math]::Round($quotaUsed / 1GB, 2)
                    $remainingGB = [math]::Round($quotaRemaining / 1GB, 2)
                    $usagePercentage = if ($quotaTotal -gt 0) { [math]::Round(($quotaUsed / $quotaTotal) * 100, 1) } else { 0 }
                }
                
                $status = "有効"
                
                # 最終更新日を整形
                if ($lastModified) {
                    $lastModifiedFormatted = (Get-Date $lastModified -Format "yyyy/MM/dd HH:mm:ss")
                } else {
                    $lastModifiedFormatted = "未更新"
                }
                
                Write-DetailLog "  ステータス: $status, 割当: $totalGB GB, 使用: $usedGB GB ($usagePercentage%)" -Level INFO
            } else {
                $failureCount++
                $totalGB = "取得不可"
                $usedGB = "取得不可"
                $remainingGB = "取得不可"
                $usagePercentage = "取得不可"
                $status = "未設定"
                $lastModifiedFormatted = "未更新"
                
                Write-DetailLog "  OneDrive情報を取得できませんでした。" -Level WARNING
            }

            # 結果を配列に追加
            $results += [PSCustomObject]@{
                ユーザーID = $user.Id
                氏名 = $user.DisplayName
                ログオンアカウント名 = $user.OnPremisesSamAccountName
                メールアドレス = $user.UserPrincipalName
                状態 = $status
                割当容量GB = $totalGB
                使用容量GB = $usedGB
                残容量GB = $remainingGB
                使用率 = $usagePercentage
                最終更新日時 = $lastModifiedFormatted
            }
        }

        Write-DetailLog "ユーザー情報の取得が完了しました。成功: $successCount, 失敗: $failureCount" -Level INFO

        # CSVファイル出力(BOM付き UTF-8)
        Write-DetailLog "CSVファイルの出力を開始します..." -Level INFO
        Export-ToCsv -Data $results -Path $csvFilePath
        
        # JavaScript生成
        Write-DetailLog "JavaScriptファイルの出力を開始します..." -Level INFO
        $jsContent = Create-JSTemplate -Data $results -Account $context.Account -GeneratedDate (Get-Date)
        Export-ToJS -JsTemplate $jsContent -Path $jsFilePath
        
        # HTML生成
        Write-DetailLog "HTMLファイルの出力を開始します..." -Level INFO
        $htmlContent = Create-HTMLTemplate -Title "OneDrive ステータスレポート" -JsFileName $jsFileNameOnly
        Export-ToHtml -HtmlTemplate $htmlContent -Path $htmlFilePath
        
        # 出力完了メッセージ
        Write-DetailLog "すべてのファイル出力が完了しました。" -Level INFO
        Write-DetailLog "レポートの保存先: $outputPath" -Level INFO
        
        # レポートを自動的に開く
        Start-Process $htmlFilePath
    }
    catch {
        Write-ErrorLog $_ "処理中にエラーが発生しました。"
        Exit 1
    }
}
catch {
    Write-ErrorLog $_ "予期しないエラーが発生しました。"
    Exit 1
}
finally {
    # 元のエンコーディングに戻す
    Restore-OriginalEncoding
    
    Write-Host "処理が完了しました。終了するには任意のキーを押してください..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}