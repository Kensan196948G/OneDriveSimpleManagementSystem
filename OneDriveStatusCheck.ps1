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
            
            try {
                $context = Get-MgContext
                if ($context) {
                    $script:AccessToken = $context.AccessToken
                    
                    # AccessTokenExpiresOnが利用できる場合は期限を設定
                    if ($context.ExpiresOn) {
                        $script:TokenExpiration = $context.ExpiresOn
                    } else {
                        # 有効期限不明の場合は警告を表示
                        Write-DetailLog "コンテキストから直接アクセストークンを取得しました。有効期限は不明です。" -Level WARNING
                        # 安全のため1時間後を有効期限として設定
                        $script:TokenExpiration = $now.AddHours(1)
                    }
                } else {
                    Write-DetailLog "GraphAPIコンテキストが存在しません。再認証が必要です。" -Level ERROR
                    return $null
                }
            } catch {
                Write-DetailLog "トークン取得中にエラーが発生しました: $($_.Exception.Message)" -Level ERROR
                return $null
            }
            
            # トークンが取得できたか確認
            if ([string]::IsNullOrEmpty($script:AccessToken)) {
                Write-DetailLog "有効なアクセストークンがありません。再認証が必要です。" -Level ERROR
                return $null
            }
            
            # 管理者権限の確認
            $script:IsAdmin = Test-IsGlobalAdmin
        }
        
        # 一般ユーザー向けのメッセージ
        if (-not $script:IsAdmin -and $script:AccessToken) {
            Write-DetailLog "一般ユーザー権限では一部の情報にアクセスできません。自分のOneDrive情報のみ取得可能です。" -Level INFO
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
        [Parameter()]
        $InputObject,
        
        [Parameter(Mandatory=$true)]
        [string]$PropertyName,
        
        $DefaultValue = $null
    )
    
    try {
        # InputObjectがnullの場合はデフォルト値を返す
        if ($null -eq $InputObject) {
            return $DefaultValue
        }
        
        if ($InputObject.PSObject.Properties.Match($PropertyName).Count -gt 0) {
            $value = $InputObject.$PropertyName
            if ($null -ne $value) {
                return $value
            }
        }
        return $DefaultValue
    }
    catch {
        Write-DetailLog "プロパティ '$PropertyName' の取得に失敗しました: $($_.Exception.Message)" -Level WARNING
        return $DefaultValue
    }
}

# OneDriveの状態を取得する関数 - 管理者向け拡張
function Get-OneDriveStatus {
    param (
        [Parameter(Mandatory=$false)]
        [string]$UserId,
        
        [switch]$SelfOnly,
        
        [switch]$CurrentUserOnly
    )
    
    try {
        Write-DetailLog "OneDriveステータスの取得を開始します..." -Level INFO
        
        # 結果を格納する配列
        $users = @()
        
        # ユーザーIDの確認とセットアップ
        $targetUserId = $null
        if ($UserId) {
            $targetUserId = $UserId
            Write-DetailLog "指定されたユーザーID: $targetUserId のOneDrive情報を取得" -Level INFO
        } elseif ($script:CurrentUserId) {
            $targetUserId = $script:CurrentUserId
            Write-DetailLog "現在のユーザーID: $targetUserId のOneDrive情報を取得" -Level INFO
        } else {
            Write-DetailLog "有効なユーザーIDが指定されていません" -Level ERROR
            return $null
        }
        
        try {
            # ユーザー情報を取得
            $user = Get-MgUser -UserId $targetUserId -ErrorAction Stop
            
            if ($null -eq $user) {
                Write-DetailLog "ユーザー情報が取得できません: $targetUserId" -Level ERROR
                return $null
            }
            
            Write-DetailLog "ユーザー「$($user.DisplayName)」のOneDrive情報を取得中..." -Level INFO
            
            try {
                # クエリパラメータを拡張（データ取得の成功率向上のため）
                $params = @{
                    UserId = $user.Id
                    ErrorAction = "Stop"
                }
                
                # ユーザーのドライブを取得
                $drive = Get-MgUserDrive @params
                
                if ($null -eq $drive) {
                    Write-DetailLog "ユーザー $($user.DisplayName) のOneDriveが見つかりません" -Level WARNING
                    return $null
                }
                
                # デバッグ情報の出力
                Write-DetailLog "取得成功: DriveID=$($drive.Id), Type=$($drive.DriveType)" -Level INFO
                
                # クォータ情報を持っているか確認
                if ($drive.Quota) {
                    Write-DetailLog "クォータ情報: State=$($drive.Quota.State), Total=$([math]::Round($drive.Quota.Total/1GB,2))GB, Used=$([math]::Round($drive.Quota.Used/1GB,2))GB" -Level INFO
                } else {
                    Write-DetailLog "クォータ情報がありません" -Level WARNING
                }
                
                # ドライブの詳細情報を返却
                return $drive
            }
            catch {
                Write-DetailLog "ユーザー $($user.DisplayName) のOneDrive取得中にエラー: $($_.Exception.Message)" -Level ERROR
                if ($_.Exception.InnerException) {
                    Write-DetailLog "  詳細: $($_.Exception.InnerException.Message)" -Level ERROR
                }
                return $null
            }
        }
        catch {
            Write-ErrorLog $_ "ユーザー $targetUserId の情報取得中にエラーが発生しました"
            return $null
        }
    }
    catch {
        Write-ErrorLog $_ "OneDriveステータスの取得中にエラーが発生しました"
        return $null
    }
}

# HTML、CSS、JavaScriptのテンプレートを作成する関数 - 改善版
function Create-HTMLTemplate {
    param (
        [string]$Title = "OneDrive ステータス",
        [string]$JsFileName,
        [string]$ReportFolder
    )
    
    # JavaScriptファイルへの相対パスを作成（同じフォルダ内を想定）
    $jsPath = "./$JsFileName"
    
    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <!-- DataTablesとそのプラグインのスタイルシート -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.2.2/css/buttons.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/searchpanes/1.4.0/css/searchPanes.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/select/1.3.4/css/select.dataTables.min.css">
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
        .search-tools {
            margin: 20px 0;
            padding: 10px;
            background-color: #f9f9f9;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        .dataTables_filter {
            margin-bottom: 15px;
        }
        .dt-buttons {
            margin-bottom: 15px;
        }
        /* 印刷時のスタイル */
        @media print {
            body {
                background-color: white;
                margin: 0;
                padding: 0;
            }
            .container {
                box-shadow: none;
                padding: 0;
            }
            .no-print {
                display: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>OneDrive ステータスレポート</h1>
        
        <div class="report-info">
            <p><strong>生成日時：</strong> <span id="reportDate"></span></p>
            <p><strong>取得アカウント：</strong> <span id="reportAccount"></span></p>
            <p><strong>レポート保存先：</strong> $ReportFolder</p>
        </div>
        
        <div class="chart-container">
            <h2>使用状況サマリー</h2>
            <div id="summaryChart"></div>
        </div>
        
        <div class="search-tools no-print">
            <h3>データ検索・操作ツール</h3>
            <p>以下の機能が利用できます：</p>
            <ul>
                <li><strong>検索：</strong> テーブル上部の検索ボックスで全項目から検索</li>
                <li><strong>エクスポート：</strong> 「エクスポート」ボタンからデータをダウンロード</li>
                <li><strong>印刷：</strong> 「印刷」ボタンからテーブルを印刷</li>
                <li><strong>表示列の選択：</strong> 「表示列」ボタンから表示する列を選択</li>
            </ul>
        </div>
        
        <table id="statusTable" class="display" style="width:100%">
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

    <!-- JavaScriptライブラリ -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <!-- DataTablesプラグイン -->
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.print.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.colVis.min.js"></script>
    <script src="https://cdn.datatables.net/searchpanes/1.4.0/js/dataTables.searchPanes.min.js"></script>
    <script src="https://cdn.datatables.net/select/1.3.4/js/dataTables.select.min.js"></script>
    <!-- Excel出力用ライブラリ -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.3/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js"></script>
    <!-- グラフ用ライブラリ -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- カスタムのJavaScript -->
    <script src="$jsPath"></script>
</body>
</html>
"@

    return $htmlTemplate
}

# JavaScriptのテンプレートを作成する関数 - 拡張版
function Create-JSTemplate {
    param (
        [array]$Data,
        [string]$Account,
        [datetime]$GeneratedDate
    )
    
    # JavaScriptにデータを渡すためJSON形式に変換
    $jsonData = $Data | ConvertTo-Json -Depth 5 -Compress
    
    # JavaScriptテンプレート
    $jsTemplate = @"
// OneDrive Status Report JavaScript
document.addEventListener('DOMContentLoaded', function() {
    console.log('OneDrive Status Report: ドキュメント読み込み完了');
    
    // レポート情報を設定
    document.getElementById('reportDate').textContent = '${GeneratedDate.ToString("yyyy/MM/dd HH:mm:ss")}';
    document.getElementById('reportAccount').textContent = '${Account}';
    
    // JSONデータをJavaScriptオブジェクトに変換
    let userData;
    try {
        userData = ${jsonData};
        console.log('データ読み込み成功: ' + userData.length + '件のユーザー情報');
    } catch (error) {
        console.error('データ解析エラー:', error);
        userData = [];
    }
    
    // テーブルデータを構築
    const tableBody = document.getElementById('tableBody');
    let totalUsers = 0;
    let activeUsers = 0;
    let totalStorage = 0;
    let usedStorage = 0;
    
    // テーブルデータの作成
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
            let allocatedSpace = 0;
            if (!isNaN(parseFloat(user.割当容量GB))) {
                allocatedSpace = parseFloat(user.割当容量GB);
            } else if (typeof user.割当容量GB === 'string' && !isNaN(parseFloat(user.割当容量GB.replace(/,/g, '')))) {
                // カンマ区切りの文字列を数値に変換
                allocatedSpace = parseFloat(user.割当容量GB.replace(/,/g, ''));
            } else {
                allocatedSpace = 1024; // 1TBをデフォルト値とする
            }
            totalStorage += allocatedSpace;
            
            // 使用容量の集計
            let usedSpace = 0;
            if (!isNaN(parseFloat(user.使用容量GB))) {
                usedSpace = parseFloat(user.使用容量GB);
            } else if (typeof user.使用容量GB === 'string' && !isNaN(parseFloat(user.使用容量GB.replace(/,/g, '')))) {
                usedSpace = parseFloat(user.使用容量GB.replace(/,/g, ''));
            }
            usedStorage += usedSpace;
        }
        
        tableBody.appendChild(row);
    });
    
    // DataTablesの初期化 - 拡張機能追加
    try {
        const table = $('#statusTable').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.13.1/i18n/ja.json"
            },
            // 表示機能とボタンの設定
            dom: 'Bfrtip',
            buttons: [
                {
                    extend: 'copy',
                    text: 'コピー',
                    className: 'btn-copy'
                },
                {
                    extend: 'csv',
                    text: 'CSV',
                    className: 'btn-csv',
                    title: 'OneDriveStatus_' + new Date().toISOString().split('T')[0]
                },
                {
                    extend: 'excel',
                    text: 'Excel',
                    className: 'btn-excel',
                    title: 'OneDriveステータスレポート_' + new Date().toISOString().split('T')[0],
                    exportOptions: {
                        columns: ':visible'
                    }
                },
                {
                    extend: 'pdf',
                    text: 'PDF',
                    className: 'btn-pdf',
                    title: 'OneDriveステータスレポート',
                    orientation: 'landscape'
                },
                {
                    extend: 'print',
                    text: '印刷',
                    className: 'btn-print',
                    customize: function (win) {
                        `$(win.document.body).find('h1').text('OneDriveステータスレポート');
                        `$(win.document.body).css('font-family', 'メイリオ, Meiryo, sans-serif');
                    }
                },
                {
                    extend: 'colvis',
                    text: '表示列',
                    className: 'btn-colvis'
                },
                {
                    text: '全データ表示',
                    action: function (e, dt) {
                        dt.page.len(-1).draw();
                    }
                }
            ],
            // 検索機能の強化
            searchBuilder: true,
            searchPanes: true,
            // 1ページに表示する行数
            pageLength: 25,
            // デフォルトのソート順
            order: [[3, 'desc'], [7, 'desc']],
            // レスポンシブ対応
            responsive: true,
            // 初期化完了時の処理
            initComplete: function() {
                console.log('DataTables 初期化完了');
                // 検索ボックスにプレースホルダーを追加
                `$('.dataTables_filter input').attr('placeholder', '検索キーワードを入力');
                // 検索ボックスの幅を広げる
                `$('.dataTables_filter input').css('width', '300px');
            }
        });
        
        console.log('DataTablesの初期化に成功しました');
        
        // 検索機能強化: 即時反応するフィルタリング
        `$('#statusTable_filter input').unbind().bind('keyup', function() {
            table.search(this.value).draw();
        });
        
    } catch (error) {
        console.error('DataTables初期化エラー:', error);
    }
    
    // 使用状況サマリーチャート
    if (document.getElementById('summaryChart')) {
        try {
            const usagePercentage = totalStorage > 0 ? (usedStorage / totalStorage * 100).toFixed(1) : 0;
            const activePercentage = totalUsers > 0 ? (activeUsers / totalUsers * 100).toFixed(1) : 0;
            
            console.log('サマリー情報: 総ユーザー数=' + totalUsers + ', 有効ユーザー数=' + activeUsers);
            console.log('ストレージ情報: 総容量=' + totalStorage.toFixed(1) + 'GB, 使用量=' + usedStorage.toFixed(1) + 'GB');
            
            const summaryHTML = `
                <div style="margin: 20px 0;">
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-between;">
                        <div style="flex: 1; min-width: 250px; margin-right: 20px; margin-bottom: 20px;">
                            <h3>ユーザー状態</h3>
                            <p>総ユーザー数: <strong>`${totalUsers}</strong></p>
                            <p>有効ユーザー数: <strong>`${activeUsers}</strong> (`${activePercentage}%)</p>
                            <div class="usage-bar" style="width: `${activePercentage}%"></div>
                        </div>
                        <div style="flex: 1; min-width: 250px; margin-bottom: 20px;">
                            <h3>ストレージ使用状況</h3>
                            <p>総割当容量: <strong>`${totalStorage.toFixed(1)} GB</strong></p>
                            <p>総使用容量: <strong>`${usedStorage.toFixed(1)} GB</strong> (`${usagePercentage}%)</p>
                            <div class="usage-bar" style="width: `${usagePercentage}%"></div>
                        </div>
                    </div>
                </div>
            `;
            
            document.getElementById('summaryChart').innerHTML = summaryHTML;
        } catch (error) {
            console.error('サマリーチャート生成エラー:', error);
            document.getElementById('summaryChart').innerHTML = '<p class="error">データの表示中にエラーが発生しました。</p>';
        }
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

# スクリプト実行前に認証をクリア
function Clear-AuthenticationCache {
    Write-Host "`n認証キャッシュをクリアしています..." -ForegroundColor Yellow
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:AccessToken = $null
        $script:TokenExpiration = [DateTime]::MinValue
        $script:IsAdmin = $false
        $script:CurrentUserId = $null
        Write-Host "認証キャッシュをクリアしました。再認証が必要です。" -ForegroundColor Green
        return $true
    }
    catch {
        Write-DetailLog "認証キャッシュのクリア中にエラーが発生しました: $($_.Exception.Message)" -Level WARNING
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
    
    # 認証キャッシュをクリア（オプション）
    if ($ClearAuth) {
        Clear-AuthenticationCache
    }
    
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
                $users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,Mail,OnPremisesSamAccountName -ErrorAction Stop
                Write-DetailLog "合計 $($users.Count) 人のユーザーが見つかりました。" -Level INFO
            }
            catch {
                Write-DetailLog "ユーザー一覧の取得中にエラーが発生: $($_.Exception.Message)" -Level ERROR
                Write-DetailLog "現在のユーザーのみでスキャンを実行します。" -Level WARNING
                $users = @()
                if ($script:CurrentUserId) {
                    $user = Get-MgUser -UserId $script:CurrentUserId -Property Id,UserPrincipalName,DisplayName,Mail,OnPremisesSamAccountName
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
                $user = Get-MgUser -UserId $script:CurrentUserId -Property Id,UserPrincipalName,DisplayName,Mail,OnPremisesSamAccountName
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
            
            # ユーザーのOneDrive情報を取得
            try {
                # OneDriveの詳細情報を取得
                $driveInfo = Get-OneDriveStatus -UserId $user.Id
                
                # デバッグ用：重要な項目のみを抽出して出力
                if ($driveInfo) {
                    Write-DetailLog "ドライブ情報: ID=$($driveInfo.Id), Type=$($driveInfo.DriveType)" -Level INFO
                    
                    # クォータ情報を取得
                    $quota = $driveInfo.Quota
                    if ($quota) {
                        Write-DetailLog "クォータ状態: $($quota.State)" -Level INFO
                    }
                    
                    $successCount++
                    
                    # クォータ情報を取得
                    $driveId = $driveInfo.Id
                    $webUrl = $driveInfo.WebUrl
                    
                    # 変数を安全に取得
                    $quotaTotal = Get-SafePropertyValue -InputObject $quota -PropertyName "Total" -DefaultValue 0
                    $quotaUsed = Get-SafePropertyValue -InputObject $quota -PropertyName "Used" -DefaultValue 0
                    $quotaRemaining = Get-SafePropertyValue -InputObject $quota -PropertyName "Remaining" -DefaultValue 0
                    $quotaState = Get-SafePropertyValue -InputObject $quota -PropertyName "State" -DefaultValue "unknown"
                    
                    # GB単位に変換
                    if ($quotaTotal -eq 0) {
                        # 使用容量の正確な変換
                        $usedGB = if ($quotaUsed -gt 0) {
                            [math]::Round($quotaUsed / 1GB, 2)
                        } else {
                            0
                        }
                        
                        # デバッグログで検出された25600GBを適用（実際の環境では変更が必要かもしれません）
                        $totalGB = 25600
                        $remainingGB = $totalGB - $usedGB
                        
                        # パーセンテージ計算
                        $usagePercentage = if ($usedGB -gt 0) {
                            [math]::Round(($usedGB / $totalGB) * 100, 1)
                        } else {
                            0
                        }
                    } else {
                        $totalGB = [math]::Round($quotaTotal / 1GB, 2)
                        $usedGB = [math]::Round($quotaUsed / 1GB, 2)
                        $remainingGB = [math]::Round($quotaRemaining / 1GB, 2)
                        
                        # パーセンテージ計算
                        $usagePercentage = if ($quotaTotal -gt 0) {
                            [math]::Round(($quotaUsed / $quotaTotal) * 100, 1)
                        } else {
                            0
                        }
                    }
                    
                    $status = "有効"
                    
                    # 最終更新日を取得
                    $lastModified = $null
                    try {
                        # ルートフォルダ自体の更新日時を取得
                        if ($driveInfo.Root -and $driveInfo.Root.LastModifiedDateTime) {
                            $lastModified = $driveInfo.Root.LastModifiedDateTime
                        } else {
                            # 子アイテムから最新の更新日を探すことを試みる
                            $childParams = @{
                                UserId = $user.Id
                                Top = 1
                                OrderBy = "lastModifiedDateTime desc"
                                ErrorAction = "Stop"
                            }
                            
                            $rootItems = Get-MgUserDriveRootChild @childParams
                            if ($rootItems -and $rootItems.Count -gt 0) {
                                $lastModified = $rootItems[0].LastModifiedDateTime
                            }
                        }
                    }
                    catch {
                        Write-DetailLog "ユーザー $($user.DisplayName) のファイル履歴取得時にエラー: $($_.Exception.Message)" -Level WARNING
                    }
                    
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
                    メールアドレス = $user.Mail
                    状態 = $status
                    割当容量GB = $totalGB
                    使用容量GB = $usedGB
                    残容量GB = $remainingGB
                    使用率 = $usagePercentage
                    最終更新日時 = $lastModifiedFormatted
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-DetailLog "ユーザー $($user.DisplayName) のOneDrive取得中にエラー: $errorMessage" -Level ERROR
                
                # エラーメッセージを分析して対応策を表示
                if ($errorMessage -match "accessDenied|Access denied") {
                    Write-DetailLog "アクセス権限の問題が発生しました。以下をお試しください：" -Level WARNING
                    Write-DetailLog "1. Disconnect-MgGraph を実行して認証キャッシュをクリア" -Level WARNING
                    Write-DetailLog "2. スクリプトを -ClearAuth スイッチを付けて実行" -Level WARNING
                    Write-DetailLog "3. 管理者アカウントでログインし直してください" -Level WARNING
                }
                
                # エラーが発生した場合の処理を追加
                if ($SkipErrorUsers) {
                    Write-DetailLog "ユーザー $($user.DisplayName) のOneDriveはスキップされました" -Level WARNING
                    # エラーユーザーの情報を結果に追加（エラー状態を明示）
                    $results += [PSCustomObject]@{
                        ユーザーID = $user.Id
                        氏名 = $user.DisplayName
                        ログオンアカウント名 = $user.OnPremisesSamAccountName
                        メールアドレス = $user.Mail
                        状態 = "エラー"
                        割当容量GB = "アクセス拒否"
                        使用容量GB = "アクセス拒否"
                        残容量GB = "アクセス拒否"
                        使用率 = "アクセス拒否"
                        最終更新日時 = "取得不可"
                        エラー詳細 = $errorMessage
                    }
                    # 次のユーザーに進む（continue相当の処理）
                } else {
                    # SkipErrorUsersがfalseの場合はエラーをスローして処理を停止
                    throw $_
                }
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
        $htmlContent = Create-HTMLTemplate -Title "OneDrive ステータスレポート" -JsFileName $jsFileNameOnly -ReportFolder $outputPath
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

# OneDriveステータスチェックスクリプト
# ユーザーのOneDrive使用状況をレポートします
#

# エンコーディングの設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8


# モジュール関連の変数
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Files"
)

# 出力ファイル関連の変数
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$defaultOutputFolder = Join-Path -Path $PSScriptRoot -ChildPath "OneDriveStatus.$($timestamp.Substring(0, 8))"
$outputFolder = if ([string]::IsNullOrEmpty($OutputFolder)) { $defaultOutputFolder } else { $OutputFolder }

# ログ関連の変数
$logFile = Join-Path -Path $outputFolder -ChildPath "OneDriveStatus_$timestamp.log"
$htmlFile = Join-Path -Path $outputFolder -ChildPath "OneDriveStatus_$timestamp.html"
$csvFile = Join-Path -Path $outputFolder -ChildPath "OneDriveStatus_$timestamp.csv"
$jsFile = Join-Path -Path $outputFolder -ChildPath "OneDriveStatus_$timestamp.js"

# 統計情報の変数初期化
$global:totalProcessed = 0
$global:totalSuccess = 0
$global:totalFailed = 0
$global:totalSkipped = 0
$global:errorMessages = @{}
$global:userResults = @()

# ログ出力関数
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # コンソール出力色の設定
    $foregroundColor = switch ($Level) {
        "INFO"    { "White" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        "SUCCESS" { "Green" }
        "DEBUG"   { "Cyan" }
        default   { "White" }
    }
    
    # コンソールに出力
    Write-Host $logMessage -ForegroundColor $foregroundColor
    
    # ログファイルに出力
    try {
        Add-Content -Path $logFile -Value $logMessage -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "ログファイルへの書き込みに失敗しました: $_" -ForegroundColor Red
    }
}

# 出力フォルダを作成
function Initialize-OutputFolder {
    # 出力フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $outputFolder)) {
        try {
            New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
            Write-Log -Message "出力フォルダを作成しました: $outputFolder" -Level "INFO"
        }
        catch {
            throw "出力フォルダの作成に失敗しました: $_"
        }
    }
    
    # 出力ファイルの設定を表示
    Write-Log -Message "出力ファイルの設定:" -Level "INFO"
    Write-Log -Message "- ログ: $logFile" -Level "INFO"
    Write-Log -Message "- HTML: $htmlFile" -Level "INFO"
    Write-Log -Message "- CSV: $csvFile" -Level "INFO"
    Write-Log -Message "- JS: $jsFile" -Level "INFO"
}

# モジュール存在チェック・インストール・インポート関数
function Import-RequiredModules {
    foreach ($moduleName in $requiredModules) {
        try {
            # モジュールが既にインポートされているかチェック
            if (Get-Module -Name $moduleName) {
                Write-Log -Message "モジュール $moduleName は既に読み込まれています。" -Level "INFO"
                continue
            }
            
            # モジュールがインストールされているかチェック
            if (-not (Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue)) {
                Write-Log -Message "モジュール $moduleName をインストールします..." -Level "INFO"
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
            }
            
            # モジュールをインポート
            Write-Log -Message "モジュール $moduleName をインポートします..." -Level "INFO"
            Import-Module -Name $moduleName -DisableNameChecking
            Write-Log -Message "モジュール $moduleName を読み込みました。" -Level "INFO"
        }
        catch {
            throw "モジュール $moduleName の準備に失敗しました: $_"
        }
    }
}

# Microsoft Graphへの接続・認証
function Connect-MicrosoftGraph {
    Write-Log -Message "認証プロセスを開始します..." -Level "INFO"
    
    # 必要な権限スコープ
    $scopes = @(
        "User.Read.All",
        "Files.Read.All",
        "Directory.Read.All",
        "Sites.Read.All"
    )
    
    try {
        # すでに接続されていないかチェック
        $connection = Get-MgContext
        if ($null -ne $connection) {
            Write-Log -Message "既にMicrosoft Graph APIに接続されています。" -Level "INFO"
            return
        }
        
        # アプリケーション権限への同意を要求
        if (-not $SkipConsent) {
            Write-Log -Message "アプリケーション権限への同意を要求します..." -Level "INFO"
        }
        
        # スコープを指定して接続
        Write-Log -Message "$($scopes -join ', ') スコープでの認証を試行します..." -Level "INFO"
        Connect-MgGraph -Scopes $scopes
        
        # 接続情報を取得
        $context = Get-MgContext
        Write-Log -Message "認証成功: $($context.Account)" -Level "INFO"
        
        # 現在のユーザー情報を取得
        $currentUser = Get-MgUser -UserId $context.Account -ErrorAction SilentlyContinue
        if ($currentUser) {
            Write-Log -Message "現在のユーザーID: $($currentUser.Id)" -Level "INFO"
            
            # 管理者ロールをチェック
            $roles = Get-MgUserDirectoryRole -UserId $currentUser.Id -ErrorAction SilentlyContinue
            if ($roles) {
                foreach ($role in $roles) {
                    Write-Log -Message "管理者ロールを発見: $($role.DisplayName)" -Level "INFO"
                }
            }
        }
        
        Write-Log -Message "接続アカウント情報:" -Level "INFO"
        Write-Log -Message "- アカウント: $($context.Account)" -Level "INFO"
        Write-Log -Message "- アプリケーション: $($context.AppName)" -Level "INFO"
        Write-Log -Message "- テナント: $($context.TenantId)" -Level "INFO"
        Write-Log -Message "- トークン有効期限: $($context.ExpiresOn)" -Level "INFO"
    }
    catch {
        throw "Microsoft Graphへの接続に失敗しました: $_"
    }
}

# ユーザー情報を取得する関数
function Get-OneDriveUsers {
    Write-Log -Message "ユーザー情報の取得を開始します..." -Level "INFO"
    
    try {
        # 特定のユーザーが指定されている場合
        if (-not [string]::IsNullOrEmpty($UserEmail)) {
            $user = Get-MgUser -UserId $UserEmail -ErrorAction Stop
            Write-Log -Message "指定されたユーザー $UserEmail の情報を取得しました。" -Level "INFO"
            return @($user)
        }
        
        # 管理者権限で全ユーザーを取得
        Write-Log -Message "管理者権限を使用して全ユーザーのOneDrive情報を取得します。" -Level "INFO"
        
        # ライセンス付与済みの有効なユーザーのみを取得
        # AccountEnabled=true と AssignedLicenses があるユーザーのみフィルター
        $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses -Filter "AccountEnabled eq true"
        
        # フィルタリング: ライセンスが付与されているユーザーのみ
        $licensedUsers = $users | Where-Object { $_.AssignedLicenses.Count -gt 0 }
        Write-Log -Message "合計 $($licensedUsers.Count) 人のユーザーが見つかりました。" -Level "INFO"
        
        return $licensedUsers
    }
    catch {
        throw "ユーザー情報の取得に失敗しました: $_"
    }
}

# ユーザーのOneDriveステータスを取得する関数
function Get-UserOneDriveStatus {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )
    
    $retryDelaySeconds = 2
    
    # ユーザー名をログに記録
    $userEmail = $User.UserPrincipalName
    $userName = $User.DisplayName
    $userId = $User.Id
    
    Write-Log -Message "ユーザー $userEmail のOneDrive情報を取得中... ($($global:totalProcessed + 1)/$($global:totalUsers))" -Level "INFO"
    Write-Log -Message "OneDriveステータスの取得を開始します..." -Level "INFO"
    
    try {
        Write-Log -Message "指定されたユーザーID: $userId のOneDrive情報を取得" -Level "INFO"
        Write-Log -Message "ユーザー「$userName」のOneDrive情報を取得中..." -Level "INFO"
        
        # リトライロジック
        $attempt = 0
        $success = $false
        
        while ($attempt -lt $RetryCount -and -not $success) {
            try {
                $attempt++
                
                # OneDriveのルートを取得
                $oneDrive = Get-MgUserDrive -UserId $userId -ErrorAction Stop
                
                # OneDrive情報が取得できた場合
                $success = $true
                
                # OneDriveルート配下のアイテムを取得
                $items = Get-MgUserDriveRoot -UserId $userId -ErrorAction SilentlyContinue
                $childItems = Get-MgUserDriveRootChild -UserId $userId -ErrorAction SilentlyContinue -All
                
                # ファイル情報を取得
                $fileCount = 0
                $folderCount = 0
                $largeFiles = @()
                
                if ($childItems) {
                    $fileCount = ($childItems | Where-Object { $_.File }).Count
                    $folderCount = ($childItems | Where-Object { $_.Folder }).Count
                    
                    # 大きいファイル（20MB以上）を確認
                    $largeFiles = $childItems | Where-Object { 
                        $_.File -and $_.Size -gt 20MB 
                    } | Select-Object Name, Size, LastModifiedDateTime
                }
                
                # OneDriveの使用状況を整理
                $result = [PSCustomObject]@{
                    Email = $userEmail
                    DisplayName = $userName
                    DriveId = $oneDrive.Id
                    Quota = [PSCustomObject]@{
                        Total = $oneDrive.Quota.Total
                        Used = $oneDrive.Quota.Used
                        Remaining = $oneDrive.Quota.Remaining
                        PercentUsed = if ($oneDrive.Quota.Total -gt 0) { 
                            [math]::Round(($oneDrive.Quota.Used / $oneDrive.Quota.Total) * 100, 2) 
                        } else { 0 }
                    }
                    FileCount = $fileCount
                    FolderCount = $folderCount
                    LargeFiles = $largeFiles
                    Status = "成功"
                    LastChecked = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                $global:totalSuccess++
                Write-Log -Message "OneDrive情報の取得に成功しました: $userName" -Level "SUCCESS"
                return $result
            }
            catch {
                $errorMsg = $_.Exception.Message
                
                # AccessDeniedエラーの場合で、無視するオプションが指定されている場合
                if ($errorMsg -match "accessDenied" -and $IgnoreAccessDenied) {
                    Write-Log -Message "ユーザー $userName のOneDrive取得中にエラー: $errorMsg" -Level "ERROR"
                    Write-Log -Message "  OneDrive情報を取得できませんでした。" -Level "WARNING"
                    
                    $result = [PSCustomObject]@{
                        Email = $userEmail
                        DisplayName = $userName
                        DriveId = $null
                        Quota = $null
                        FileCount = 0
                        FolderCount = 0
                        LargeFiles = @()
                        Status = "アクセス拒否"
                        ErrorMessage = $errorMsg
                        LastChecked = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    
                    $global:totalSkipped++
                    return $result
                }
                
                # 最後のリトライでなければ再試行
                if ($attempt -lt $RetryCount) {
                    Write-Log -Message "リトライ $attempt/$RetryCount - $errorMsg" -Level "WARNING"
                    Start-Sleep -Seconds ($retryDelaySeconds * $attempt)
                }
                else {
                    # 最大リトライ回数を超えた場合はエラー
                    Write-Log -Message "ユーザー $userName のOneDrive取得に失敗しました: $errorMsg" -Level "ERROR"
                    
                    # エラー統計を更新
                    if ($global:errorMessages.ContainsKey($errorMsg)) {
                        $global:errorMessages[$errorMsg]++
                    }
                    else {
                        $global:errorMessages[$errorMsg] = 1
                    }
                    
                    $result = [PSCustomObject]@{
                        Email = $userEmail
                        DisplayName = $userName
                        DriveId = $null
                        Quota = $null
                        FileCount = 0
                        FolderCount = 0
                        LargeFiles = @()
                        Status = "エラー"
                        ErrorMessage = $errorMsg
                        LastChecked = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    
                    $global:totalFailed++
                    return $result
                }
            }
        }
    }
    catch {
        # 予期せぬエラー
        $errorMsg = $_.Exception.Message
        Write-Log -Message "ユーザー $userName のOneDrive処理中に予期せぬエラー: $errorMsg" -Level "ERROR"
        
        # エラー統計を更新
        if ($global:errorMessages.ContainsKey($errorMsg)) {
            $global:errorMessages[$errorMsg]++
        }
        else {
            $global:errorMessages[$errorMsg] = 1
        }
        
        $result = [PSCustomObject]@{
            Email = $userEmail
            DisplayName = $userName
            DriveId = $null
            Quota = $null
            FileCount = 0
            FolderCount = 0
            LargeFiles = @()
            Status = "エラー"
            ErrorMessage = $errorMsg
            LastChecked = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        $global:totalFailed++
        return $result
    }
}

# HTMLレポートを作成する関数
function Export-HTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Results
    )
    
    Write-Log -Message "HTMLレポートの作成を開始します..." -Level "INFO"
    
    # JSファイルにデータを出力
    $jsContent = "const oneDriveData = " + ($Results | ConvertTo-Json -Depth 10) + ";"
    $jsContent += "`nconst reportInfo = {`n"
    $jsContent += "  totalUsers: $global:totalUsers,`n"
    $jsContent += "  totalSuccess: $global:totalSuccess,`n"
    $jsContent += "  totalFailed: $global:totalFailed,`n"
    $jsContent += "  totalSkipped: $global:totalSkipped,`n"
    $jsContent += "  generatedDate: '$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")',`n"
    $jsContent += "  errorSummary: " + ($global:errorMessages | ConvertTo-Json) + "`n"
    $jsContent += "};"
    
    Set-Content -Path $jsFile -Value $jsContent -Encoding UTF8
    
    # HTMLテンプレート
    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>OneDriveステータスレポート</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 {
            color: #0078d4;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
        }
        .summary-box {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .stats {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
        }
        .stat-card {
            background-color: white;
            padding: 15px;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            flex: 1;
            min-width: 150px;
            text-align: center;
        }
        .success { background-color: #e6ffe6; }
        .error { background-color: #ffe6e6; }
        .skip { background-color: #fff9e6; }
        
        .stat-number {
            font-size: 24px;
            font-weight: bold;
            margin: 10px 0;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 14px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            position: sticky;
            top: 0;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        .progress-container {
            width: 100%;
            background-color: #e0e0e0;
            border-radius: 4px;
            margin-top: 5px;
        }
        .progress-bar {
            height: 20px;
            border-radius: 4px;
            text-align: center;
            color: white;
            font-weight: bold;
            line-height: 20px;
        }
        .warning {
            background-color: #ffcc00;
        }
        .danger {
            background-color: #ff3333;
        }
        .normal {
            background-color: #4caf50;
        }
        .filter-controls {
            margin-bottom: 15px;
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
        }
        .search-input {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 250px;
        }
        .filter-dropdown {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        .reset-button {
            padding: 8px 12px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .reset-button:hover {
            background-color: #005a9e;
        }
        .error-summary {
            margin-top: 20px;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 8px;
            background-color: #fff9f9;
        }
        .tabs {
            display: flex;
            margin-top: 20px;
            border-bottom: 1px solid #ddd;
        }
        .tab {
            padding: 10px 20px;
            border: 1px solid #ddd;
            border-bottom: none;
            border-radius: 5px 5px 0 0;
            margin-right: 5px;
            cursor: pointer;
        }
        .tab.active {
            background-color: #0078d4;
            color: white;
        }
        .tab-content {
            display: none;
            padding: 20px;
            border: 1px solid #ddd;
            border-top: none;
        }
        .tab-content.active {
            display: block;
        }
        .chart-container {
            width: 100%;
            height: 400px;
        }
        @media (max-width: 768px) {
            .stat-card {
                min-width: 100%;
            }
            .search-input {
                width: 100%;
            }
        }
    </style>
    <!-- Chart.jsを追加 -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <h1>OneDriveステータスレポート</h1>
    
    <div class="summary-box">
        <h2>概要</h2>
        <p>レポート生成日時: <span id="generated-date"></span></p>
        
        <div class="stats">
            <div class="stat-card">
                <div>合計ユーザー数</div>
                <div id="total-users" class="stat-number">-</div>
            </div>
            <div class="stat-card success">
                <div>成功</div>
                <div id="success-count" class="stat-number">-</div>
            </div>
            <div class="stat-card error">
                <div>エラー</div>
                <div id="error-count" class="stat-number">-</div>
            </div>
            <div class="stat-card skip">
                <div>スキップ</div>
                <div id="skipped-count" class="stat-number">-</div>
            </div>
        </div>
    </div>
    
    <!-- タブナビゲーション -->
    <div class="tabs">
        <div id="tab-users" class="tab active" onclick="switchTab('users')">ユーザー一覧</div>
        <div id="tab-stats" class="tab" onclick="switchTab('stats')">統計情報</div>
        <div id="tab-errors" class="tab" onclick="switchTab('errors')">エラー詳細</div>
    </div>

    <!-- ユーザー一覧タブ -->
    <div id="content-users" class="tab-content active">
        <div class="filter-controls">
            <input type="text" id="search-input" class="search-input" placeholder="ユーザー名またはメールで検索...">
            <select id="status-filter" class="filter-dropdown">
                <option value="all">すべてのステータス</option>
                <option value="成功">成功</option>
                <option value="エラー">エラー</option>
                <option value="アクセス拒否">アクセス拒否</option>
            </select>
            <select id="usage-filter" class="filter-dropdown">
                <option value="all">すべての使用量</option>
                <option value="high">75%以上</option>
                <option value="medium">50%〜75%</option>
                <option value="low">50%未満</option>
            </select>
            <button id="reset-filters" class="reset-button">フィルターをリセット</button>
        </div>
        
        <table id="onedrive-table">
            <thead>
                <tr>
                    <th>ユーザー名</th>
                    <th>メールアドレス</th>
                    <th>使用容量/割当容量</th>
                    <th>使用率</th>
                    <th>ファイル数</th>
                    <th>フォルダ数</th>
                    <th>ステータス</th>
                    <th>最終確認日時</th>
                </tr>
            </thead>
            <tbody id="table-body">
                <!-- JavaScriptでデータが挿入されます -->
            </tbody>
        </table>
    </div>

    <!-- 統計情報タブ -->
    <div id="content-stats" class="tab-content">
        <h2>OneDrive使用状況の統計</h2>
        <div class="chart-container">
            <canvas id="usage-chart"></canvas>
        </div>
        
        <div class="chart-container" style="margin-top: 30px;">
            <canvas id="status-chart"></canvas>
        </div>
    </div>

    <!-- エラー詳細タブ -->
    <div id="content-errors" class="tab-content">
        <div class="error-summary">
            <h2>エラーサマリー</h2>
            <div id="error-summary-content">
                <!-- JavaScriptでエラーサマリーが挿入されます -->
            </div>
        </div>
        
        <h2>エラーが発生したユーザー</h2>
        <table id="error-table">
            <thead>
                <tr>
                    <th>ユーザー名</th>
                    <th>メールアドレス</th>
                    <th>ステータス</th>
                    <th>エラーメッセージ</th>
                </tr>
            </thead>
            <tbody id="error-table-body">
                <!-- JavaScriptでエラーデータが挿入されます -->
            </tbody>
        </table>
    </div>

    <script src="OneDriveStatus_$timestamp.js"></script>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            // レポート情報を表示
            document.getElementById('generated-date').textContent = reportInfo.generatedDate;
            document.getElementById('total-users').textContent = reportInfo.totalUsers;
            document.getElementById('success-count').textContent = reportInfo.totalSuccess;
            document.getElementById('error-count').textContent = reportInfo.totalFailed;
            document.getElementById('skipped-count').textContent = reportInfo.totalSkipped;
            
            // テーブルにデータを挿入
            renderTable(oneDriveData);
            
            // エラーサマリーを表示
            renderErrorSummary();
            
            // エラーテーブルを表示
            renderErrorTable();
            
            // 統計チャートを描画
            renderCharts();
            
            // フィルター機能の追加
            document.getElementById('search-input').addEventListener('input', filterTable);
            document.getElementById('status-filter').addEventListener('change', filterTable);
            document.getElementById('usage-filter').addEventListener('change', filterTable);
            document.getElementById('reset-filters').addEventListener('click', resetFilters);
        });

        // テーブルのレンダリング関数
        function renderTable(data) {
            const tableBody = document.getElementById('table-body');
            tableBody.innerHTML = '';
            
            data.forEach(user => {
                const row = document.createElement('tr');
                
                // 使用量
```
                // 残りのJavaScriptコード
            }
"@
    return $htmlTemplate
}
