#
# OneDriveStatusCheck.ps1のエラー修正スクリプト
# 構文エラーがあるOneDriveStatusCheck.ps1を修正します
#

# UTF-8エンコーディング設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 対象ファイルパス
$targetFilePath = Join-Path -Path $PSScriptRoot -ChildPath "OneDriveStatusCheck.ps1"

# バックアップを作成
$backupPath = "$targetFilePath.bak"
Copy-Item -Path $targetFilePath -Destination $backupPath -Force
Write-Host "バックアップを作成しました: $backupPath" -ForegroundColor Green

# 修正するコンテンツを準備
$fixedContent = @'
#
# OneDriveステータスチェックスクリプト
# ユーザーのOneDrive使用状況をレポートします
#

# エンコーディングの設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# パラメーター定義
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = "",
    
    [Parameter(Mandatory = $false)]
    [switch]$AllUsers = $false,
    
    [Parameter(Mandatory = $false)]
    [string]$UserEmail = "",
    
    [Parameter(Mandatory = $false)]
    [int]$MaxErrors = 10,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipConsent = $false,
    
    [Parameter(Mandatory = $false)]
    [int]$RetryCount = 3,
    
    [Parameter(Mandatory = $false)]
    [switch]$IgnoreAccessDenied = $true
)

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
$global:totalUsers = 0

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
                    Write-Log -Message "リトライ $attempt/$RetryCount`：$errorMsg" -Level "WARNING"
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
            render