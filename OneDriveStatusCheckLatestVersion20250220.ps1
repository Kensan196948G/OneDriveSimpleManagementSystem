# OneDriveStatusCheckLatestVersion20250220.ps1
# HTML末尾に <script src="OneDriveStatus_XXXX.js"></script> を正しく配置し
# CSV出力や検索・印刷などDataTablesの機能が動作するようにする
# 
# 更新履歴:
# 2025/02/25 - Microsoft Graph API権限スコープ設定の強化、認証処理の改善

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
        $script:OutputEncoding = [System.Text.Encoding]::UTF8
        
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
        $script:OutputEncoding = $script:originalOutputEncoding
        
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

# スクリプト開始時に表示
Show-AppRegistrationInstructions

# エンコーディングを一時切り替え
Set-EncodingEnvironment

# フォルダパスチェックを追加
$expectedPath = "C:\kitting\OneDrive運用ツール"
if (-not (Test-Path $expectedPath)) {
    Write-Host "エラー: フォルダ '$expectedPath' が見つかりません。" -ForegroundColor Red
    Write-Host "Enterキーを押して終了します..."
    [void][System.Console]::ReadLine()
    exit 1
}

try {
    $ErrorActionPreference = "Continue"
    
    # 実行時刻を使ってベースのファイル名を決定
    $scriptStartTime = Get-Date -Format "yyyyMMddHHmmss"
    $fileNameBase = "OneDriveStatus_${scriptStartTime}"
    
    # 出力フォルダの作成（日付ベース）
    $outputFolderName = "OneDriveStatus." + (Get-Date -Format "yyyyMMdd")
    $outputPath = Join-Path $expectedPath $outputFolderName
    
    # フォルダが既に存在する場合は削除して再作成
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
    
    # エラーログの初期化（スクリプト開始時に必ず実行）
    $script:logFilePath = Join-Path $outputPath ("${fileNameBase}.log")
    
    # 出力ファイルのパスを設定（すべて出力フォルダ内に）
    $htmlFilePath = Join-Path $outputPath ("${fileNameBase}.html")
    $csvFilePath = Join-Path $outputPath ("${fileNameBase}.csv")
    $jsFilePath = Join-Path $outputPath ("${fileNameBase}.js")

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

    # モジュールのインストールと読み込み
    if (-not (Install-RequiredModules)) {
        throw "必要なモジュールの設定に失敗しました。"
    }

    try {
        Write-Host "Microsoft Graph API への接続を開始します..." -ForegroundColor Yellow
        
        # OneDrive関連の必要最小限のスコープを定義
        # 必要最小限のスコープに絞ることで承認されやすくなります
        $requiredScopes = @(
            # 基本的な権限
            "User.Read.All",        # ユーザー情報の読み取り
            "Files.Read.All"        # すべてのファイルの読み取り
        )
        
        # 追加の権限が必要な場合は以下をコメント解除
        <#
        $additionalScopes = @(
            "Directory.Read.All",           # ディレクトリの読み取り
            "Files.ReadWrite.All",         # すべてのファイルの読み書き
            "Files.Read.Selected",         # 選択したファイルの読み取り
            "Files.ReadWrite.Selected",    # 選択したファイルの読み書き
            "Files.ReadWrite.AppFolder",   # アプリフォルダーのファイル読み書き
            "Sites.Read.All",             # すべてのサイトの読み取り
            "Sites.ReadWrite.All",        # すべてのサイトの読み書き
            "Sites.Manage.All",           # すべてのサイトの管理
            "Sites.FullControl.All",      # すべてのサイトの完全制御
            "Sites.Selected"              # 選択したサイトへのアクセス
        )
        
        # 追加の権限が必要な場合はコメント解除
        # $requiredScopes += $additionalScopes
        #>

        # 既存のパーミッションチェックと安全な接続処理（強化版）
        function Connect-MgGraphSafely {
            param (
                [array]$RequiredScopes
            )

            try {
                Write-DetailLog "認証プロセスを開始します..." -Level INFO
                Write-DetailLog "グローバル管理者アカウントでサインインしてください。" -Level INFO
                Write-DetailLog "1. Microsoftアカウント（例: username@mirai-const.co.jp）を入力" -Level INFO
                Write-DetailLog "2. HENGEONEのパスワードを入力" -Level INFO
                
                # 既存の接続を確認
                $existingContext = Get-MgContext
                if ($existingContext) {
                    Write-DetailLog "既存の接続を検出: $($existingContext.Account)" -Level INFO
                    
                    # ユーザーの役割を確認
                    try {
                        $currentUser = Get-MgUser -UserId $existingContext.Account -ErrorAction Stop
                        Write-DetailLog "現在のユーザー: $($currentUser.DisplayName)" -Level INFO
                        
                        # 既存のスコープを詳細表示
                        Write-DetailLog "現在の権限スコープ:" -Level INFO
                        foreach ($scope in $existingContext.Scopes) {
                            Write-DetailLog " - $scope" -Level INFO
                        }
                        
                        # 既存のスコープで十分か確認
                        $missingScopes = $RequiredScopes | Where-Object { $existingContext.Scopes -notcontains $_ }
                        if ($missingScopes.Count -gt 0) {
                            Write-DetailLog "以下の権限が不足しています:" -Level WARNING
                            foreach ($scope in $missingScopes) {
                                Write-DetailLog " - $scope" -Level WARNING
                            }
                            Write-DetailLog "再認証を行います..." -Level INFO
                            Disconnect-MgGraph -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 2
                        } else {
                            Write-DetailLog "必要な権限はすべて付与されています。" -Level INFO
                            return $existingContext
                        }
                    }
                    catch {
                        Write-DetailLog "ユーザー情報の取得に失敗しました。再認証を行います... エラー: $($_.Exception.Message)" -Level WARNING
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }
                }

                # 認証方法の選択
                Write-Host "`n認証方法を選択してください：" -ForegroundColor Yellow
                Write-Host "1: 通常の認証（推奨）- Microsoftアカウントとパスワードを入力" -ForegroundColor Green
                Write-Host "2: デバイスコード認証 - コードを使用して別デバイスで認証" -ForegroundColor Green
                Write-Host "3: 管理者同意付きデバイスコード認証（問題解決用）" -ForegroundColor Cyan
                $authChoice = Read-Host "`n選択 (1, 2, or 3)"

                # 認証パラメータの設定
                $connectParams = @{
                    Scopes = $RequiredScopes
                    ErrorAction = "Stop"
                }

                switch ($authChoice) {
                    "2" {
                        $connectParams["UseDeviceAuthentication"] = $true
                    }
                    "3" {
                        $connectParams["UseDeviceAuthentication"] = $true
                        # 管理者の同意を強制的に要求するフラグを追加
                        $connectParams["ForceConsent"] = $true
                    }
                }

                # 認証を実行
                Write-DetailLog "以下の権限スコープで認証を開始します:" -Level INFO
                foreach ($scope in $RequiredScopes) {
                    Write-DetailLog " - $scope" -Level INFO
                }
                
                Connect-MgGraph @connectParams
                $newContext = Get-MgContext

                if (-not $newContext) {
                    throw "認証コンテキストを取得できません。"
                }

                # 取得したスコープを確認
                Write-DetailLog "認証成功: $($newContext.Account)" -Level INFO
                Write-DetailLog "取得した権限スコープ:" -Level INFO
                foreach ($scope in $newContext.Scopes) {
                    Write-DetailLog " - $scope" -Level INFO
                }
                
                # 必要なスコープがすべて取得できたか確認
                $stillMissingScopes = $RequiredScopes | Where-Object { $newContext.Scopes -notcontains $_ }
                if ($stillMissingScopes.Count -gt 0) {
                    Write-DetailLog "警告: 以下の権限が取得できませんでした:" -Level WARNING
                    foreach ($scope in $stillMissingScopes) {
                        Write-DetailLog " - $scope" -Level WARNING
                    }
                    Write-DetailLog "一部の機能が制限される可能性があります。" -Level WARNING
                }
                
                return $newContext
            }
            catch {
                Write-DetailLog "認証プロセスでエラーが発生しました: $($_.Exception.Message)" -Level ERROR
                if ($_.Exception.InnerException) {
                    Write-DetailLog "詳細エラー: $($_.Exception.InnerException.Message)" -Level ERROR
                }
                throw
            }
        }

        # 安全な接続を実行
        try {
            $context = Connect-MgGraphSafely -RequiredScopes $requiredScopes
        }
        catch {
            Write-ErrorLog $_ "Microsoft Graph API への接続に失敗しました。"
            throw
        }

        # 接続情報を表示
        Write-DetailLog "接続アカウント情報:" -Level INFO
        Write-DetailLog "- アカウント: $($context.Account)" -Level INFO
        Write-DetailLog "- アプリケーション: $($context.AppName)" -Level INFO
        Write-DetailLog "- テナント: $($context.TenantId)" -Level INFO
        Write-DetailLog "- トークン有効期限: $($context.ExpiresOn)" -Level INFO

    }
    catch {
        Write-ErrorLog $_ "Microsoft Graph API への接続でエラーが発生しました。"
        Exit 1
    }

    function Get-OneDriveStatus {
        param (
            [string]$userId,
            [int]$retryCount = 3,
            [int]$retryDelaySeconds = 2
        )
        
        try {
            # まずユーザーの存在確認
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            if (-not $user) {
                Write-DetailLog "ユーザー $userId が見つかりません。" -Level ERROR
                return $null
            }

            # 現在のアクセストークンを取得
            $context = Get-MgContext
            if (-not $context -or -not $context.AccessToken) {
                Write-DetailLog "有効なアクセストークンがありません。再認証が必要です。" -Level ERROR
                return $null
            }

            try {
                # v1.0 APIでの取得を試行
                $drive = Get-MgUserDrive -UserId $userId -ErrorAction Stop
                Write-DetailLog "ユーザー $userId のOneDrive情報を取得しました。" -Level INFO
                return $drive
            }
            catch {
                Write-DetailLog "標準APIでの取得に失敗: $($_.Exception.Message)" -Level WARNING
                Write-DetailLog "代替方法を試行します..." -Level WARNING
                
                for ($i = 1; $i -le $retryCount; $i++) {
                    try {
                        # アクセストークンが古い場合に備えて更新
                        $refreshedContext = Get-MgContext
                        
                        # 代替方法：RESTを直接使用
                        $apiUrl = "https://graph.microsoft.com/v1.0/users/$userId/drive"
                        $headers = @{
                            "Authorization" = "Bearer $($refreshedContext.AccessToken)"
                            "Content-Type" = "application/json"
                        }
                        
                        Write-DetailLog "REST API呼び出し試行 $i/$retryCount: $apiUrl" -Level INFO
                        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get -ErrorAction Stop
                        Write-DetailLog "REST APIでユーザー $userId のOneDrive情報を取得しました。" -Level INFO
                        return $response
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        Write-DetailLog "REST API呼び出し失敗 ($i/$retryCount): $errorMessage" -Level ERROR
                        
                        if ($i -lt $retryCount) {
                            Write-DetailLog "$retryDelaySeconds 秒後に再試行します..." -Level INFO
                            Start-Sleep -Seconds $retryDelaySeconds
                        }
                    }
                }
                
                Write-DetailLog "すべての試行が失敗しました。ユーザー $userId のOneDrive情報を取得できません。" -Level ERROR
                return $null
            }
        }
        catch {
            Write-ErrorLog $_ "ユーザー $userId のOneDrive情報取得に失敗しました。"
            return $null
        }
    }

    # ユーザー情報の取得
    Write-DetailLog "ユーザー情報の取得を開始します..." -Level INFO
    $users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,OnPremisesSamAccountName
    Write-DetailLog "合計 $($users.Count) 人のユーザーが見つかりました。" -Level INFO
    $results = @()
    $processedCount = 0
    $successCount = 0
    $failureCount = 0

    foreach ($user in $users) {
        $processedCount++
        Write-DetailLog "ユーザー $($user.UserPrincipalName) のOneDrive情報を取得中... ($processedCount/$($users.Count))" -Level INFO
        $driveInfo = Get-OneDriveStatus -userId $user.Id

        if ($driveInfo) {
            $successCount++
            $quota = $driveInfo.Quota
            $state = $driveInfo.State
            $lastModified = $driveInfo.LastModifiedDateTime
            # 変数を安全に取得
            $quotaTotal = if ($null -ne $quota -and $null -ne $quota.Total) { $quota.Total } else { 0 }
            $quotaUsed = if ($null -ne $quota -and $null -ne $quota.Used) { $quota.Used } else { 0 }
            $quotaRemaining = if ($null -ne $quota -and $null -ne $quota.Remaining) { $quota.Remaining } else { 0 }
            
            if ($null -eq $quota -or $null -eq $quotaTotal -or 0 -eq $quotaTotal) {
                $totalGB = "無制限（1TB）"
                # 1GBは1073741824バイト
                $usedGB = if ($quotaUsed -gt 0) { [math]::Round($quotaUsed / 1073741824, 2) } else { 0 }
                $remainingGB = "無制限（1TB）"
                $usagePercentage = "計測不可"
            } else {
                # 1GBは1073741824バイト
                $totalGB = [math]::Round($quotaTotal / 1073741824, 2)
                $usedGB = [math]::Round($quotaUsed / 1073741824, 2)
                $remainingGB = [math]::Round($quotaRemaining / 1073741824, 2)
                $usagePercentage = if ($quotaTotal -gt 0) { [math]::Round($quotaUsed * 100 / $quotaTotal, 1) } else { 0 }
            }
            $status = "有効"
            if ($lastModified) {
                $lastModifiedFormatted = (Get-Date $lastModified -Format "yyyy/MM/dd HH:mm:ss")
            } else {
                $lastModifiedFormatted = "未更新"
            }
        } else {
            $failureCount++
            $totalGB = "取得不可"
            $usedGB = "取得不可"
            $remainingGB = "取得不可"
            $usagePercentage = "取得不可"
            $status = "未設定"
            $lastModifiedFormatted = "未更新"
        }

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
    $results | ForEach-Object {
        # 文字化け防止のために各フィールドをUTF-8エンコードで処理
        $_.PSObject.Properties | ForEach-Object {
            if ($_.Value) {
                $_.Value = [string]$_.Value
            }
        }
    } | Export-Csv -Path $csvFilePath -Encoding UTF8 -NoTypeInformation
    
    # BOMを確実に付与するために再度書き込み
    $csvRaw = Get-Content $csvFilePath -Raw -Encoding UTF8
    $utf8BOM = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($csvFilePath, $csvRaw, $utf8BOM)
    Write-DetailLog "CSVファイルの出力が完了しました: $csvFilePath" -Level INFO

    # JSファイルの内容 (DataTables の設定など)
    Write-DetailLog "JavaScriptファイルの出力を開始します..." -Level INFO
    $jsContent = @'
$(document).ready(function() {
    var table = $("#dataTable").DataTable({
        language: {
            url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/ja.json"
        },
        dom: 'Bfrtip',
        buttons: [
            {
                extend: 'csv',
                text: 'CSV出力',
                charset: 'utf-8',
                bom: true,
                filename: "REPLACE_CSV_FILENAME",
                exportOptions: {
                    columns: ':visible'
                }
            },
            {
                extend: 'excel',
                text: 'Excel出力',
                filename: "REPLACE_CSV_FILENAME",
                exportOptions: {
                    columns: ':visible'
                }
            },
            {
                extend: 'print',
                text: '印刷',
                exportOptions: {
                    columns: ':visible'
                }
            }
        ],
        pageLength: 25,
        orderCellsTop: true,
        fixedHeader: true,
        initComplete: function () {
            this.api().columns().every(function (index) {
                var column = this;
                var headerCell = $(".filters th").eq($(column.header()).index());
                var select = $('<select class="column-filter"><option value="">全て表示</option></select>')
                    .appendTo(headerCell)
                    .on('change', function () {
                        var val = $.fn.dataTable.util.escapeRegex($(this).val());
                        column.search(val ? '^' + val + '$' : '', true, false).draw();
                    });
                var uniqueValues = [];
                column.data().each(function (value, i) {
                    if (uniqueValues.indexOf(value) === -1) {
                        uniqueValues.push(value);
                    }
                });
                uniqueValues.sort();
                uniqueValues.forEach(function (value) {
                    if (value !== '') {
                        select.append('<option value="' + value + '">' + value + '</option>');
                    }
                });
            });
        }
    });
    // Enterキーで検索実行
    $('#dataTable_filter input').unbind();
    $('#dataTable_filter input').bind('keyup', function(e) {
        if(e.keyCode === 13) {
            table.search(this.value).draw();
        }
    });
});
'@
    # JS中のREPLACE_CSV_FILENAMEを fileNameBase に置き換え
    $jsContent = $jsContent -replace "REPLACE_CSV_FILENAME", $fileNameBase

    # JSファイルを BOMなしUTF-8 で保存
    [System.IO.File]::WriteAllText($jsFilePath, $jsContent, [System.Text.Encoding]::UTF8)
    Write-DetailLog "JavaScriptファイルの出力が完了しました: $jsFilePath" -Level INFO

    # HTML 出力 (BOM付き UTF-8)
    Write-DetailLog "HTMLファイルの出力を開始します..." -Level INFO
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>OneDrive ステータス</title>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/select/1.3.4/css/select.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.2.2/css/buttons.dataTables.min.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/select/1.3.4/js/dataTables.select.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/dataTables.buttons.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.3/jszip.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.2.2/js/buttons.print.min.js"></script>
    <style>
        .dataTables_wrapper {
            padding: 20px;
            font-family: 'メイリオ', Meiryo, sans-serif;
        }
        table.dataTable thead th {
            background-color: #f2f2f2;
            padding: 10px;
            border: 1px solid #ddd;
        }
        table.dataTable tbody td {
            padding: 8px;
            border: 1px solid #ddd;
        }
        .dataTables_filter {
            margin-bottom: 15px;
        }
        .dt-buttons {
            margin-bottom: 15px;
        }
        .column-filter {
            width: 100%;
            padding: 5px;
            margin-top: 5px;
            box-sizing: border-box;
        }
        .dataTables_filter input {
            padding: 5px;
            border: 1px solid #ccc;
            border-radius: 4px;
        }
        .buttons-csv, .buttons-excel, .buttons-print {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-right: 5px;
        }
        .buttons-csv:hover, .buttons-excel:hover, .buttons-print:hover {
            background-color: #45a049;
        }
    </style>
</head>
<body>
    <h1>OneDrive ステータスレポート</h1>
    <table id="dataTable" class="display">
        <thead>
            <tr>
                <th>ユーザーID</th>
                <th>氏名</th>
                <th>ログオンアカウント名</th>
                <th>メールアドレス</th>
                <th>状態</th>
                <th>割当容量GB</th>
                <th>使用容量GB</th>
                <th>残容量GB</th>
                <th>使用率</th>
                <th>最終更新日時</th>
            </tr>
            <tr class="filters">
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
            </tr>
        </thead>
        <tbody>
"@

    # テーブル行を追加
    $htmlBody = ""
    foreach ($r in $results) {
        $htmlBody += @"
            <tr>
                <td>$($r.ユーザーID)</td>
                <td>$($r.氏名)</td>
                <td>$($r.ログオンアカウント名)</td>
                <td>$($r.メールアドレス)</td>
                <td>$($r.状態)</td>
                <td>$($r.割当容量GB)</td>
                <td>$($r.使用容量GB)</td>
                <td>$($r.残容量GB)</td>
                <td>$($r.使用率)</td>
                <td>$($r.最終更新日時)</td>
            </tr>
"@
    }

    # HTMLフッター
    $htmlFooter = @"
        </tbody>
    </table>
    <script src="${fileNameBase}.js"></script>
</body>
</html>
"@

    # HTML全体を結合して出力
    $htmlContent = $htmlHeader + $htmlBody + $htmlFooter
    $utf8BOM = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($htmlFilePath, $htmlContent, $utf8BOM)
    Write-DetailLog "HTMLファイルの出力が完了しました: $htmlFilePath" -Level INFO

    # 出力フォルダを開く
    Write-DetailLog "出力フォルダを開きます: $outputPath" -Level INFO
    Start-Process "explorer.exe" -ArgumentList $outputPath
}
catch {
    Write-ErrorLog $_ "処理中にエラーが発生しました。"
    Write-Host "エラーが発生しました。詳細はログファイルを確認してください: $script:logFilePath" -ForegroundColor Red
}
finally {
    # エンコーディングを元に戻す
    Restore-OriginalEncoding
    Write-Host "処理が完了しました。Enterキーを押して終了します..." -ForegroundColor Green
    [void][System.Console]::ReadLine()
}
