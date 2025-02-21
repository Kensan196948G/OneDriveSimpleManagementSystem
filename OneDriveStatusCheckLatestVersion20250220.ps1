# OneDriveStatusCheckLatestVersion20250220.ps1
# HTML末尾に <script src="OneDriveStatus_XXXX.js"></script> を正しく配置し
# CSV出力や検索・印刷などDataTablesの機能が動作するようにする

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

function Set-EncodingEnvironment {
    $script:originalOutputEncoding = [Console]::OutputEncoding
    $script:originalInputEncoding  = [Console]::InputEncoding
    $currentOutputEncoding         = [Console]::OutputEncoding
    $currentOutputCodePage         = $currentOutputEncoding.CodePage

    Write-Host "現在の出力エンコーディング: $($currentOutputEncoding.EncodingName) (CodePage: $($currentOutputEncoding.CodePage))" -ForegroundColor Yellow
    if ($currentOutputCodePage -ne 65001) {
        Write-Host "出力エンコーディングをUTF-8に変更しています..." -ForegroundColor Yellow
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
        $newOutputEncoding        = [Console]::OutputEncoding
        Write-Host "変更後の出力エンコーディング: $($newOutputEncoding.EncodingName) (CodePage: $($newOutputEncoding.CodePage))" -ForegroundColor Green
    } else {
        Write-Host "出力エンコーディングはすでにUTF-8です。変更は不要です。" -ForegroundColor Green
    }
}

function Restore-OriginalEncoding {
    if ($script:originalOutputEncoding -ne $null) {
        Write-Host "元のエンコーディングに戻しています..." -ForegroundColor Yellow
        [Console]::OutputEncoding = $script:originalOutputEncoding
        [Console]::InputEncoding  = $script:originalInputEncoding
        $restoredOutputEncoding   = [Console]::OutputEncoding
        Write-Host "復元後の出力エンコーディング: $($restoredOutputEncoding.EncodingName) (CodePage: $($restoredOutputEncoding.CodePage))" -ForegroundColor Green
    }
}

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

    # 出力フォルダの作成
    $outputFolderName = "OneDriveStatus." + (Get-Date -Format "yyyyMMdd")
    $outputPath = Join-Path $expectedPath $outputFolderName
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

    # 実行時刻を使ってベースのファイル名を決定
    $scriptStartTime = Get-Date -Format "yyyyMMddHHmmss"
    $fileNameBase = "OneDriveStatus_${scriptStartTime}"

    # 出力ファイルのパスを更新
    $logFilePath = Join-Path $outputPath ("${fileNameBase}.log")
    $htmlFilePath = Join-Path $outputPath ("${fileNameBase}.html")
    $csvFilePath = Join-Path $outputPath ("${fileNameBase}.csv")
    $jsFilePath = Join-Path $outputPath ("${fileNameBase}.js")

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
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
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

    # 実行時刻を使ってベースのファイル名を決定
    $scriptStartTime = Get-Date -Format "yyyyMMddHHmmss"
    $fileNameBase    = "OneDriveStatus_${scriptStartTime}"

    # スクリプト実行フォルダ
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

    # 出力ファイル(同じベース名で揃える)
    $logFilePath  = Join-Path $scriptPath ("${fileNameBase}.log")
    $htmlFilePath = Join-Path $scriptPath ("${fileNameBase}.html")
    $csvFilePath  = Join-Path $scriptPath ("${fileNameBase}.csv")
    $jsFilePath   = Join-Path $scriptPath ("${fileNameBase}.js")

    try {
        Connect-MgGraph -Scopes "User.Read.All","Files.ReadWrite.All","Sites.Read.All"
        Write-DetailLog "Microsoft Graph API へ接続できました。" -Level INFO
    }
    catch {
        Write-DetailLog "Microsoft Graph API へ接続できませんでした。" -Level ERROR
        Exit 1
    }

    function Get-OneDriveStatus {
        param ([string]$userId)
        try {
            $driveInfo = Get-MgUserDrive -UserId $userId -ErrorAction Stop
            return $driveInfo
        }
        catch {
            if ($_.Exception.Message -match "403") {
                Write-DetailLog "ユーザー $userId のOneDrive へのアクセスが拒否されました(403)。" -Level WARNING
                return $null
            } elseif ($_.Exception.Message -match "accessDenied") {
                Write-DetailLog "ユーザー $userId のOneDrive情報取得失敗: [accessDenied] : Access denied" -Level ERROR
                return $null
            } else {
                Write-DetailLog "ユーザー $userId のOneDrive情報取得失敗: $($_.Exception.Message)" -Level ERROR
                return $null
            }
        }
    }

    # ユーザー情報の取得
    $users   = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,OnPremisesSamAccountName
    $results = @()

    foreach ($user in $users) {
        Write-DetailLog "ユーザー $($user.UserPrincipalName) のOneDrive情報を取得中..."
        $driveInfo = Get-OneDriveStatus -userId $user.Id

        if ($driveInfo) {
            $quota        = $driveInfo.Quota
            $state        = $driveInfo.State
            $lastModified = $driveInfo.LastModifiedDateTime
            if ($quota.Total -eq $null -or $quota.Total -eq 0) {
                $totalGB         = "無制限（1TB）"
                $usedGB          = [math]::Round($quota.Used /1GB,2)
                $remainingGB     = "無制限（1TB）"
                $usagePercentage = "計測不可"
            } else {
                $totalGB         = [math]::Round($quota.Total /1GB,2)
                $usedGB          = [math]::Round($quota.Used  /1GB,2)
                $remainingGB     = [math]::Round($quota.Remaining /1GB,2)
                $usagePercentage = [math]::Round(($quota.Used / $quota.Total)*100,1)
            }
            $status = "有効"
            if ($lastModified) {
                $lastModifiedFormatted = (Get-Date $lastModified -Format "yyyy/MM/dd HH:mm:ss")
            } else {
                $lastModifiedFormatted = "未更新"
            }
        } else {
            $totalGB              = "取得不可"
            $usedGB               = "取得不可"
            $remainingGB          = "取得不可"
            $usagePercentage      = "取得不可"
            $status               = "未設定"
            $lastModifiedFormatted= "未更新"
        }

        $results += [PSCustomObject]@{
            ユーザーID        = $user.Id
            氏名              = $user.DisplayName
            ログオンアカウント名 = $user.OnPremisesSamAccountName
            メールアドレス     = $user.UserPrincipalName
            状態              = $status
            割当容量GB         = $totalGB
            使用容量GB         = $usedGB
            残容量GB           = $remainingGB
            使用率            = $usagePercentage
            最終更新日時       = $lastModifiedFormatted
        }
    }

    # CSVファイル出力(BOM付き UTF-8)
    $results | Export-Csv -Path $csvFilePath -Encoding UTF8 -NoTypeInformation
    $csvRaw   = Get-Content $csvFilePath -Raw
    $utf8BOM  = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($csvFilePath, $csvRaw, $utf8BOM)

    # JSファイルの内容 (DataTables の設定など)
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

    # HTML 出力 (BOM付き UTF-8)
    $htmlContent = @"
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
    foreach ($r in $results) {
        $htmlContent += @"
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
    $fullHtml = $htmlContent + $htmlFooter
    [System.IO.File]::WriteAllText($htmlFilePath, $fullHtml, (New-Object System.Text.UTF8Encoding($true)))

    Write-Host "" -ForegroundColor Green
    Write-Host "処理完了: 以下のファイルが生成されました。" -ForegroundColor Green
    Write-Host "  ログ : $logFilePath" -ForegroundColor Cyan
    Write-Host "  CSV : $csvFilePath" -ForegroundColor Cyan
    Write-Host "  HTML: $htmlFilePath" -ForegroundColor Cyan
    Write-Host "  JS  : $jsFilePath" -ForegroundColor Cyan
}
catch {
    Write-DetailLog "エラーが発生しました: $($_.Exception.Message)" -Level ERROR
    throw
}
finally {
    Restore-OriginalEncoding
    Write-Host "" -ForegroundColor Yellow
    Write-Host "エンコーディングを元に戻しました。Enterキーを押して終了します..." -ForegroundColor Yellow
    [void][System.Console]::ReadLine()
}
