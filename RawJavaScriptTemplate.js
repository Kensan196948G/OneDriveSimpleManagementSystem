// OneDrive Status Report JavaScript
// このファイルはOneDriveStatusCheck.ps1の中でテンプレートとして使用されるJavaScriptです
// 単体でのテストやトラブルシューティングに使用できます

document.addEventListener('DOMContentLoaded', function() {
    // レポート情報を設定（実際の値はPowerShellから注入されます）
    document.getElementById('reportDate').textContent = '[生成日時]';
    document.getElementById('reportAccount').textContent = '[アカウント名]';
    
    // JSONデータをJavaScriptオブジェクトに変換（実際のデータはPowerShellから注入されます）
    const userData = [
        {
            "氏名": "テストユーザー",
            "ログオンアカウント名": "testuser",
            "メールアドレス": "testuser@example.com", 
            "状態": "有効",
            "割当容量GB": 1024,
            "使用容量GB": 256,
            "残容量GB": 768,
            "使用率": 25,
            "最終更新日時": "2025/02/25 12:34:56"
        }
    ];
    
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
    $(document).ready(function() {
        $('#statusTable').DataTable({
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
                        <p>総ユーザー数: <strong>${totalUsers}</strong></p>
                        <p>有効ユーザー数: <strong>${activeUsers}</strong> (${activePercentage}%)</p>
                        <div class="usage-bar" style="width: ${activePercentage}%"></div>
                    </div>
                    <div style="flex: 1;">
                        <h3>ストレージ使用状況</h3>
                        <p>総割当容量: <strong>${totalStorage.toFixed(1)} GB</strong></p>
                        <p>総使用容量: <strong>${usedStorage.toFixed(1)} GB</strong> (${usagePercentage}%)</p>
                        <div class="usage-bar" style="width: ${usagePercentage}%"></div>
                    </div>
                </div>
            </div>
        `;
        
        document.getElementById('summaryChart').innerHTML = summaryHTML;
    }
});
