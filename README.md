📂 OneDrive運用ツール

## 概要
このツールは、OneDrive のステータスをチェックし、ユーザーごとの使用状況をレポートとして出力するためのスクリプトです。レポートは CSV、HTML、JavaScript ファイルとして出力され、DataTables を使用して検索や印刷などの機能を提供します。

## 📁 ファイル構成
- `OneDriveStatus.bat`: バッチファイル。PowerShell スクリプトを実行するためのエントリーポイント。
- `OneDriveStatusCheckLatestVersion20250220.ps1`: PowerShell スクリプト。OneDrive のステータスをチェックし、レポートを生成します。
- `README.md`: このファイル。ツールの概要と利用手順を記載しています。

## 🛠 利用手順

### 1. 必要な権限の確認
Azure AD アプリケーションに以下の権限が付与されていることを確認してください。
- `User.Read.All`
- `Files.ReadWrite.All`
- `Sites.Read.All`

### 2. ツールの実行
1. `OneDriveStatus.bat` を管理者として実行します。
2. 実行ポリシーの確認メッセージが表示されます。現在のポリシーが `RemoteSigned` であることを確認してください。
3. PowerShell スクリプトが実行され、OneDrive のステータスチェックが開始されます。
4. スクリプトの実行が完了すると、以下のファイルが生成されます。
   - 📄 ログファイル: `OneDriveStatus_YYYYMMDDHHMMSS.log`
   - 📄 CSVファイル: `OneDriveStatus_YYYYMMDDHHMMSS.csv`
   - 📄 HTMLファイル: `OneDriveStatus_YYYYMMDDHHMMSS.html`
   - 📄 JavaScriptファイル: `OneDriveStatus_YYYYMMDDHHMMSS.js`

### 3. レポートの確認
生成された HTML ファイルをブラウザで開くと、DataTables を使用したレポートが表示されます。検索や印刷、CSV 出力などの機能を利用できます。

## ⚠️ 注意事項
- スクリプトを実行するには、管理者権限が必要です。
- Azure AD アプリケーションの権限が不足している場合、アクセス拒否エラーが発生することがあります。その場合は、必要な権限を追加してください。

## 🛠 トラブルシューティング
### エラー: `[accessDenied] : Access denied`
- Azure AD アプリケーションに必要な権限が付与されているか確認してください。
- 管理者に依頼して、必要な権限を追加してもらってください。

### エラー: `Function capacity 4096 has been exceeded for this scope`
- PowerShell セッション内の関数が多すぎる場合に発生します。スクリプトの冒頭でセッション内の全関数を削除する処理を追加していますが、問題が解決しない場合は、セッションを再起動してください。

### エラー: `ファイルまたはアセンブリ 'Microsoft.Graph.Authentication' が見つかりません`
- `Microsoft.Graph.Authentication` モジュールがインストールされているか確認してください。
- スクリプト内で `Install-RequiredModules` 関数を実行し、必要なモジュールがインストールされていることを確認してください。

## 📜 ライセンス
このツールは MIT ライセンスの下で提供されています。詳細は `LICENSE` ファイルを参照してください。