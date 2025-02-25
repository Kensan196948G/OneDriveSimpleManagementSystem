# OneDriveStatusCheck 実行順序

## 概要
OneDriveユーザーステータスと使用容量確認ツールの実行手順

## 実行順序

1. **管理者権限でPowerShell起動**
   - PowerShellを右クリック→「管理者として実行」

2. **作業ディレクトリ移動**
   ```
   cd C:\kitting\OneDrive運用ツール
   ```

3. **実行ポリシー変更**
   ```
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
   - プロンプトに「Y」と回答

4. **スクリプト実行**
   ```
   .\OneDriveStatusCheck.ps1
   ```

5. **アプリ権限確認**
   - アプリ権限の説明が表示される
   - 「y」を入力するとAzureポータルの権限設定ページへ移動可能

6. **認証方法選択**
   - 「1」: 通常認証
   - 「2」: デバイスコード認証

7. **サインイン処理**
   - Microsoft 365アカウントでサインイン
   - 必要に応じてMFA認証完了

8. **データ収集・処理**
   - スクリプトがOneDriveデータを収集・処理
   - 処理状況がコンソールに表示

9. **レポート生成**
   - 処理完了後、以下の場所にファイルが生成:
     ```
     C:\kitting\OneDrive運用ツール\OneDriveStatus.YYYYMMDD\
     ```
   - 出力ファイル:
     - OneDriveStatus_YYYYMMDDHHmmss.html
     - OneDriveStatus_YYYYMMDDHHmmss.csv
     - OneDriveStatus_YYYYMMDDHHmmss.js
     - OneDriveStatus_YYYYMMDDHHmmss.log

10. **レポート確認**
    - HTMLファイルをブラウザで開く
    - CSVファイルをExcelで開く

11. **終了処理**
    - 完了メッセージ表示後、任意のキーで終了

## エラー発生時

- 認証エラー → 管理者権限でログイン確認
- モジュールエラー → `Install-Module Microsoft.Graph -Force` を実行
- タイムアウト → インターネット接続確認、再実行
- 実行エラー → ログファイル確認

---

*最終更新: 2025年3月10日*
