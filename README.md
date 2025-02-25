# 📊 OneDrive運用ツール マニュアル

## 🌟 概要
このツールは組織内のOneDriveの使用状況を確認するためのツールです。
全ユーザーのOneDriveのステータス、使用容量、最終更新日時などを取得し、レポートを作成します。

## ⚙️ 動作要件
- 🖥️ Windows 10/11
- 🔷 PowerShell 5.1以上
- 🛡️ 管理者権限
- 👤 Microsoft 365管理者アカウント（グローバル管理者推奨）

## 🚀 インストール方法
1. このフォルダを`C:\kitting\OneDrive運用ツール`に配置してください
2. 初回実行時に必要なモジュールは自動的にインストールされます

## 📝 使用方法

### OneDriveステータス確認ツールの起動
1. `OneDriveStatusCheck.bat`をダブルクリックします
2. 🔒 UAC（ユーザーアカウント制御）のダイアログが表示されたら「はい」をクリックします
3. 🔑 認証方法を選択します:
   - 🟢 **通常の認証（推奨）**: Microsoft 365管理者アカウントでサインイン
   - 🔵 **デバイスコード認証**: 別デバイスで認証コードを使用してサインイン
   - 🟣 **詳細認証**: 拡張スコープでの認証（問題解決時に使用）

4. Microsoft 365管理者アカウントでサインインします
5. 実行完了後、自動的にレポートフォルダが開きます

### 📁 出力ファイル
実行すると以下のファイルが生成されます:
- 📰 **HTMLレポート**: データテーブル形式でのレポート（フィルタリングや検索が可能）
- 📊 **CSVファイル**: Excelなどで開ける形式のデータ
- 📝 **ログファイル**: 実行ログ（トラブルシューティング用）

## ❓ トラブルシューティング

### ⚠️ エラー: ForceConsentパラメーターが見つからない
- エラーメッセージ: `パラメーター名 'ForceConsent' に一致するパラメーターが見つかりません`
- 解決策:
  - ✅ `OneDriveStatusCheckLatestVersion20250227.ps1` を使用してください
  - ✅ このバージョンでは認証の問題を修正しています

### 🔠 文字化けの問題が発生する場合
1. `CharacterEncodingTool.ps1` を使用して文字コードを変換します
   ```powershell
   .\CharacterEncodingTool.ps1 -InputFile "問題のファイル" -SourceEncoding "shift-jis" -TargetEncoding "utf-8"
   ```
2. もしくは `文字化け修正.bat` を実行してGUIで修正します

### 🔒 アクセス権限の問題
- 「アクセスが拒否されました」などのエラーが表示される場合:
  1. Azure ADポータルで権限を確認してください
  2. Microsoft Graph Command Line Toolsアプリに以下の権限が付与されていることを確認:
     - 📖 User.Read.All
     - 📂 Files.Read.All

## 📆 更新履歴
- 🆕 **2025/02/27** - ForceConsentパラメーターエラーの修正、文字化け対応の強化
- 🔄 **2025/02/25** - Microsoft Graph API権限スコープ設定の強化、認証処理の改善

## 📋 スクリプト一覧
| ファイル名 | 説明 |
|------------|------|
| 📜 OneDriveStatusCheckLatestVersion20250227.ps1 | メインスクリプト（最新版） |
| 📜 OneDriveStatusCheck.bat | 実行用バッチファイル |
| 📜 CharacterEncodingTool.ps1 | 文字コード変換ツール |
| 📜 CharacterEncodingFixer.ps1 | 文字化け診断・修正GUIツール |
| 📜 文字化け修正.bat | 文字化け修正ツール起動用バッチ |
| 📄 README.md | このマニュアル |
| 📄 文字化け対応マニュアル.md | 文字化け対応の詳細ガイド |
| 📄 OneDriveStatusCheckLatestVersion20250227実行順序.md | スクリプト実行順序の詳細説明 |