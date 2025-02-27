# 📋 OneDriveStatusCheck 実行順序

## 📝 概要
OneDriveユーザーステータスと使用容量確認ツールの実行手順

## 🔄 実行順序

### 1️⃣ メインメニューからの実行（推奨）
   - 🚀 `OneDriveReportShortcut.bat` をダブルクリック
   - 📋 メニューから「1: OneDriveステータスレポート作成」を選択

### 2️⃣ 管理者権限でPowerShell起動（直接実行する場合）
   - 👑 PowerShellを右クリック→「管理者として実行」

### 3️⃣ 作業ディレクトリ移動
   ```powershell
   cd C:\kitting\OneDrive運用ツール
   ```

### 4️⃣ 実行ポリシー変更
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
   - ✅ プロンプトに「Y」と回答

### 5️⃣ スクリプト実行
   ```powershell
   .\OneDriveStatusCheck.ps1
   ```

### 6️⃣ アプリ権限確認
   - 🔐 アプリ権限の説明が表示される
   - 👆 「y」を入力するとAzureポータルの権限設定ページへ移動可能

### 7️⃣ 認証方法選択
   - 🔑 「1」: 通常認証
   - 📱 「2」: デバイスコード認証

### 8️⃣ サインイン処理
   - 👤 Microsoft 365アカウントでサインイン
   - 🔒 必要に応じてMFA認証完了

### 9️⃣ データ収集・処理
   - 📊 スクリプトがOneDriveデータを収集・処理
   - 📈 処理状況がコンソールに表示

### 🔟 レポート生成
   - 📁 処理完了後、以下の場所にファイルが生成:
     ```
     C:\kitting\OneDrive運用ツール\OneDriveStatus.YYYYMMDD\
     ```
   - 📄 出力ファイル:
     - 🌐 OneDriveStatus_YYYYMMDDHHmmss.html
     - 📊 OneDriveStatus_YYYYMMDDHHmmss.csv
     - 📜 OneDriveStatus_YYYYMMDDHHmmss.js
     - 📝 OneDriveStatus_YYYYMMDDHHmmss.log

### 1️⃣1️⃣ レポート確認
   - 🌐 HTMLファイルをブラウザで開く
   - 📊 CSVファイルをExcelで開く

### 1️⃣2️⃣ 終了処理
   - ✅ 完了メッセージ表示後、任意のキーで終了

## ⚠️ エラー発生時

### 🚨 よくあるエラーと対処法

#### 1️⃣ 認証エラー
   - 🚫 **エラーメッセージ**: `アクセストークンの取得に失敗しました` または `有効なアクセストークンがありません`
   - 🔧 **対処法**: 
     - 🔄 一度サインアウトし、再度認証する
     - 👑 管理者アカウントを使用する
     - 🔍 Azureポータルでアプリ権限を確認する

#### 2️⃣ モジュールエラー
   - 🚫 **エラーメッセージ**: `'Microsoft.Graph' モジュールがありません`
   - 🔧 **対処法**: 
     ```powershell
     Install-Module Microsoft.Graph -Force
     Install-Module Microsoft.Graph.Authentication -Force
     ```

#### 3️⃣ Nullエラー
   - 🚫 **エラーメッセージ**: `Cannot bind argument to parameter 'InputObject' because it is null`
   - 🔧 **対処法**:
     - 👑 一般ユーザーのOneDriveにアクセスできない場合は、管理者アカウントで再試行
     - 📝 ログファイルを確認して詳細なエラー情報を収集

#### 4️⃣ タイムアウト
   - 🚫 **エラーメッセージ**: `操作がタイムアウトしました`
   - 🔧 **対処法**:
     - 🌐 インターネット接続を確認
     - 🔌 会社のプロキシ設定を確認
     - ⏱️ しばらく時間をおいてから再試行

### 🔍 詳細なトラブルシューティング

エラーが継続する場合:

1. 📝 ログファイル(`OneDriveStatus_*.log`)を確認
2. 🗑️ 以下のコマンドで認証状態をクリア:
   ```powershell
   Disconnect-MgGraph
   Clear-MgContext
   ```
3. 🔄 PowerShellを再起動して再試行

---

*📅 最終更新: 2025年3月15日*
