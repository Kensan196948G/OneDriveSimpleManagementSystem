# OneDriveStatusCheckLatestVersion20250220.ps1 実行順序

## 1. 初期化処理とエラーハンドリング設定
1. PowerShellセッション内の全関数を削除（再定義防止）
2. ログ出力関数（Write-DetailLog）の定義
3. エラーハンドリング関数（Write-ErrorLog）の定義
4. グローバルエラーハンドラーの設定
5. エンコーディング関連関数の定義と実行
   - Set-EncodingEnvironment
   - Restore-OriginalEncoding
6. エンコーディングをUTF-8に一時切り替え

## 2. 環境チェックと出力準備
1. フォルダパスの存在確認（C:\kitting\OneDrive運用ツール）
2. エラーログの初期化
   - 日付ベースのフォルダ作成
   - ログファイルの作成と初期化
3. 出力フォルダの作成
   - フォルダ名: OneDriveStatus.YYYYMMDD
   - 既存フォルダがある場合は削除して再作成
4. 出力ファイルの設定
   - ログファイル: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.log
   - HTMLファイル: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.html
   - CSVファイル: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.csv
   - JavaScriptファイル: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.js

## 3. Microsoft Graph モジュールの設定
1. 必要なモジュールのリスト確認
   - Microsoft.Graph.Users
   - Microsoft.Graph.Files
   - Microsoft.Graph.Authentication
2. 既存のモジュールセッションのクリーンアップ
3. 必要なモジュールのインストールと読み込み

## 4. Microsoft Graph API 認証
1. 必要なスコープの定義
2. 既存の接続確認と必要に応じてクリーンアップ
3. ユーザーに認証方法の選択を提示
   - 通常の認証（推奨）
   - デバイスコード認証
4. 選択された方法での認証実行
5. 認証コンテキストの取得と検証
6. 必要な権限の確認

## 5. OneDrive情報の取得
1. 全ユーザーの基本情報を取得
2. 各ユーザーに対して：
   - ユーザーの存在確認
   - OneDrive情報の取得試行
   - エラー時の代替方法実行
3. 取得したデータの整形
   - 容量情報の計算
   - 日時のフォーマット
   - ステータス判定

## 6. レポート生成
1. CSV出力
   - データのCSV形式への変換
   - BOM付きUTF-8でファイル出力
   - 出力先: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.csv

2. JavaScript出力
   - DataTables設定の生成
   - イベントハンドラの設定
   - 出力先: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.js

3. HTML出力
   - ヘッダー部分の生成
   - テーブル構造の作成
   - データ行の追加
   - フッター部分の追加
   - 出力先: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.html

4. ログファイル
   - 処理の詳細ログを記録
   - 出力先: ./OneDriveStatus.YYYYMMDD/OneDriveStatus_YYYYMMDDHHMMSS.log

## 7. 終了処理
1. 処理結果の表示
2. エンコーディングを元の設定に復元
3. 完了メッセージの表示

## エラーハンドリング
1. エラー情報の記録
   - エラーの種類
   - エラーメッセージ
   - 発生場所
   - スタックトレース
   - 内部エラー情報

2. エラーレベル別の処理
   - INFO: 通常の処理状況
   - WARNING: 警告（処理は継続）
   - ERROR: エラー（状況に応じて処理中断）

3. エラーログの出力
   - タイムスタンプ付きでログファイルに記録
   - コンソールへの色分け表示
   - エラーサマリーの生成

4. エラー発生時の対応
   - OneDrive情報取得エラー: 該当ユーザーをスキップ
   - 認証エラー: 再認証を試行
   - 致命的エラー: エラーログを記録して終了

5. 終了処理
   - エラー情報の要約出力
   - エラー総数の表示
   - ログファイルパスの表示
   - エンコーディングの復元

## 注意事項
- スクリプト実行には管理者権限が必要
- Microsoft Graph API の認証には適切な権限が必要
- エンコーディングの変更により文字化けが発生する可能性がある場合は即時対応

## 出力ファイル構成
すべてのファイルは日付ベースのフォルダ内に出力
```
OneDriveStatus.YYYYMMDD/
├── OneDriveStatus_YYYYMMDDHHMMSS.log  # 処理ログとエラー情報
├── OneDriveStatus_YYYYMMDDHHMMSS.html # レポート本体
├── OneDriveStatus_YYYYMMDDHHMMSS.js   # DataTables設定
└── OneDriveStatus_YYYYMMDDHHMMSS.csv  # データエクスポート
