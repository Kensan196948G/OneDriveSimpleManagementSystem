# ⚠️ OneDrive運用ツールのトラブルシューティングガイド

## 🔒 Access Denied エラーの対処方法

### 🚫 エラー内容
```
[ERROR] ユーザー XXX のOneDrive取得中にエラー: [accessDenied] : Access denied
```

### 🔍 原因
このエラーは以下の理由で発生する可能性があります:

1. 👑 ログインしている管理者アカウントに十分な権限がない
2. 🔐 対象ユーザーのOneDriveに特別なアクセス制限が設定されている
3. 🆕 ユーザーのOneDriveがまだプロビジョニングされていない
4. 🔑 APIのアクセス許可が不十分である

### 🛠️ 解決方法

#### 1️⃣ 管理者権限の確認
- 👑 Microsoft 365管理センターで、使用しているアカウントに「SharePoint管理者」または「グローバル管理者」の役割が割り当てられていることを確認してください。
- 🔄 必要に応じて、より高い権限を持つアカウントで再実行してください。

#### 2️⃣ スコープ（権限）の確認と再認証
スクリプトを実行する際に、より広範な権限で再認証を行います:

1. 🗑️ 認証キャッシュをクリア:
```powershell
Disconnect-MgGraph
```

2. 🔄 スクリプトを再実行し、認証画面で「同意する」ボタンを押して全ての権限を許可

#### 3️⃣ オプション：個別ユーザーをスキップする設定
特定のユーザーでエラーが継続する場合は、スキップする設定を使用できます:
1. ⚙️ スクリプト実行時にパラメーターを追加:
```powershell
.\OneDriveStatusCheck.ps1 -SkipErrorUsers $true
```

#### 4️⃣ OneDriveのプロビジョニング状態の確認
🆕 対象ユーザーがOneDriveを一度も使用していない場合、OneDriveがプロビジョニングされていない可能性があります。
該当ユーザーに一度OneDriveにログインしてもらうか、管理者が代わりにプロビジョニングを行ってください。

#### 5️⃣ Microsoft Graph APIのアクセス許可の確認
🔐 Azure ADポータルでアプリケーションの権限を確認し、以下の権限が付与されていることを確認:
- 📂 Files.Read.All
- 🌐 Sites.Read.All
- 👥 User.Read.All
- 📚 Directory.Read.All

## 📚 その他のトラブルシューティング情報

より詳細なトラブルシューティング情報については、MANUALの「⚠️ トラブルシューティング」セクションを参照してください。

### 🔍 よくあるその他の問題

#### 🔤 文字化けの問題
- 症状: スクリプト実行時に日本語が文字化けする
- 解決策: 
  1. メインメニューから「文字化け修正ツール起動」を選択
  2. または `修正ツール/EncodingFixer.bat` を実行してエンコーディングを修正

#### ⏱️ 処理時間が長い
- 症状: 大規模環境で実行すると処理に時間がかかる
- 解決策: 夜間や週末など、システム負荷の少ない時間帯に実行を計画

#### 🔄 認証エラーが繰り返し発生
- 症状: 認証を何度試みてもエラーが発生する
- 解決策: 
  1. `Disconnect-MgGraph` で認証キャッシュをクリア
  2. PowerShellを管理者として再起動
  3. 再度認証を試行
