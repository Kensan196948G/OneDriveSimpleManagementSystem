# PowerShellのエンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "PowerShellのエンコーディングが UTF-8 に設定されました。"
Write-Host "現在の出力エンコーディング: $($OutputEncoding.EncodingName)"
Write-Host "現在のコンソール出力エンコーディング: $([Console]::OutputEncoding.EncodingName)"
Write-Host "例: カルビ、実行するには実行ポリシーを変更する必要があります。"
Write-Host ""
Write-Host "PowerShellの実行ポリシーを一時的に変更するには、管理者権限でPowerShellを開き、次のコマンドを実行します："
Write-Host "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass"
