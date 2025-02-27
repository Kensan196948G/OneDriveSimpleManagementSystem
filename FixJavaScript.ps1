# JavaScriptテンプレートのエラーを修正するためのヘルパースクリプト

# 指定されたJSファイルを読み込み、PowerShellの変数展開問題を修正して保存する
param (
    [Parameter(Mandatory=$true)]
    [string]$JsFilePath
)

# ファイルの存在確認
if (-not (Test-Path $JsFilePath)) {
    Write-Host "エラー: 指定されたファイル ($JsFilePath) が見つかりません。" -ForegroundColor Red
    exit 1
}

try {
    # ファイルを読み込む
    $jsContent = Get-Content -Path $JsFilePath -Raw -Encoding UTF8

    # PowerShellによって誤って解釈される可能性のある記号を修正
    $fixedContent = $jsContent -replace '\$\(document\)', '$(document)'
    $fixedContent = $fixedContent -replace '\$\(', '$('
    $fixedContent = $fixedContent -replace '\$\{([^}]+)\}', '${$1}'

    # 修正したコンテンツを同じファイルに保存
    $fixedContent | Out-File -FilePath $JsFilePath -Encoding UTF8 -Force

    Write-Host "JavaScriptファイルを修正しました: $JsFilePath" -ForegroundColor Green
    return $true
}
catch {
    Write-Host "JavaScriptファイルの修正中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    return $false
}
