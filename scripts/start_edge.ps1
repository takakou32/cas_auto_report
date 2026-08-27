<#
.SYNOPSIS
    デバッグポート付きでEdgeを起動する

.DESCRIPTION
    既存のEdgeプロセスを確認し、デバッグポート付きでEdgeを起動する。
    既にEdgeが起動している場合は警告を出す。

.PARAMETER Port
    CDPデバッグポート番号（デフォルト: 9222）
#>

param(
    [int]$Port = 9222
)

# 既存のEdgeプロセスを確認
$existingEdge = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
if ($existingEdge) {
    Write-Warning "Edgeが既に起動しています。デバッグポートを有効にするには、全てのEdgeを閉じてから再実行してください。"
    Write-Host ""
    Write-Host "全Edgeを終了するには:"
    Write-Host "  Stop-Process -Name msedge -Force"
    Write-Host ""
    $confirm = Read-Host "Edgeを強制終了して続行しますか？ (y/N)"
    if ($confirm -eq "y") {
        Stop-Process -Name msedge -Force
        Start-Sleep -Seconds 2
    } else {
        Write-Host "中断しました。"
        exit 1
    }
}

# デバッグポート付きでEdgeを起動
Write-Host "Edge起動中 (CDP port: $Port)..."
Start-Process "msedge.exe" "--remote-debugging-port=$Port"

# ポートの待機
Write-Host "CDPポートの起動を待機中..."
$maxRetry = 10
for ($i = 0; $i -lt $maxRetry; $i++) {
    Start-Sleep -Seconds 1
    try {
        $response = Invoke-RestMethod "http://localhost:$Port/json/version" -ErrorAction Stop
        Write-Host "CDP接続確認OK"
        Write-Host "  Browser: $($response.Browser)"
        Write-Host "  WebSocket: $($response.webSocketDebuggerUrl)"
        Write-Host ""
        Write-Host "次のステップ: exeからアプリを起動してください"
        exit 0
    } catch {
        Write-Host "  待機中... ($($i + 1)/$maxRetry)"
    }
}

Write-Warning "CDPポートの起動を確認できませんでした。手動で確認してください:"
Write-Host "  Invoke-RestMethod http://localhost:$Port/json"
