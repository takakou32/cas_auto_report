<#
.SYNOPSIS
    デバッグポート付きでEdgeを起動する

.DESCRIPTION
    既存のEdgeプロセスを確認し、デバッグポート付きでEdgeを起動する。
    既にEdgeが起動している場合は警告を出す。

.PARAMETER Port
    CDPデバッグポート番号（デフォルト: 9222）

.PARAMETER UserDataDir
    Edgeの保存場所（プロファイル）を別の場所にする。
    Edge 136以降は、普段の保存場所のままだとデバッグポートの指定が無視されるため、
    自分でEdgeを立てて繋ぐときは指定が必須（例: C:\edgework）。
    指定すると普段のEdgeとは別のウィンドウとして立ち上がるので、開いているEdgeを閉じる必要もない。
    ただし保存場所が別＝ログイン状態やお気に入りは引き継がれない。

.EXAMPLE
    .\scripts\start_edge.ps1 -UserDataDir C:\edgework
#>

param(
    [int]$Port = 9222,
    [string]$UserDataDir = ""
)

# 保存場所を別にする場合は、普段のEdgeとは別インスタンスになるので既存プロセスの確認は不要
$existingEdge = if ($UserDataDir) { $null } else { Get-Process -Name "msedge" -ErrorAction SilentlyContinue }
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
$edgeArgs = "--remote-debugging-port=$Port"
if ($UserDataDir) {
    $edgeArgs += " --user-data-dir=`"$UserDataDir`" --no-first-run --no-default-browser-check"
    Write-Host "Edge起動中 (CDP port: $Port / 保存場所: $UserDataDir)..."
} else {
    Write-Host "Edge起動中 (CDP port: $Port)..."
}
Start-Process "msedge.exe" $edgeArgs

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

Write-Warning "CDPポートの起動を確認できませんでした。"
Write-Host ""
if (-not $UserDataDir) {
    # Edge 136以降は、普段の保存場所のままだとデバッグポートの指定が黙って無視される。
    # 別の保存場所を指定したときだけ有効になる（公式の回避策は他に無い）。
    Write-Host "Edge 136以降は、普段の保存場所のままだとデバッグポートの指定が無視されます。"
    Write-Host "保存場所を指定して立て直してください:"
    Write-Host "  .\scripts\start_edge.ps1 -Port $Port -UserDataDir C:\edgework"
    Write-Host ""
}
Write-Host "手動で確認するには:"
Write-Host "  Invoke-RestMethod http://localhost:$Port/json"
