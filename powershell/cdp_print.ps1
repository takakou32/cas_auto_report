<#
.SYNOPSIS
    CAS Auto Report (PowerShell版) - CDP経由でEdgeにアタッチし、帳票出力操作を宛名番号ごとに繰り返す

.DESCRIPTION
    外部依存なし（.NET の ClientWebSocket で Chrome DevTools Protocol を直接操作）。
    帳票ごとに1回だけ操作を記録し、あとは対象者リスト（宛名番号）の分だけ再生する。
    出力された帳票の回収は別ツールの担当。本ツールは「出力操作を繰り返す」ところまで。

    前提:
      1. Edgeを全プロセス終了
      2. scripts/start_edge.ps1 でデバッグポート付きEdgeを起動
      3. exeからアプリを起動（Edgeにタブが追加される）
      4. このスクリプトを実行: .\powershell\cdp_print.ps1

.PARAMETER Config
    設定ファイル(JSON)のパス（デフォルト: config/config.json）

.PARAMETER List
    接続中のEdgeのタブ一覧を表示して終了する

.PARAMETER CdpUrl
    CDPのURL（指定時は設定ファイルの cdp_url を上書き）

.PARAMETER Record
    操作記録モード。-Job と -Atena が必要

.PARAMETER Job
    帳票名。記録時は保存先 jobs/<Job>.json、実行時は処理する帳票（省略時は jobs/ の全件）

.PARAMETER Atena
    記録時に使う宛名番号。記録した手順の中のこの番号を {{atena}} に置き換える

.PARAMETER DryRun
    出力実行ボタン（final）とそれ以降の手順を行わず、手前まで再生する。台帳には dry と記録する

.PARAMETER Limit
    処理する件数の上限（全帳票の合計）。0なら制限なし

.PARAMETER Evidence
    証跡スクリーンショット（指定時は設定ファイルの evidence を上書き）: none | before | after | both

.EXAMPLE
    .\powershell\cdp_print.ps1 -List
.EXAMPLE
    .\powershell\cdp_print.ps1 -Record -Job R001_住民票 -Atena 1234567890
.EXAMPLE
    .\powershell\cdp_print.ps1 -Job R001_住民票 -DryRun -Limit 1
#>

[CmdletBinding()]
param(
    [Alias("c")]
    [string]$Config = "config/config.json",
    [switch]$List,
    [string]$CdpUrl,
    [switch]$Record,
    [string]$Job,
    [string]$Atena,
    [switch]$DryRun,
    [int]$Limit = 0,
    [ValidateSet("", "none", "before", "after", "both")]
    [string]$Evidence = ""
)

$ErrorActionPreference = "Stop"
$script:CdpId = 0

# 実行結果の記録先（リポジトリ直下・追記のみ）
$script:LedgerPath = "ledger.csv"

# .NET の相対パス基準(WriteAllBytes等)を PowerShell のカレントに合わせる。
# これをしないと、別ディレクトリから起動した際に出力先がずれる。
[System.IO.Directory]::SetCurrentDirectory((Get-Location).Path)

# ---------------------------------------------------------------------------
# 設定読み込み
# ---------------------------------------------------------------------------
function Get-CapConfig {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        Write-Host "設定ファイルが見つかりません: $Path"
        Write-Host "config/config.sample.json をコピーして config/config.json を作成してください"
        exit 1
    }
    return Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

# ---------------------------------------------------------------------------
# CDP HTTPエンドポイント
# ---------------------------------------------------------------------------
function Get-CdpTabs {
    param([string]$BaseUrl)
    try {
        $tabs = Invoke-RestMethod -Uri "$BaseUrl/json" -Method Get
    } catch {
        Write-Host "接続エラー: $BaseUrl"
        Write-Host "  Edgeがデバッグポート付きで起動しているか確認してください"
        exit 1
    }
    # ページタイプのタブのみ対象（service_worker等を除外）
    return @($tabs | Where-Object { $_.type -eq "page" })
}

# ---------------------------------------------------------------------------
# CDP WebSocket 通信
# ---------------------------------------------------------------------------
function Connect-CdpSocket {
    param([string]$WsUrl)
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = [Uri]$WsUrl
    $ws.ConnectAsync($uri, [System.Threading.CancellationToken]::None).GetAwaiter().GetResult() | Out-Null
    return $ws
}

function Send-CdpRaw {
    param($Ws, [string]$Json)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Json)
    $segment = [System.ArraySegment[byte]]::new($bytes)
    $Ws.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true,
        [System.Threading.CancellationToken]::None).GetAwaiter().GetResult() | Out-Null
}

function Receive-CdpRaw {
    param($Ws)
    $buffer = New-Object byte[] 131072
    $sb = New-Object System.Text.StringBuilder
    do {
        $segment = [System.ArraySegment[byte]]::new($buffer)
        $result = $Ws.ReceiveAsync($segment, [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
        [void]$sb.Append([System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count))
    } while (-not $result.EndOfMessage)
    return $sb.ToString()
}

# CDPコマンドを送信し、対応するid応答を待つ（途中のイベント通知は読み飛ばす）
function Invoke-CdpCommand {
    param($Ws, [string]$Method, $Params = $null)
    $script:CdpId++
    $id = $script:CdpId
    $msg = @{ id = $id; method = $Method }
    if ($null -ne $Params) { $msg.params = $Params }
    $json = $msg | ConvertTo-Json -Depth 20 -Compress
    Send-CdpRaw -Ws $Ws -Json $json

    while ($true) {
        $raw = Receive-CdpRaw -Ws $Ws
        $obj = $raw | ConvertFrom-Json
        if ($obj.id -eq $id) {
            if ($obj.error) { throw "CDPエラー [$Method]: $($obj.error.message)" }
            return $obj.result
        }
        # idが一致しないものはイベント通知なので無視
    }
}

# ページコンテキストでJavaScriptを評価する
function Invoke-PageScript {
    param($Ws, [string]$Expression, [bool]$AwaitPromise = $false)
    $result = Invoke-CdpCommand -Ws $Ws -Method "Runtime.evaluate" -Params @{
        expression   = $Expression
        awaitPromise = $AwaitPromise
        returnByValue = $true
    }
    if ($result.exceptionDetails) {
        $desc = $result.exceptionDetails.exception.description
        if (-not $desc) { $desc = $result.exceptionDetails.text }
        throw "JS実行エラー: $desc"
    }
    return $result.result.value
}

# ページ遷移中はJS実行コンテキストが破棄され Runtime.evaluate が失敗する。
# 遷移由来の一時的エラーは少し待って新しいコンテキストでリトライする。
function Invoke-PageScriptSafe {
    param($Ws, [string]$Expression, [bool]$AwaitPromise = $false, [int]$Retries = 40)
    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            return Invoke-PageScript -Ws $Ws -Expression $Expression -AwaitPromise $AwaitPromise
        } catch {
            $msg = "$_"
            if ($msg -match "Execution context was destroyed" -or
                $msg -match "Cannot find context" -or
                $msg -match "Inspected target navigated or closed" -or
                $msg -match "uniqueContextId") {
                Start-Sleep -Milliseconds 200
                continue
            }
            throw
        }
    }
    throw "ページ評価がナビゲーションにより安定しませんでした"
}

# 文字列を安全なJSリテラルに変換（JSON文字列はJS文字列としても妥当）
function ConvertTo-JsLiteral {
    param([string]$Value)
    return ($Value | ConvertTo-Json -Compress)
}

# ---------------------------------------------------------------------------
# 待機処理
#   1) ready_selector が指定されていれば、その要素が表示されるまで待つ
#   2) readyState=complete かつ DOMが stable_ms の間変化しなくなるまで待つ（最大 timeout）
#   3) 仕上げに settle_ms 待つ
# これにより「readyStateはcompleteだが中身はまだローディング中」の画面を操作するのを防ぐ。
# stable_ms / load_timeout_ms / ready_selector は $script: 変数で上書き可（Invoke-PrintRunで設定）。
# ---------------------------------------------------------------------------
function Wait-PageReady {
    param($Ws, [int]$SettleMs = 800, [int]$TimeoutMs = 0)

    $timeout  = if ($TimeoutMs -gt 0) { $TimeoutMs }
                elseif ($script:LoadTimeoutMs) { [int]$script:LoadTimeoutMs } else { 30000 }
    $stable   = if ($null -ne $script:StableMs) { [int]$script:StableMs } else { 1000 }
    $selector = if ($script:ReadySelector) { [string]$script:ReadySelector } else { "" }

    # 1) 目印要素の出現待ち（任意）
    if ($selector) {
        $sel = ConvertTo-JsLiteral $selector
        $expr = @"
new Promise((resolve, reject) => {
  const deadline = Date.now() + $timeout;
  (function check() {
    const el = document.querySelector($sel);
    if (el && el.offsetParent !== null) return resolve(true);
    if (Date.now() > deadline) return reject(new Error('ready_selector timeout: ' + $sel));
    setTimeout(check, 100);
  })();
})
"@
        Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true | Out-Null
    }

    # 2) readyState完了 ＋ DOM安定待ち（コンテンツの挿入が止まるまで）
    $expr = @"
new Promise((resolve) => {
  const idle = $stable, deadline = Date.now() + $timeout;
  let last = Date.now();
  let obs = null;
  try {
    obs = new MutationObserver(() => { last = Date.now(); });
    obs.observe(document.documentElement, { childList: true, subtree: true });
  } catch (e) {}
  (function check() {
    const now = Date.now();
    if (document.readyState === 'complete' && (now - last) >= idle) {
      if (obs) obs.disconnect();
      return resolve('idle');
    }
    if (now > deadline) {
      if (obs) obs.disconnect();
      return resolve('timeout');
    }
    setTimeout(check, 100);
  })();
})
"@
    Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true | Out-Null

    if ($SettleMs -gt 0) { Start-Sleep -Milliseconds $SettleMs }
}

# ---------------------------------------------------------------------------
# アクション実行
#   cas_cap と違い、対象が見つからない・待機タイムアウトは「スキップ」ではなく例外にする。
#   ズレた画面のまま進んで、誤った対象で出力ボタンを押す事故を防ぐため。
# ---------------------------------------------------------------------------
function Invoke-CapAction {
    param($Ws, $Action, [int]$SettleMs)

    switch ($Action.type) {
        "click" {
            # ハイブリッド特定：セレクタで当てた要素を「記録時のボタン名(text/aria-label)」で検証する。
            # 権限差などでDOMの順番が変わり、位置セレクタが“別要素”に当たった場合はラベルで探し直す。
            $sel = ConvertTo-JsLiteral $Action.selector
            $txt = ConvertTo-JsLiteral ([string]$Action.text)
            $to  = $script:ActionTimeoutMs
            $expr = @"
new Promise((resolve) => {
  const sel = $sel, text = $txt, deadline = Date.now() + $to;
  function visible(e){ return e && (e.offsetParent !== null || (e.getClientRects && e.getClientRects().length > 0)); }
  function txtOf(e){
    var s=(e.innerText||e.textContent||'').trim();
    if(!s){ try { s=((e.getAttribute('aria-label')||e.getAttribute('title'))||'').trim(); } catch(_){} }
    return s;
  }
  function matches(a,b){
    if(!a||!b) return false;
    if(a===b) return true;
    return (b.length>=2 && a.indexOf(b)>=0) || (a.length>=2 && b.indexOf(a)>=0);
  }
  function byText(){
    if (!text) return null;
    var list = Array.prototype.slice.call(document.querySelectorAll('a,button,[role=button],[role=tab],[role=menuitem],[role=link],[role=option],li,[tabindex],[onclick]')).filter(visible);
    return list.find(function(e){ return txtOf(e) === text; })
        || list.find(function(e){ return matches(txtOf(e), text); });
  }
  // click() は HTML の要素にしかない。SVG(アイコン等)には無いので、そのまま呼ぶと
  // 「e.click is not a function」で落ちる。その場合はマウス操作を作って投げる。
  // 投げ先は要素そのもの。イベントは祖先へ伝わるので、ボタン側の処理も動く。
  function rawClick(e){
    var r=e.getBoundingClientRect();
    var x=r.left+r.width/2, y=r.top+r.height/2;
    // detail(押した回数)・pointerType(何で押したか)まで入れる。
    // これを見て動きを決めるページがあり、無いと「エラーは出ないが何も起きない」になる。
    var m={bubbles:true, cancelable:true, composed:true, view:window,
           clientX:x, clientY:y, screenX:x, screenY:y,
           button:0, buttons:1, detail:1};
    var p={}; for(var k in m) p[k]=m[k];
    p.pointerId=1; p.pointerType='mouse'; p.isPrimary=true; p.width=1; p.height=1; p.pressure=0.5;
    try { e.dispatchEvent(new PointerEvent('pointerdown', p)); } catch(_){}
    e.dispatchEvent(new MouseEvent('mousedown', m));
    // click() と違い、投げるだけではフォーカスが移らないので明示的に当てる
    try { if(typeof e.focus==='function') e.focus({preventScroll:true}); } catch(_){}
    var pu={}; for(var k2 in p) pu[k2]=p[k2];
    pu.buttons=0; pu.pressure=0;
    try { e.dispatchEvent(new PointerEvent('pointerup', pu)); } catch(_){}
    var mu={}; for(var k3 in m) mu[k3]=m[k3];
    mu.buttons=0;
    e.dispatchEvent(new MouseEvent('mouseup', mu));
    e.dispatchEvent(new MouseEvent('click', mu));
  }
  function go(e,how){
    e.scrollIntoView({block:'center'});
    // 押す相手は記録した要素そのもの。祖先に押し替えると、
    // 押された場所を見て動きを変えるページで結果が変わるため。
    if(typeof e.click==='function'){ e.click(); } else { rawClick(e); }
    return resolve(how);
  }
  (function check(){
    var el = null; try { el = document.querySelector(sel); } catch(e){}
    // 1) セレクタが当たり、かつ(テキスト未記録 or ラベル一致) → それをクリック
    if (visible(el) && (!text || matches(txtOf(el), text))) return go(el, 'clicked');
    // 2) ラベル一致の要素を探す（順番が変わっても“ボタン名”で当てる）
    var c = byText();
    if (c) return go(c, 'text');
    // 3) テキスト情報が無い時のみ、位置一致のセレクタ要素をクリック（アイコン等）
    if (visible(el) && !text) return go(el, 'clicked-notext');
    // 4) テキストはあるが一致要素が無い → まだ描画中かもしれないので待つ
    if (Date.now() > deadline) return resolve('notfound');
    setTimeout(check, 150);
  })();
})
"@
            $st = Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true
            switch ($st) {
                'notfound'       { throw "クリック対象が見つかりません: $($Action.selector) (ボタン名: $($Action.text))" }
                'text'           { Write-Host "  (ボタン名一致でクリック: $($Action.text))" }
                'clicked-notext' { Write-Host "  (位置一致でクリック: $($Action.selector))" }
            }
        }
        "fill" {
            $sel = ConvertTo-JsLiteral $Action.selector
            $val = ConvertTo-JsLiteral $Action.value
            $to  = $script:ActionTimeoutMs
            $expr = @"
new Promise((resolve) => {
  const sel = $sel, val = $val, deadline = Date.now() + $to;
  (function check(){
    var el = null; try { el = document.querySelector(sel); } catch(e){}
    if (el) {
      el.focus(); el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return resolve('ok');
    }
    if (Date.now() > deadline) return resolve('notfound');
    setTimeout(check, 150);
  })();
})
"@
            $st = Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true
            if ($st -eq 'notfound') { throw "入力対象が見つかりません: $($Action.selector)" }
        }
        "setcheck" {
            # チェックボックス/ラジオを「指定の状態」にする。
            # 既にその状態ならクリックしない（クリック＝トグルなので、押すと逆になる）。
            $sel  = ConvertTo-JsLiteral $Action.selector
            $lab  = ConvertTo-JsLiteral ([string]$Action.label)
            $want = if ($Action.checked) { "true" } else { "false" }
            $to   = $script:ActionTimeoutMs
            $expr = @"
new Promise((resolve) => {
  const sel = $sel, label = $lab, want = $want, deadline = Date.now() + $to;
  function visible(e){ return e && (e.offsetParent !== null || (e.getClientRects && e.getClientRects().length > 0)); }
  // ラベル文字列から入力欄を探す（セレクタが当たらないときの保険）
  function byLabel(){
    if (!label) return null;
    var boxes = Array.prototype.slice.call(document.querySelectorAll('input[type=checkbox],input[type=radio]')).filter(visible);
    for (var i=0;i<boxes.length;i++){
      var b=boxes[i], t='';
      var lb = b.id ? document.querySelector('label[for="' + b.id + '"]') : null;
      if (!lb && b.closest) lb = b.closest('label');
      if (lb) t = (lb.innerText||lb.textContent||'').trim();
      if (!t) { try { t = ((b.getAttribute('aria-label')||b.getAttribute('title'))||'').trim(); } catch(_){} }
      if (t === label) return b;
    }
    return null;
  }
  function settle(el){
    // フレームワークが再描画で状態を戻すことがあるので、少しの間だけ確認を続ける
    var limit = Date.now() + 1000;
    (function verify(){
      if (!!el.checked === want) return resolve('set');
      if (Date.now() > limit) return resolve('failed');
      setTimeout(verify, 100);
    })();
  }
  (function check(){
    var el = null; try { el = document.querySelector(sel); } catch(e){}
    if (!visible(el)) el = byLabel();
    if (el) {
      if (!!el.checked === want) return resolve('same');
      el.scrollIntoView({block:'center'});
      el.click();
      return settle(el);
    }
    if (Date.now() > deadline) return resolve('notfound');
    setTimeout(check, 150);
  })();
})
"@
            $st = Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true
            switch ($st) {
                'notfound' { throw "チェック対象が見つかりません: $($Action.selector) (項目名: $($Action.label))" }
                'failed'   { throw "チェック状態を変更できませんでした: $($Action.selector) (項目名: $($Action.label))" }
                'set'      { Write-Host "  (チェック変更: $($Action.label) -> $want)" }
            }
        }
        "wait" {
            $timeout = if ($Action.timeout) { [int]$Action.timeout } else { 5000 }
            if ($Action.selector) {
                $sel = ConvertTo-JsLiteral $Action.selector
                $expr = @"
new Promise((resolve, reject) => {
  const deadline = Date.now() + $timeout;
  (function check() {
    if (document.querySelector($sel)) return resolve(true);
    if (Date.now() > deadline) return reject(new Error('wait timeout: ' + $sel));
    setTimeout(check, 100);
  })();
})
"@
                Invoke-PageScript -Ws $Ws -Expression $expr -AwaitPromise $true | Out-Null
            } else {
                Start-Sleep -Milliseconds $timeout
            }
        }
        "goto" {
            if ($script:SpaMode) {
                # SPA(認証あり)向け: リロードせず history.pushState でルートだけ変更し、
                # ロード済みアプリの認証状態を保ったまま画面遷移する（同一オリジンのみ）。
                $u = ConvertTo-JsLiteral $Action.url
                $expr = @"
(function(u){
  try {
    var t = new URL(u, location.href);
    if (t.origin !== location.origin) { location.href = u; return 'hard'; }
    history.pushState({}, '', t.pathname + t.search + t.hash);
    window.dispatchEvent(new PopStateEvent('popstate', { state: history.state }));
    window.dispatchEvent(new Event('hashchange'));
    return 'soft';
  } catch (e) { location.href = u; return 'err'; }
})($u)
"@
                Invoke-PageScriptSafe -Ws $Ws -Expression $expr | Out-Null
                Wait-PageReady -Ws $Ws -SettleMs $SettleMs
            } else {
                Invoke-CdpCommand -Ws $Ws -Method "Page.navigate" -Params @{ url = $Action.url } | Out-Null
                Wait-PageReady -Ws $Ws -SettleMs $SettleMs
            }
        }
        "expect_url" {
            # 直前のクリックで狙った画面へ移ったかを確認する（移動はしない）。
            # 記録時と同じURLになるまで待ち、ならなければその件を失敗にする。
            # クエリ文字列は実行のたびに変わりやすいので、オリジン＋パス＋ハッシュで比べる。
            $u  = ConvertTo-JsLiteral $Action.url
            $to = if ($script:LoadTimeoutMs) { [int]$script:LoadTimeoutMs } else { 30000 }
            $expr = @"
new Promise((resolve) => {
  const want = $u, deadline = Date.now() + $to;
  function norm(u){
    try { var t = new URL(u, location.href); return t.origin + t.pathname + t.hash; }
    catch (e) { return u; }
  }
  const target = norm(want);
  (function check(){
    if (norm(location.href) === target) return resolve('ok');
    if (Date.now() > deadline) return resolve('mismatch:' + location.href);
    setTimeout(check, 100);
  })();
})
"@
            $st = "" + (Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true)
            if ($st -like "mismatch:*") {
                throw "画面遷移が記録と違います: 期待 $($Action.url) / 実際 $($st.Substring(9))"
            }
            Wait-PageReady -Ws $Ws -SettleMs $SettleMs
        }
        "select" {
            $sel = ConvertTo-JsLiteral $Action.selector
            $val = ConvertTo-JsLiteral $Action.value
            $to  = $script:ActionTimeoutMs
            $expr = @"
new Promise((resolve) => {
  const sel = $sel, val = $val, deadline = Date.now() + $to;
  (function check(){
    var el = null; try { el = document.querySelector(sel); } catch(e){}
    if (el) {
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return resolve('ok');
    }
    if (Date.now() > deadline) return resolve('notfound');
    setTimeout(check, 150);
  })();
})
"@
            $st = Invoke-PageScriptSafe -Ws $Ws -Expression $expr -AwaitPromise $true
            if ($st -eq 'notfound') { throw "選択対象が見つかりません: $($Action.selector)" }
        }
        "keyboard" {
            $key = ConvertTo-JsLiteral $Action.key
            $expr = @"
new Promise((resolve) => {
  const el = document.activeElement || document.body;
  const opt = { key: $key, bubbles: true };
  el.dispatchEvent(new KeyboardEvent('keydown', opt));
  el.dispatchEvent(new KeyboardEvent('keypress', opt));
  el.dispatchEvent(new KeyboardEvent('keyup', opt));
  resolve(true);
})
"@
            Invoke-PageScript -Ws $Ws -Expression $expr -AwaitPromise $true | Out-Null
        }
        default {
            Write-Host "  未知のアクション: $($Action.type)"
        }
    }
}

# ---------------------------------------------------------------------------
# ビューポート（デバイスメトリクス）を上書きしてウィンドウ幅に依存しない描画にする
# ---------------------------------------------------------------------------
function Set-Viewport {
    param($Ws, [int]$Width, [int]$Height, [double]$Scale)
    $params = @{
        width             = $Width
        height            = $Height
        deviceScaleFactor = $Scale
        mobile            = $false
    }
    Invoke-CdpCommand -Ws $Ws -Method "Emulation.setDeviceMetricsOverride" -Params $params | Out-Null
}

# ---------------------------------------------------------------------------
# スクリーンショット（証跡用）
#   証跡は「どの対象者で出力ボタンを押したか」が分かれば足りるので、
#   cas_cap のようなモーダル/iframe展開の全画面対応はせず、表示領域のみを撮る。
# ---------------------------------------------------------------------------
function Save-Screenshot {
    param($Ws, [string]$Path)
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $result = Invoke-CdpCommand -Ws $Ws -Method "Page.captureScreenshot" -Params @{ format = "png" }
    [IO.File]::WriteAllBytes($Path, [Convert]::FromBase64String($result.data))
}

# ---------------------------------------------------------------------------
# タブ一覧表示
# ---------------------------------------------------------------------------
function Show-Tabs {
    param([string]$BaseUrl)
    $tabs = Get-CdpTabs -BaseUrl $BaseUrl
    Write-Host "接続成功。ページ数: $($tabs.Count)"
    for ($i = 0; $i -lt $tabs.Count; $i++) {
        Write-Host "  [$i] $($tabs[$i].title)"
        Write-Host "      $($tabs[$i].url)"
    }
}

# ---------------------------------------------------------------------------
# 実行結果の記録（ledger.csv）
#   追記のみ。再実行時は job×atena の最新行で判定する（ok→スキップ、fail→再試行）。
# ---------------------------------------------------------------------------
function ConvertTo-CsvField {
    param([string]$Value)
    if (-not $Value) { return "" }
    # 改行はCSVの行を壊すので空白に潰す
    $v = $Value -replace "`r`n", " " -replace "`r", " " -replace "`n", " "
    if ($v -match '[",]') { return '"' + $v.Replace('"', '""') + '"' }
    return $v
}

function Get-LedgerState {
    param([string]$Path)
    $state = @{}
    if (-not (Test-Path $Path)) { return $state }
    $rows = @(Import-Csv -Path $Path -Encoding UTF8)
    foreach ($r in $rows) {
        if (-not $r.job) { continue }
        # 同じ組み合わせが複数あれば後の行（＝新しい実行）で上書きする
        $state["$($r.job)`t$($r.atena)"] = $r.status
    }
    return $state
}

function Add-LedgerRow {
    param([string]$Path, [string]$JobName, [string]$AtenaNo, [string]$Status, [string]$ErrorText)
    # Excelでそのまま開けるようUTF-8 BOM付きで書く（BOMは新規作成時のみ付く）
    $enc = New-Object System.Text.UTF8Encoding($true)
    if (-not (Test-Path $Path)) {
        [System.IO.File]::AppendAllText($Path, "job,atena,timestamp,status,error`r`n", $enc)
    }
    $line = "{0},{1},{2},{3},{4}`r`n" -f (ConvertTo-CsvField $JobName),
        (ConvertTo-CsvField $AtenaNo),
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        (ConvertTo-CsvField $Status),
        (ConvertTo-CsvField $ErrorText)
    [System.IO.File]::AppendAllText($Path, $line, $enc)
}

# ---------------------------------------------------------------------------
# 対象者リスト（targets/<Job>.csv）
#   1列目が宛名番号。1行目はヘッダ(atena)として読み飛ばす。2列目以降は無視する。
# ---------------------------------------------------------------------------
function Get-TargetAtenas {
    param([string]$Path)
    $list = New-Object System.Collections.ArrayList
    if (-not (Test-Path $Path)) {
        Write-Warning "対象者リストがありません(この帳票はスキップ): $Path"
        return $list
    }
    $lines = @(Get-Content -Path $Path -Encoding UTF8)
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $cell = ($lines[$i] -split ",")[0]
        if ($null -eq $cell) { continue }
        $cell = $cell.Trim().Trim('"').Trim()
        if ($cell) { [void]$list.Add($cell) }
    }
    return $list
}

# ---------------------------------------------------------------------------
# 手順ファイル（jobs/<Job>.json）
# ---------------------------------------------------------------------------
function Get-JobPaths {
    param([string]$JobName)
    if ($JobName) {
        $p = Join-Path "jobs" "$JobName.json"
        if (-not (Test-Path $p)) {
            Write-Host "手順ファイルが見つかりません: $p"
            Write-Host "  先に記録してください: .\powershell\cdp_print.ps1 -Record -Job $JobName -Atena <宛名番号>"
            exit 1
        }
        return @($p)
    }
    if (-not (Test-Path "jobs")) { return @() }
    return @(Get-ChildItem -Path "jobs" -Filter "*.json" | Sort-Object Name | ForEach-Object { $_.FullName })
}

# {{atena}} を実際の宛名番号に差し込む。
# click の照合テキストは「宛名番号だけ」に絞る（氏名の書式は画面によって揺れるため）。
function Expand-JobAction {
    param($Action, [string]$AtenaNo)
    $new = @{}
    foreach ($p in $Action.PSObject.Properties) { $new[$p.Name] = $p.Value }
    if ($new.ContainsKey("selector") -and $new["selector"]) {
        $new["selector"] = ([string]$new["selector"]).Replace("{{atena}}", $AtenaNo)
    }
    if ($new.ContainsKey("value") -and $null -ne $new["value"]) {
        $new["value"] = ([string]$new["value"]).Replace("{{atena}}", $AtenaNo)
    }
    if ($new.ContainsKey("text") -and $new["text"]) {
        $t = [string]$new["text"]
        if ($t.Contains("{{atena}}")) { $new["text"] = $AtenaNo }
    }
    # goto / expect_url のURL（検索がGETで宛名番号がURLに乗る作りへの備え）
    if ($new.ContainsKey("url") -and $new["url"]) {
        $new["url"] = ([string]$new["url"]).Replace("{{atena}}", $AtenaNo)
    }
    return [PSCustomObject]$new
}

# 1人分の手順を頭から再生する。途中で失敗したら例外を投げて、その件を打ち切る。
function Invoke-JobForAtena {
    param($Ws, $JobDef, [string]$AtenaNo, [string]$JobName, [int]$SettleMs,
          [bool]$DryRunMode, [string]$Evidence)

    $stopped = $false
    foreach ($a in @($JobDef.actions)) {
        $act = Expand-JobAction -Action $a -AtenaNo $AtenaNo
        $isFinal = [bool]$act.final

        if ($isFinal -and ($Evidence -eq "before" -or $Evidence -eq "both")) {
            Save-Screenshot -Ws $Ws -Path (Join-Path (Join-Path "output" $JobName) "${AtenaNo}_before.png")
        }
        if ($isFinal -and $DryRunMode) {
            # DryRun: 出力実行ボタンは押さず、それ以降の手順（確認ダイアログのはい等）も行わない
            Write-Host "  (DryRun: 出力実行ボタン以降は行いません: $($act.text))"
            $stopped = $true
            break
        }

        Invoke-CapAction -Ws $Ws -Action $act -SettleMs $SettleMs
    }

    # 証跡(after)は手順を最後まで流し終えてから撮る
    if (-not $stopped -and ($Evidence -eq "after" -or $Evidence -eq "both")) {
        Wait-PageReady -Ws $Ws -SettleMs $SettleMs
        Save-Screenshot -Ws $Ws -Path (Join-Path (Join-Path "output" $JobName) "${AtenaNo}_after.png")
    }
}

# ---------------------------------------------------------------------------
# 実行（再生）本体
# ---------------------------------------------------------------------------
function Invoke-PrintRun {
    param($Cfg, [string]$JobName, [bool]$DryRunMode, [int]$LimitCount, [string]$EvidenceMode)

    $baseUrl  = if ($Cfg.cdp_url) { $Cfg.cdp_url } else { "http://localhost:9222" }
    $settleMs = if ($null -ne $Cfg.settle_ms) { [int]$Cfg.settle_ms } else { 800 }
    $interval = if ($null -ne $Cfg.interval_ms) { [int]$Cfg.interval_ms } else { 1000 }
    $maxFail  = if ($null -ne $Cfg.max_consecutive_fail) { [int]$Cfg.max_consecutive_fail } else { 5 }
    $evidence = if ($EvidenceMode) { $EvidenceMode }
                elseif ($Cfg.evidence) { [string]$Cfg.evidence } else { "none" }

    # ビューポート設定（未指定ならデスクトップ幅をデフォルトにして縦長化を防ぐ）
    $vpWidth  = if ($Cfg.viewport.width)  { [int]$Cfg.viewport.width }  else { 1280 }
    $vpHeight = if ($Cfg.viewport.height) { [int]$Cfg.viewport.height } else { 800 }
    $vpScale  = if ($Cfg.viewport.device_scale_factor) { [double]$Cfg.viewport.device_scale_factor } else { 1 }

    # 待機設定（描画途中の画面を操作しないため）
    #   stable_ms      : DOMがこの時間変化しなくなったら「描画完了」とみなす
    #   load_timeout_ms: 上記を待つ最大時間
    #   ready_selector : 指定するとこの要素が表示されるまで待つ（最も確実）
    $script:StableMs      = if ($null -ne $Cfg.stable_ms) { [int]$Cfg.stable_ms } else { 1000 }
    $script:LoadTimeoutMs = if ($null -ne $Cfg.load_timeout_ms) { [int]$Cfg.load_timeout_ms } else { 30000 }
    $script:ReadySelector = if ($Cfg.ready_selector) { [string]$Cfg.ready_selector } else { "" }
    # 要素クリック/入力で対象を待つ最大時間（超えたらその件は失敗）
    $script:ActionTimeoutMs = if ($null -ne $Cfg.action_timeout_ms) { [int]$Cfg.action_timeout_ms } else { 5000 }

    $jobPaths = Get-JobPaths -JobName $JobName
    if ($jobPaths.Count -eq 0) {
        Write-Host "処理する手順ファイルがありません（jobs/*.json）"
        Write-Host "  先に記録してください: .\powershell\cdp_print.ps1 -Record -Job <帳票名> -Atena <宛名番号>"
        exit 1
    }

    $ledger = Get-LedgerState -Path $script:LedgerPath
    $processed = 0
    $stopAll = $false

    foreach ($jobPath in $jobPaths) {
        if ($stopAll) { break }

        $name = [System.IO.Path]::GetFileNameWithoutExtension($jobPath)
        $jobDef = Get-Content -Path $jobPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $actions = @($jobDef.actions)
        if ($actions.Count -eq 0) {
            Write-Warning "手順が空です(スキップ): $jobPath"
            continue
        }

        $targets = @(Get-TargetAtenas -Path (Join-Path "targets" "$name.csv"))
        if ($targets.Count -eq 0) { continue }

        # SPA判定は手順ファイル（記録時に検出した値）を優先する
        $script:SpaMode = if ($null -ne $jobDef.spa_mode) { [bool]$jobDef.spa_mode }
                          elseif ($null -ne $Cfg.spa_mode) { [bool]$Cfg.spa_mode } else { $false }
        $keyword = if ($jobDef.target_url_keyword) { [string]$jobDef.target_url_keyword }
                   elseif ($Cfg.target_url_keyword) { [string]$Cfg.target_url_keyword } else { "" }

        # 対象タブを探す
        $tabs = Get-CdpTabs -BaseUrl $baseUrl
        $target = $tabs | Where-Object { $_.url -like "*$keyword*" } | Select-Object -First 1
        if (-not $target) {
            Write-Host "対象タブが見つかりません (keyword: $keyword)"
            Write-Host "タブ一覧:"
            foreach ($t in $tabs) { Write-Host "  $($t.url)" }
            exit 1
        }

        Write-Host ""
        Write-Host "=== 帳票: $name （対象 $($targets.Count) 件）==="
        Write-Host "対象タブ: $($target.title)"

        $ws = Connect-CdpSocket -WsUrl $target.webSocketDebuggerUrl
        try {
            Invoke-CdpCommand -Ws $ws -Method "Page.enable" | Out-Null
            Set-Viewport -Ws $ws -Width $vpWidth -Height $vpHeight -Scale $vpScale

            $consecutiveFail = 0
            foreach ($atenaNo in $targets) {
                if ($LimitCount -gt 0 -and $processed -ge $LimitCount) {
                    Write-Host "上限 $LimitCount 件に達したため終了します。"
                    $stopAll = $true
                    break
                }
                # 台帳に成功記録がある組は飛ばす（中断→再実行で続きから）
                if ($ledger["$name`t$atenaNo"] -eq "ok") {
                    Write-Host "[$name] $atenaNo : スキップ（処理済み）"
                    continue
                }

                Write-Host "[$name] $atenaNo : 処理中..."
                # DryRunは「試しに流しただけ」なので ok とは区別する（本番のスキップ判定はokのみ）
                $status = if ($DryRunMode) { "dry" } else { "ok" }
                $errText = ""
                try {
                    Invoke-JobForAtena -Ws $ws -JobDef $jobDef -AtenaNo $atenaNo -JobName $name `
                        -SettleMs $settleMs -DryRunMode $DryRunMode -Evidence $evidence
                } catch {
                    $status = "fail"
                    $errText = "$_"
                }
                Add-LedgerRow -Path $script:LedgerPath -JobName $name -AtenaNo $atenaNo `
                    -Status $status -ErrorText $errText
                $ledger["$name`t$atenaNo"] = $status
                $processed++

                if ($status -ne "fail") {
                    if ($status -eq "dry") { Write-Host "  完了（DryRun）" } else { Write-Host "  完了" }
                    $consecutiveFail = 0
                } else {
                    Write-Warning "  失敗: $errText"
                    $consecutiveFail++
                    if ($consecutiveFail -ge $maxFail) {
                        # 連続で失敗する＝システム側の異常か手順の陳腐化。この帳票は打ち切る。
                        Write-Warning "連続 $consecutiveFail 件失敗したため、帳票 '$name' を中断します。"
                        break
                    }
                }

                if ($interval -gt 0) { Start-Sleep -Milliseconds $interval }
            }
        } finally {
            try {
                $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "done",
                    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult() | Out-Null
            } catch {}
            $ws.Dispose()
        }
    }

    Write-Host ""
    Write-Host "処理件数: $processed 件（記録: $($script:LedgerPath)）"
    Write-Host "完了"
}

# ---------------------------------------------------------------------------
# 操作記録（レコーディング）
# ---------------------------------------------------------------------------
# ページに監視用JSを注入し、クリック・入力を localStorage に溜める。
# 画面遷移そのものはPS側がURL変化を監視して検出し、goto アクションとして再現する。
# cas_cap との違い: チェックボックス/ラジオはクリックではなく setcheck として記録する。
$script:RecorderJs = @'
(function(){
  // 1ドキュメントにつき1回だけロード回数を加算（SPAソフト遷移では新ドキュメントにならない＝増えない）
  if (!window.__capLoadCounted){ window.__capLoadCounted = true;
    try { sessionStorage.setItem("__capLoad", String((parseInt(sessionStorage.getItem("__capLoad")||"0",10))+1)); } catch(e){}
  }
  function uniq(sel){ try { return document.querySelectorAll(sel).length === 1; } catch(e){ return false; } }
  // 自動生成され毎回変わるID（PrimeVueのpv_id_、ReactのuseId、各UIライブラリ等）は使わない
  function volatileId(id){
    return !id
      || /^pv_id_/.test(id)
      || /^:r[0-9a-z]+:?$/i.test(id)
      || /^(headlessui|radix|mui|el-id|ember|svelte|aria-)/i.test(id)
      || /[0-9]{4,}/.test(id)
      || /_[0-9]+(_|$)/.test(id);
  }
  // ハッシュ的・状態的でない安定したクラスだけ残す
  function stableClasses(el){
    if (!el.classList) return [];
    return Array.prototype.slice.call(el.classList).filter(function(c){
      return c && c.length>1
        && !/[0-9]{3,}/.test(c)
        && !/^(ng-|v-|jsx-|css-|sc-|is-|has-|active|selected|open|show|hover|focus)/.test(c)
        && !/--[0-9a-f]{4,}/.test(c)
        && !/[0-9a-f]{6,}/.test(c);
    });
  }
  function attrSel(el){
    var tag=el.tagName.toLowerCase();
    var attrs=["data-pc-section","data-pc-name","data-testid","data-test","data-cy","name","aria-label","title","role","placeholder"];
    for (var i=0;i<attrs.length;i++){
      var a=attrs[i]; var v=el.getAttribute && el.getAttribute(a);
      if(v){ var s=tag+"["+a+"=\""+(""+v).replace(/"/g,'\\"')+"\"]"; if(uniq(s)) return s; }
    }
    return null;
  }
  function cssPath(el){
    if (!el || el.nodeType !== 1) return "";
    if (el.id && !volatileId(el.id) && uniq("#"+CSS.escape(el.id))) return "#"+CSS.escape(el.id);
    var sa=attrSel(el); if(sa) return sa;
    var parts=[]; var node=el; var depth=0;
    while(node && node.nodeType===1 && node!==document.body && node!==document.documentElement && depth<12){
      if(node.id && !volatileId(node.id)){ parts.unshift("#"+CSS.escape(node.id)); break; }
      var seg=node.tagName.toLowerCase();
      var cls=stableClasses(node);
      if(cls.length){ seg += "."+cls.map(function(c){ return CSS.escape(c); }).join("."); }
      var parent=node.parentNode;
      if(parent){
        var sibs=Array.prototype.slice.call(parent.children).filter(function(c){
          if(c.tagName!==node.tagName) return false;
          if(!cls.length) return true;
          return cls.every(function(k){ return c.classList && c.classList.contains(k); });
        });
        if(sibs.length>1){ seg += ":nth-of-type("+(Array.prototype.indexOf.call(parent.children,node)+1)+")"; }
      }
      parts.unshift(seg);
      var cand=parts.join(" > ");
      try { if(document.querySelectorAll(cand).length===1) return cand; } catch(e){}
      node=parent; depth++;
    }
    return parts.join(" > ");
  }
  function push(ev){ try{ ev.url=location.href; var k="__capRec"; var arr=JSON.parse(localStorage.getItem(k)||"[]"); arr.push(ev); localStorage.setItem(k,JSON.stringify(arr)); }catch(e){} }

  // クリック対象を決める。標準的な要素が無ければ、クリック地点から数階層さかのぼって
  // 「クリック可能そうな祖先(cursor:pointer / role / aria-label / onclick / tabindex)」を探し、
  // それも無ければ アイコン要素(svg/i/icon系class)そのものを対象にする（虫眼鏡等のアイコンボタン対策）。
  function clickTarget(t0){
    if(!t0 || t0===document.body || t0===document.documentElement) return null;
    var t=t0.closest("a,button,[role=button],[role=tab],[role=menuitem],[role=link],[role=option],li,[onclick],[tabindex]");
    if(t) return t;
    var el=t0, hops=0;
    while(el && el!==document.body && el!==document.documentElement && hops<5){
      var cur=""; try{ cur=getComputedStyle(el).cursor; }catch(_){}
      if(cur==="pointer") return el;
      if(el.getAttribute && (el.getAttribute("role")||el.getAttribute("aria-label")||el.getAttribute("onclick")||el.hasAttribute("tabindex"))) return el;
      el=el.parentElement; hops++;
    }
    var tag=(t0.tagName||"").toLowerCase();
    var cls=(t0.getAttribute && (t0.getAttribute("class")||"")) || "";
    if(tag==="svg"||tag==="i"||tag==="use"||/icon|search|magnif|glass/i.test(cls)){
      return t0.closest("button,a,[role],[onclick],[tabindex],span,i,svg") || t0;
    }
    return null;
  }
  // チェックボックス/ラジオ本体を探す（ラベルをクリックした場合も本体に辿り着く）
  function checkTarget(t0){
    if(!t0) return null;
    function isBox(e){ if(!e||(e.tagName||"").toLowerCase()!=="input") return null; var ty=(e.type||"").toLowerCase(); return (ty==="checkbox"||ty==="radio")?e:null; }
    var b=isBox(t0); if(b) return b;
    var lb=t0.closest ? t0.closest("label") : null;
    if(lb){
      var c=lb.control || (lb.htmlFor ? document.getElementById(lb.htmlFor) : null) || lb.querySelector("input[type=checkbox],input[type=radio]");
      return isBox(c);
    }
    return null;
  }
  // チェック項目の見出し（label要素 → aria-label/title）
  function checkLabel(el){
    var lb = el.id ? document.querySelector('label[for="'+el.id+'"]') : null;
    if(!lb && el.closest) lb = el.closest("label");
    if(lb){ var s=(lb.innerText||lb.textContent||"").trim(); if(s) return s.slice(0,80); }
    var v = el.getAttribute && (el.getAttribute("aria-label")||el.getAttribute("title"));
    return v ? (""+v).trim().slice(0,80) : "";
  }
  // ラベル文字列（テキスト → 自身/祖先の aria-label / title）
  function labelOf(t){
    var s=(t.innerText||t.textContent||"").trim();
    if(s) return s.slice(0,80);
    var el=t, hops=0;
    while(el && hops<3){
      var v = el.getAttribute && (el.getAttribute("aria-label")||el.getAttribute("title"));
      if(v) return (""+v).trim().slice(0,80);
      el=el.parentElement; hops++;
    }
    return "";
  }
  var clickH = function(e){
    // チェックボックス/ラジオは「最終的にどの状態であるべきか」を記録する。
    // クリック直後はまだ checked が切り替わっていないので、次のタスクで読む。
    var box=checkTarget(e.target);
    if(box){
      var bsel=cssPath(box), blab=checkLabel(box);
      setTimeout(function(){ push({type:"setcheck", selector:bsel, label:blab, checked: !!box.checked}); }, 0);
      return;
    }
    var t=clickTarget(e.target);
    if(!t) return;
    var tg=(t.tagName||"").toLowerCase();
    if(tg==="input"||tg==="textarea"){
      var ty=(t.type||"").toLowerCase();
      if(ty!=="submit"&&ty!=="button") return; // テキスト入力はfill、チェック類はsetcheckで扱う
    }
    push({type:"click", selector:cssPath(t), text:labelOf(t)});
  };
  var changeH = function(e){
    var el=e.target; var tag=(el.tagName||"").toLowerCase();
    if(tag==="select"){ push({type:"select", selector:cssPath(el), value:el.value}); }
    else if(el.type==="checkbox"||el.type==="radio"){ /* setcheckで記録済み */ }
    else if(tag==="input"||tag==="textarea"){ push({type:"fill", selector:cssPath(el), value:el.value}); }
  };

  // 古いハンドラがあれば除去して最新を付け直す。
  // これにより「ページを開いたまま録り直し」ても古いcssPath実装が残らない。
  try { if(window.__capClickH)  document.removeEventListener("click",  window.__capClickH,  true); } catch(e){}
  try { if(window.__capChangeH) document.removeEventListener("change", window.__capChangeH, true); } catch(e){}
  window.__capClickH = clickH;
  window.__capChangeH = changeH;
  document.addEventListener("click",  clickH,  true);
  document.addEventListener("change", changeH, true);
})();
'@

# 現在のURLと、溜まった操作イベントをまとめて回収してバッファをクリアするJS
$script:DrainJs = @'
(function(){try{var k="__capRec";var arr=JSON.parse(localStorage.getItem(k)||"[]");localStorage.setItem(k,"[]");var ld=0;try{ld=parseInt(sessionStorage.getItem("__capLoad")||"0",10)}catch(e){}return JSON.stringify({url:location.href,load:ld,events:arr});}catch(e){return JSON.stringify({url:"",load:0,events:[]});}})()
'@

# 記録した手順を1本の actions 配列として確定する
function Add-RecAction {
    param($Action, [bool]$Verbose)
    if (-not $Action) { return }
    # 同じ対象への連続した入力/チェックは最後の状態だけ残す
    $prev = if ($script:RecActions.Count -gt 0) { $script:RecActions[$script:RecActions.Count - 1] } else { $null }
    if ($prev -and $prev.type -eq $Action.type -and $prev.selector -eq $Action.selector -and
        ($Action.type -eq "fill" -or $Action.type -eq "select" -or $Action.type -eq "setcheck")) {
        $script:RecActions[$script:RecActions.Count - 1] = $Action
    } else {
        [void]$script:RecActions.Add($Action)
    }
    if ($Verbose) {
        $desc = switch ($Action.type) {
            "click"      { "click $($Action.selector) ($($Action.text))" }
            "goto"       { "goto $($Action.url)" }
            "expect_url" { "expect_url $($Action.url)" }
            "setcheck"   { "setcheck $($Action.label) = $($Action.checked)" }
            "fill"       { "fill $($Action.selector) = $($Action.value)" }
            default      { "$($Action.type) $($Action.selector)" }
        }
        Write-Host "  手順 $($script:RecActions.Count): $desc"
    }
}

# 保留中のクリックを「URLを変えないページ内クリック」として確定する
function Complete-PendingClick {
    param([bool]$Verbose)
    if ($script:RecPendingClick) {
        Add-RecAction -Action $script:RecPendingClick -Verbose $Verbose
        $script:RecPendingClick = $null
        $script:RecClickArmed = 0
    }
}

# ドレイン結果(JSON文字列)を解釈し、クリック・入力・URL遷移を発生順に手順化する
function Add-RecordSample {
    param([string]$Json, [bool]$Verbose)
    if (-not $Json) { return }
    $obj = $null
    try { $obj = $Json | ConvertFrom-Json } catch { return }
    if (-not $obj) { return }

    # 1) イベント(クリック/入力/チェック)を発生順に処理
    foreach ($e in @($obj.events)) {
        if ($e.type -eq "click") {
            # クリックはすぐ確定せず“保留”する。数ポーリング以内にURLが変われば
            #   → 遷移リンク等 → 宛先URLへの goto として確定（URLで確実に再現できる）
            # URLが変わらなければ
            #   → ページ内クリック(検索ボタン/一覧の行等) → click として確定（ボタン名で照合）
            Complete-PendingClick -Verbose $Verbose
            $clickAct = [ordered]@{ type = "click"; selector = $e.selector }
            if ($e.text) { $clickAct.text = [string]$e.text }
            $script:RecPendingClick = $clickAct
            $script:RecClickArmed = 3
        }
        elseif ($e.type -eq "setcheck") {
            Complete-PendingClick -Verbose $Verbose
            $act = [ordered]@{ type = "setcheck"; selector = $e.selector }
            if ($e.label) { $act.label = [string]$e.label }
            $act.checked = [bool]$e.checked
            Add-RecAction -Action $act -Verbose $Verbose
        }
        elseif ($e.type -eq "fill" -or $e.type -eq "select") {
            Complete-PendingClick -Verbose $Verbose
            $act = [ordered]@{ type = $e.type; selector = $e.selector; value = $e.value }
            Add-RecAction -Action $act -Verbose $Verbose
        }
    }

    # 2) URL変化の処理（イベント処理の後）
    $url = $obj.url
    if ($url -and $url -ne $script:RecLastUrl) {
        # SPA判定: 同一オリジンでURLが変わったのにロード回数が増えていない＝ソフト遷移
        if ($script:RecLastUrl) {
            $sameOrigin = $false
            try { $sameOrigin = (([Uri]$url).GetLeftPart([System.UriPartial]::Authority) -eq ([Uri]$script:RecLastUrl).GetLeftPart([System.UriPartial]::Authority)) } catch {}
            if ($sameOrigin -and [int]$obj.load -le $script:RecLastLoad) { $script:RecSawSpa = $true }
        }
        if ($script:RecPendingClick) {
            # 保留クリックが起こした遷移 → クリックはそのまま残し、
            # その直後に「このURLになったことの確認」を入れる（押すこと自体に意味があるため）
            Complete-PendingClick -Verbose $Verbose
            Add-RecAction -Action ([ordered]@{ type = "expect_url"; url = $url }) -Verbose $Verbose
        } else {
            # 直前にクリックが無いURL変化（記録開始時の画面・アドレスバー直打ち等）だけ goto にする
            Add-RecAction -Action ([ordered]@{ type = "goto"; url = $url }) -Verbose $Verbose
        }
        $script:RecLastUrl = $url
    }
    if ($null -ne $obj.load) { $script:RecLastLoad = [int]$obj.load }
    # 猶予を1ポーリング分ずつ減衰。0になっても保留クリックが残っていれば
    # 「URLを変えないページ内クリック」として確定する。
    if ($script:RecClickArmed -gt 0) {
        $script:RecClickArmed--
        if ($script:RecClickArmed -eq 0 -and $script:RecPendingClick) {
            Add-RecAction -Action $script:RecPendingClick -Verbose $Verbose
            $script:RecPendingClick = $null
        }
    }
}

# 記録した手順の中の宛名番号を {{atena}} に置き換える。置換した箇所数を返す。
function Convert-AtenaPlaceholder {
    param($Actions, [string]$AtenaNo)
    $count = 0
    foreach ($a in $Actions) {
        if ($a.type -eq "fill" -and [string]$a["value"] -eq $AtenaNo) {
            # 宛名番号を入力した欄
            $a["value"] = "{{atena}}"
            $count++
        }
        if ($a.type -eq "click") {
            # 検索結果の行クリック等でボタン名に番号が写り込むケース（部分文字列を置換）
            if ($a.Contains("text") -and [string]$a["text"] -like "*$AtenaNo*") {
                $a["text"] = ([string]$a["text"]).Replace($AtenaNo, "{{atena}}")
                $count++
            }
            if ([string]$a["selector"] -like "*$AtenaNo*") {
                $a["selector"] = ([string]$a["selector"]).Replace($AtenaNo, "{{atena}}")
                $count++
            }
        }
        # 検索がGETで宛名番号がURLに乗る作りへの備え（goto / expect_url の両方）
        if ($a.Contains("url") -and [string]$a["url"] -like "*$AtenaNo*") {
            $a["url"] = ([string]$a["url"]).Replace($AtenaNo, "{{atena}}")
            $count++
        }
    }
    return $count
}

# 記録した手順を、そのまま再生できる手順ファイルとして保存する（毎回上書き）
function Save-JobFile {
    param($BaseCfg, [string]$JobName, $Actions, [string]$OutPath, [bool]$SpaMode = $false)

    $out = [ordered]@{
        job                = $JobName
        target_url_keyword = $BaseCfg.target_url_keyword
        spa_mode           = ($SpaMode -or [bool]$BaseCfg.spa_mode)
        actions            = @($Actions)
    }

    $dir = Split-Path $OutPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $jsonText = $out | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding($false)))
}

function Start-JobRecording {
    param($Cfg, [string]$JobName, [string]$AtenaNo, [string]$OutPath)

    $baseUrl = if ($Cfg.cdp_url) { $Cfg.cdp_url } else { "http://localhost:9222" }
    $keyword = if ($Cfg.target_url_keyword) { $Cfg.target_url_keyword } else { "" }

    $tabs = Get-CdpTabs -BaseUrl $baseUrl
    $target = $tabs | Where-Object { $_.url -like "*$keyword*" } | Select-Object -First 1
    if (-not $target) {
        Write-Host "対象タブが見つかりません (keyword: $keyword)"
        Write-Host "タブ一覧:"
        foreach ($t in $tabs) { Write-Host "  $($t.url)" }
        exit 1
    }

    Write-Host "Edge接続成功"
    Write-Host "記録対象タブ: $($target.title)"

    $ws = Connect-CdpSocket -WsUrl $target.webSocketDebuggerUrl
    $script:RecActions      = New-Object System.Collections.ArrayList  # 確定した手順（順序どおり）
    $script:RecLastUrl      = $null
    $script:RecPendingClick = $null  # 後決め用に保留中のクリック（goto か click かは後で確定）
    $script:RecClickArmed   = 0      # クリック起因の遅延遷移を紐付ける残り猶予ポーリング数
    $script:RecLastLoad     = 0      # 直近のドキュメントロード回数
    $script:RecSawSpa       = $false # SPAソフト遷移を1度でも検出したか
    try {
        Invoke-CdpCommand -Ws $ws -Method "Page.enable" | Out-Null
        Invoke-CdpCommand -Ws $ws -Method "Runtime.enable" | Out-Null

        # 以降に開く（遷移後の）ページにも自動で注入されるよう登録
        Invoke-CdpCommand -Ws $ws -Method "Page.addScriptToEvaluateOnNewDocument" `
            -Params @{ source = $script:RecorderJs } | Out-Null
        # 現在表示中のページにも即時注入し、バッファを初期化
        Invoke-PageScript -Ws $ws -Expression $script:RecorderJs | Out-Null
        Invoke-PageScript -Ws $ws -Expression "try{localStorage.setItem('__capRec','[]')}catch(e){}" | Out-Null

        Write-Host ""
        Write-Host "=== 操作記録を開始しました（帳票: $JobName / 宛名番号: $AtenaNo）==="
        Write-Host "メニューの基準画面から、出力完了まで一通り操作してください。"
        Write-Host "クリック・画面遷移・入力・チェックの操作を順に記録します。"
        Write-Host "記録を終了するには、このウィンドウで Enter キーを押してください。"
        Write-Host ""

        while ($true) {
            $json = $null
            try { $json = Invoke-PageScript -Ws $ws -Expression $script:DrainJs } catch { $json = $null }
            Add-RecordSample -Json $json -Verbose $true

            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq "Enter") { break }
            }
            Start-Sleep -Milliseconds 400
        }

        # 終了直前の状態を回収し、保留中の操作があれば最後の手順として確定
        $json = $null
        try { $json = Invoke-PageScript -Ws $ws -Expression $script:DrainJs } catch { $json = $null }
        Add-RecordSample -Json $json -Verbose $false
        Complete-PendingClick -Verbose $false
    } finally {
        try {
            $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "done",
                [System.Threading.CancellationToken]::None).GetAwaiter().GetResult() | Out-Null
        } catch {}
        $ws.Dispose()
    }

    $acts = $script:RecActions

    Write-Host ""
    Write-Host "記録した手順数: $($acts.Count)"
    if ($script:RecSawSpa) {
        Write-Host "SPA(クライアントサイド遷移)を検出 → spa_mode=true で保存します（再生時はリロードせず遷移）。"
    }
    if ($acts.Count -eq 0) {
        Write-Host "手順が記録されませんでした。保存はスキップします。"
        return
    }

    # クリックが1つも無い手順は「押す操作」が抜けている＝帳票が出ない。保存せずに中止する。
    $clickIdx = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $acts.Count; $i++) {
        if ($acts[$i].type -eq "click") { [void]$clickIdx.Add($i) }
    }
    if ($clickIdx.Count -eq 0) {
        Write-Host ""
        Write-Warning "クリックが1つも記録されませんでした。手順として成立しないため保存を中止します。"
        Write-Warning "  メニューの基準画面から、出力完了まで通しで操作してから Enter を押してください。"
        return
    }

    # 宛名番号を差し替え可能にする
    $replaced = Convert-AtenaPlaceholder -Actions $acts -AtenaNo $AtenaNo
    if ($replaced -eq 0) {
        Write-Warning "宛名番号 '$AtenaNo' が手順のどこにも見つかりませんでした。"
        Write-Warning "  対象者ごとに差し替わらない手順になっています。記録し直してください。"
    } else {
        Write-Host "宛名番号を {{atena}} に置換: $replaced 箇所"
    }

    # 「出力実行」ボタンがどれかを選んでもらう。
    # 出力ボタンの後に確認ダイアログの「はい」がある画面だと、最後のクリック＝出力ボタンとは限らない。
    $cands = @($clickIdx | Select-Object -Last 5)
    $defNo = $cands.Count
    Write-Host ""
    Write-Host "どれが「出力実行」ボタンですか？（記録した最後の $($cands.Count) クリック）"
    for ($n = 0; $n -lt $cands.Count; $n++) {
        $ca = $acts[$cands[$n]]
        Write-Host ("  {0}: {1}  [{2}]" -f ($n + 1), $ca["text"], $ca["selector"])
    }
    Write-Host "選んだクリック以降の手順は、-DryRun のとき実行しません。"
    $ans = Read-Host "番号を入力してください（Enterのみなら $defNo = 最後のクリック）"
    $sel = $defNo
    if ($ans) {
        if ($ans -match '^\d+$' -and [int]$ans -ge 1 -and [int]$ans -le $cands.Count) {
            $sel = [int]$ans
        } else {
            Write-Warning "番号が正しくありません。最後のクリック($defNo)を出力実行ボタンとして扱います。"
        }
    }
    $fa = $acts[$cands[$sel - 1]]
    $fa["final"] = $true
    Write-Host "  出力実行ボタン: $($fa["text"]) [$($fa["selector"])] （final=true）"

    Save-JobFile -BaseCfg $Cfg -JobName $JobName -Actions $acts -OutPath $OutPath -SpaMode $script:RecSawSpa
    Write-Host ""
    Write-Host "保存先: $OutPath"
    Write-Host "対象者リストを用意してください: targets/$JobName.csv （1列目に宛名番号、1行目はヘッダ）"
    Write-Host "試し実行: .\powershell\cdp_print.ps1 -Job `"$JobName`" -DryRun -Limit 1"
}

# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
if ($List) {
    $baseUrl = if ($CdpUrl) { $CdpUrl } else { "http://localhost:9222" }
    Show-Tabs -BaseUrl $baseUrl
} elseif ($Record) {
    if (-not $Job)   { Write-Host "-Record には -Job（帳票名）を指定してください"; exit 1 }
    if (-not $Atena) { Write-Host "-Record には -Atena（記録に使う宛名番号）を指定してください"; exit 1 }
    $cfg = Get-CapConfig -Path $Config
    if ($CdpUrl) { $cfg.cdp_url = $CdpUrl }
    Start-JobRecording -Cfg $cfg -JobName $Job -AtenaNo $Atena -OutPath (Join-Path "jobs" "$Job.json")
} else {
    $cfg = Get-CapConfig -Path $Config
    if ($CdpUrl) { $cfg.cdp_url = $CdpUrl }
    Invoke-PrintRun -Cfg $cfg -JobName $Job -DryRunMode ([bool]$DryRun) -LimitCount $Limit -EvidenceMode $Evidence
}
