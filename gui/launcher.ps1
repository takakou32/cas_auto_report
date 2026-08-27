<#
.SYNOPSIS
    CAS Auto Report ランチャー GUI - start_edge / cdp_print を各オプションで起動する

.DESCRIPTION
    WinFormsベースの簡易ランチャー。各ツールを実行モードを選んで起動する。
    実行は新しいPowerShellコンソールで行うため、出力やプロンプト（記録の停止など）を
    そのまま確認・操作できる。

.EXAMPLE
    powershell -STA -ExecutionPolicy Bypass -File gui\launcher.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# リポジトリルート（このスクリプトの1つ上）
$RepoRoot = Split-Path $PSScriptRoot -Parent

# 指定コマンドを新しいPowerShellコンソールで実行する（リポジトリルートで動作）
function Start-InConsole {
    param([string]$Command)

    # 引用符の受け渡し崩れを避けるため、内側・外側とも EncodedCommand 化する。
    # 内側: 本体を子PowerShellで実行（cdp_print.ps1 内の exit / 例外でも終了コードを取得できる）
    $inner = "Set-Location -LiteralPath '$RepoRoot'; $Command"
    $innerEnc = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($inner))

    # 外側(親ウィンドウ): 内側を実行 → 成功なら自動で閉じる / 失敗時のみ一時停止
    $wrapper = "powershell -NoProfile -ExecutionPolicy Bypass -EncodedCommand $innerEnc; " +
               "if (`$LASTEXITCODE -ne 0) { Write-Host ''; " +
               "Write-Host 'エラーで終了しました。内容を確認してください。' -ForegroundColor Red; " +
               "[void](Read-Host 'Enter キーで閉じます') }"
    $wrapperEnc = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($wrapper))

    # -WorkingDirectory で .NET の相対パス基準もリポジトリルートに揃える（-NoExit は付けない）
    Start-Process powershell -WorkingDirectory $RepoRoot -ArgumentList @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $wrapperEnc
    )
}

# ---------------------------------------------------------------------------
# フォーム
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "CAS Auto Report ランチャー"
$form.Size = New-Object System.Drawing.Size(540, 510)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- 1. Edge起動 ----------------------------------------------------------
$gbEdge = New-Object System.Windows.Forms.GroupBox
$gbEdge.Text = "1. Edge起動 (start_edge.ps1)"
$gbEdge.Location = New-Object System.Drawing.Point(12, 12)
$gbEdge.Size = New-Object System.Drawing.Size(500, 60)
$form.Controls.Add($gbEdge)

$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Text = "ポート"
$lblPort.Location = New-Object System.Drawing.Point(15, 26)
$lblPort.Size = New-Object System.Drawing.Size(40, 20)
$gbEdge.Controls.Add($lblPort)

$txtPort = New-Object System.Windows.Forms.TextBox
$txtPort.Text = "9222"
$txtPort.Location = New-Object System.Drawing.Point(60, 23)
$txtPort.Size = New-Object System.Drawing.Size(70, 22)
$gbEdge.Controls.Add($txtPort)

$btnEdge = New-Object System.Windows.Forms.Button
$btnEdge.Text = "Edge起動"
$btnEdge.Location = New-Object System.Drawing.Point(390, 21)
$btnEdge.Size = New-Object System.Drawing.Size(100, 26)
$gbEdge.Controls.Add($btnEdge)

# --- 2. 共通設定 ----------------------------------------------------------
$gbCommon = New-Object System.Windows.Forms.GroupBox
$gbCommon.Text = "2. 共通設定"
$gbCommon.Location = New-Object System.Drawing.Point(12, 80)
$gbCommon.Size = New-Object System.Drawing.Size(500, 90)
$form.Controls.Add($gbCommon)

$lblConfig = New-Object System.Windows.Forms.Label
$lblConfig.Text = "Config"
$lblConfig.Location = New-Object System.Drawing.Point(15, 26)
$lblConfig.Size = New-Object System.Drawing.Size(70, 20)
$gbCommon.Controls.Add($lblConfig)

$txtConfig = New-Object System.Windows.Forms.TextBox
$txtConfig.Text = "config/config.json"
$txtConfig.Location = New-Object System.Drawing.Point(85, 23)
$txtConfig.Size = New-Object System.Drawing.Size(310, 22)
$gbCommon.Controls.Add($txtConfig)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "参照..."
$btnBrowse.Location = New-Object System.Drawing.Point(405, 22)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 24)
$gbCommon.Controls.Add($btnBrowse)

$lblCdp = New-Object System.Windows.Forms.Label
$lblCdp.Text = "CDP URL(任意)"
$lblCdp.Location = New-Object System.Drawing.Point(15, 57)
$lblCdp.Size = New-Object System.Drawing.Size(90, 20)
$gbCommon.Controls.Add($lblCdp)

$txtCdp = New-Object System.Windows.Forms.TextBox
$txtCdp.Text = ""
$txtCdp.Location = New-Object System.Drawing.Point(105, 54)
$txtCdp.Size = New-Object System.Drawing.Size(290, 22)
$gbCommon.Controls.Add($txtCdp)

# --- 3. 実行モード --------------------------------------------------------
$gbMode = New-Object System.Windows.Forms.GroupBox
$gbMode.Text = "3. 実行モード"
$gbMode.Location = New-Object System.Drawing.Point(12, 178)
$gbMode.Size = New-Object System.Drawing.Size(500, 190)
$form.Controls.Add($gbMode)

$rbRun = New-Object System.Windows.Forms.RadioButton
$rbRun.Text = "実行（帳票出力）"
$rbRun.Location = New-Object System.Drawing.Point(15, 22)
$rbRun.Size = New-Object System.Drawing.Size(140, 22)
$rbRun.Checked = $true
$gbMode.Controls.Add($rbRun)

$rbRecord = New-Object System.Windows.Forms.RadioButton
$rbRecord.Text = "記録"
$rbRecord.Location = New-Object System.Drawing.Point(165, 22)
$rbRecord.Size = New-Object System.Drawing.Size(80, 22)
$gbMode.Controls.Add($rbRecord)

$rbList = New-Object System.Windows.Forms.RadioButton
$rbList.Text = "タブ一覧 (-List)"
$rbList.Location = New-Object System.Drawing.Point(255, 22)
$rbList.Size = New-Object System.Drawing.Size(140, 22)
$gbMode.Controls.Add($rbList)

# 記録モードの入力欄
$lblRecJob = New-Object System.Windows.Forms.Label
$lblRecJob.Text = "記録: 帳票名"
$lblRecJob.Location = New-Object System.Drawing.Point(15, 54)
$lblRecJob.Size = New-Object System.Drawing.Size(80, 20)
$gbMode.Controls.Add($lblRecJob)

$txtRecJob = New-Object System.Windows.Forms.TextBox
$txtRecJob.Text = ""
$txtRecJob.Location = New-Object System.Drawing.Point(100, 51)
$txtRecJob.Size = New-Object System.Drawing.Size(150, 22)
$gbMode.Controls.Add($txtRecJob)

$lblAtena = New-Object System.Windows.Forms.Label
$lblAtena.Text = "宛名番号"
$lblAtena.Location = New-Object System.Drawing.Point(265, 54)
$lblAtena.Size = New-Object System.Drawing.Size(60, 20)
$gbMode.Controls.Add($lblAtena)

$txtAtena = New-Object System.Windows.Forms.TextBox
$txtAtena.Text = ""
$txtAtena.Location = New-Object System.Drawing.Point(330, 51)
$txtAtena.Size = New-Object System.Drawing.Size(155, 22)
$gbMode.Controls.Add($txtAtena)

# 実行モードの入力欄
$lblRunJob = New-Object System.Windows.Forms.Label
$lblRunJob.Text = "実行: 帳票"
$lblRunJob.Location = New-Object System.Drawing.Point(15, 88)
$lblRunJob.Size = New-Object System.Drawing.Size(80, 20)
$gbMode.Controls.Add($lblRunJob)

$cmbJob = New-Object System.Windows.Forms.ComboBox
$cmbJob.DropDownStyle = "DropDownList"
$cmbJob.Location = New-Object System.Drawing.Point(100, 85)
$cmbJob.Size = New-Object System.Drawing.Size(240, 22)
$gbMode.Controls.Add($cmbJob)

$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "再読込"
$btnReload.Location = New-Object System.Drawing.Point(350, 84)
$btnReload.Size = New-Object System.Drawing.Size(80, 24)
$gbMode.Controls.Add($btnReload)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "DryRun（出力ボタンを押さない）"
$chkDryRun.Location = New-Object System.Drawing.Point(15, 118)
$chkDryRun.Size = New-Object System.Drawing.Size(230, 22)
$gbMode.Controls.Add($chkDryRun)

$chkEvidence = New-Object System.Windows.Forms.CheckBox
$chkEvidence.Text = "証跡スクショを撮る"
$chkEvidence.Location = New-Object System.Drawing.Point(255, 118)
$chkEvidence.Size = New-Object System.Drawing.Size(160, 22)
$gbMode.Controls.Add($chkEvidence)

$lblLimit = New-Object System.Windows.Forms.Label
$lblLimit.Text = "件数上限(0=無制限)"
$lblLimit.Location = New-Object System.Drawing.Point(15, 148)
$lblLimit.Size = New-Object System.Drawing.Size(120, 20)
$gbMode.Controls.Add($lblLimit)

$txtLimit = New-Object System.Windows.Forms.TextBox
$txtLimit.Text = "0"
$txtLimit.Location = New-Object System.Drawing.Point(140, 145)
$txtLimit.Size = New-Object System.Drawing.Size(60, 22)
$gbMode.Controls.Add($txtLimit)

$lblModeHint = New-Object System.Windows.Forms.Label
$lblModeHint.Text = "記録: 基準画面から一通り操作し、コンソールで Enter を押すと jobs/<帳票名>.json に保存。"
$lblModeHint.Location = New-Object System.Drawing.Point(215, 148)
$lblModeHint.Size = New-Object System.Drawing.Size(275, 34)
$lblModeHint.ForeColor = [System.Drawing.Color]::DimGray
$gbMode.Controls.Add($lblModeHint)

# --- 4. 実行 --------------------------------------------------------------
$gbRun = New-Object System.Windows.Forms.GroupBox
$gbRun.Text = "4. 実行"
$gbRun.Location = New-Object System.Drawing.Point(12, 376)
$gbRun.Size = New-Object System.Drawing.Size(500, 60)
$form.Controls.Add($gbRun)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "実行"
$btnRun.Location = New-Object System.Drawing.Point(390, 20)
$btnRun.Size = New-Object System.Drawing.Size(100, 30)
$gbRun.Controls.Add($btnRun)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ""
$lblStatus.Location = New-Object System.Drawing.Point(12, 442)
$lblStatus.Size = New-Object System.Drawing.Size(500, 30)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($lblStatus)

# ---------------------------------------------------------------------------
# イベント
# ---------------------------------------------------------------------------

# jobs/ の手順ファイルを帳票の選択肢に読み込む（先頭は「(全部)」）
$loadJobs = {
    $cmbJob.Items.Clear()
    [void]$cmbJob.Items.Add("(全部)")
    $dir = Join-Path $RepoRoot "jobs"
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Filter "*.json" | Sort-Object Name | ForEach-Object {
            [void]$cmbJob.Items.Add([System.IO.Path]::GetFileNameWithoutExtension($_.Name))
        }
    }
    $cmbJob.SelectedIndex = 0
}.GetNewClosure()

# モードに応じて入力欄を有効/無効にする
$updateFields = {
    $txtRecJob.Enabled   = $rbRecord.Checked
    $txtAtena.Enabled    = $rbRecord.Checked
    $cmbJob.Enabled      = $rbRun.Checked
    $btnReload.Enabled   = $rbRun.Checked
    $chkDryRun.Enabled   = $rbRun.Checked
    $chkEvidence.Enabled = $rbRun.Checked
    $txtLimit.Enabled    = $rbRun.Checked
}.GetNewClosure()
$rbRun.Add_CheckedChanged($updateFields)
$rbRecord.Add_CheckedChanged($updateFields)
$rbList.Add_CheckedChanged($updateFields)

$btnReload.Add_Click($loadJobs)

# 設定ファイル参照
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "JSON (*.json)|*.json|すべて (*.*)|*.*"
    $dlg.InitialDirectory = Join-Path $RepoRoot "config"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtConfig.Text = $dlg.FileName
    }
})

# Edge起動
$btnEdge.Add_Click({
    $port = $txtPort.Text.Trim()
    if ($port -notmatch '^\d+$') {
        [System.Windows.Forms.MessageBox]::Show("ポートは数値で入力してください。", "入力エラー") | Out-Null
        return
    }
    Start-InConsole "& '.\scripts\start_edge.ps1' -Port $port"
    $lblStatus.Text = "Edgeを起動しました (port: $port)"
})

# 実行（帳票出力 / 記録 / タブ一覧）
$btnRun.Add_Click({
    $cfg = $txtConfig.Text.Trim()
    $cdp = $txtCdp.Text.Trim()
    $a = "-Config `"$cfg`""

    if ($rbList.Checked) {
        $a += " -List"
    }
    elseif ($rbRecord.Checked) {
        $nm = $txtRecJob.Text.Trim()
        $at = $txtAtena.Text.Trim()
        if (-not $nm) {
            [System.Windows.Forms.MessageBox]::Show("記録する帳票名を入力してください。", "入力エラー") | Out-Null
            return
        }
        if (-not $at) {
            [System.Windows.Forms.MessageBox]::Show("記録に使う宛名番号を入力してください。", "入力エラー") | Out-Null
            return
        }
        $a += " -Record -Job `"$nm`" -Atena `"$at`""
    }
    else {
        $lim = $txtLimit.Text.Trim()
        if ($lim -and $lim -notmatch '^\d+$') {
            [System.Windows.Forms.MessageBox]::Show("件数上限は数値で入力してください。", "入力エラー") | Out-Null
            return
        }
        if ($cmbJob.SelectedIndex -gt 0) { $a += " -Job `"$($cmbJob.SelectedItem)`"" }
        if ($chkDryRun.Checked)   { $a += " -DryRun" }
        if ($chkEvidence.Checked) { $a += " -Evidence both" }
        if ($lim -and [int]$lim -gt 0) { $a += " -Limit $lim" }
    }
    if ($cdp) { $a += " -CdpUrl `"$cdp`"" }

    Start-InConsole "& '.\powershell\cdp_print.ps1' $a"
    $lblStatus.Text = "実行しました: $a"
})

& $loadJobs
& $updateFields

[void]$form.ShowDialog()
$form.Dispose()
