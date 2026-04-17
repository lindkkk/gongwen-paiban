# format.ps1 v4 -- single-form main UI with optional modal style editor
# UTF-8 BOM + CRLF. Windows PowerShell 5.1+.

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$InputDocx
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) {
    try { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path } catch {}
}
if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = (Get-Location).Path }

$exe     = Join-Path $scriptDir "gongwen-paiban.exe"
$logFile = Join-Path $scriptDir "paiban-log.txt"

function Log([string]$msg) {
    $line = "[ps1 $(Get-Date -Format 'HH:mm:ss')] $msg"
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch {
        try { [Console]::Out.WriteLine($line) } catch {}
    }
}

function Safe-ShowError([string]$msg) {
    Log "ERROR: $msg"
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show($msg, "公文排版", 'OK', 'Error')
    } catch {
        Log "Safe-ShowError MsgBox failed: $($_.Exception.Message)"
    }
}

# ==========================================================
#   字号映射
# ==========================================================
$SizeNameToPt = [ordered]@{
    "初号" = 42.0; "小初" = 36.0; "一号" = 26.0; "小一" = 24.0;
    "二号" = 22.0; "小二" = 18.0; "三号" = 16.0; "小三" = 15.0;
    "四号" = 14.0; "小四" = 12.0; "五号" = 10.5; "小五" = 9.0; "六号" = 7.5
}
function Pt-To-SizeName([double]$pt) {
    foreach ($k in $SizeNameToPt.Keys) {
        if ([Math]::Abs($SizeNameToPt[$k] - $pt) -lt 0.01) { return $k }
    }
    return "${pt} pt"
}
function SizeName-To-Pt([string]$text) {
    $t = $text.Trim()
    if ($SizeNameToPt.Contains($t)) { return [double]$SizeNameToPt[$t] }
    $t = $t -replace '\s*pt\s*$','' -replace '\s+',''
    $val = 0.0
    if ([double]::TryParse($t, [ref]$val) -and $val -gt 0) { return $val }
    return -1.0
}

# ==========================================================
#   默认 spec（与 C# PresetDefault 对齐）
# ==========================================================
function New-DefaultSpec([string]$role) {
    switch ($role) {
        "title"    { return @{ font='方正小标宋简体'; size_pt=22.0; bold=$false; italic=$false; alignment='center';
                               line_spacing_mode='exact'; line_spacing_value=28.0;
                               spacing_before_pt=0.0; spacing_after_pt=0.0; first_line_indent_chars=0 } }
        "h1"       { return @{ font='黑体'; size_pt=16.0; bold=$false; italic=$false; alignment='';
                               line_spacing_mode='multiple'; line_spacing_value=1.5;
                               spacing_before_pt=6.0; spacing_after_pt=6.0; first_line_indent_chars=2 } }
        "h2"       { return @{ font='楷体_GB2312'; size_pt=16.0; bold=$true; italic=$false; alignment='';
                               line_spacing_mode='multiple'; line_spacing_value=1.5;
                               spacing_before_pt=0.0; spacing_after_pt=0.0; first_line_indent_chars=2 } }
        "h3"       { return @{ font='仿宋_GB2312'; size_pt=16.0; bold=$true; italic=$false; alignment='';
                               line_spacing_mode='multiple'; line_spacing_value=1.5;
                               spacing_before_pt=0.0; spacing_after_pt=0.0; first_line_indent_chars=2 } }
        "body"     { return @{ font='仿宋_GB2312'; size_pt=16.0; bold=$false; italic=$false; alignment='';
                               line_spacing_mode='multiple'; line_spacing_value=1.5;
                               spacing_before_pt=0.0; spacing_after_pt=0.0; first_line_indent_chars=2 } }
        "footnote" { return @{ font='仿宋_GB2312'; size_pt=14.0; bold=$false; italic=$false; alignment='';
                               line_spacing_mode='multiple'; line_spacing_value=1.0;
                               spacing_before_pt=0.0; spacing_after_pt=0.0; first_line_indent_chars=0 } }
    }
}

function Get-AllDefaultSpecs {
    return @{
        title = (New-DefaultSpec 'title');    h1 = (New-DefaultSpec 'h1');
        h2    = (New-DefaultSpec 'h2');       h3 = (New-DefaultSpec 'h3');
        body  = (New-DefaultSpec 'body');     footnote = (New-DefaultSpec 'footnote')
    }
}

# ==========================================================
#   六-Tab 样式编辑器（模态）
#   返回 hashtable 或 $null（取消）
# ==========================================================
function Show-StyleEditor {
    param([hashtable]$initSpecs, $Owner = $null)

    $FontList = @('方正小标宋简体','宋体','仿宋','仿宋_GB2312','黑体','楷体','楷体_GB2312',
                  '微软雅黑','华文仿宋','华文中宋','华文楷体','华文宋体','思源宋体','思源黑体',
                  'Times New Roman','Arial','Calibri')
    $SizeList = @('初号','小初','一号','小一','二号','小二','三号','小三','四号','小四','五号','小五','六号')
    $AlignList = @('(不设置)','居左','居中','居右','两端对齐')
    $AlignValues = @('','left','center','right','justify')
    $LineModeList = @('倍数','固定磅值')
    $LineModeValues = @('multiple','exact')

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "自定义各级样式"
    $form.Size = New-Object System.Drawing.Size(600, 470)
    # 有 Owner 就相对父居中，否则居屏幕；模态子窗口不独立占任务栏
    if ($Owner) { $form.StartPosition = "CenterParent" } else { $form.StartPosition = "CenterScreen" }
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $form.ShowInTaskbar = $false
    $form.ShowIcon = $false

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(10, 10)
    $tabs.Size = New-Object System.Drawing.Size(570, 360)
    $tabs.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $form.Controls.Add($tabs)

    $roleOrder = @(
        @{ k='title';    label='主标题' },
        @{ k='h1';       label='一级' },
        @{ k='h2';       label='二级' },
        @{ k='h3';       label='三级' },
        @{ k='body';     label='正文' },
        @{ k='footnote'; label='脚注' }
    )

    $ctrlBag = @{}  # key -> @{cbFont,cbSize,cbBold,cbItalic,cbAlign,cbLineMode,nudLine,nudBefore,nudAfter,nudIndent}

    foreach ($ro in $roleOrder) {
        $k = $ro.k
        $sp = $initSpecs[$k]
        $tp = New-Object System.Windows.Forms.TabPage
        $tp.Text = $ro.label
        $tabs.TabPages.Add($tp)

        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = 'Fill'
        $tp.Controls.Add($panel)

        function Add-Label2($parent, $text, $x, $y, $w=80) {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = $text
            $lbl.Location = New-Object System.Drawing.Point($x, $y)
            $lbl.Size = New-Object System.Drawing.Size($w, 22)
            $lbl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
            $lbl.TextAlign = 'MiddleLeft'
            $parent.Controls.Add($lbl)
        }

        Add-Label2 $panel "字体:" 20 15
        $cbFont = New-Object System.Windows.Forms.ComboBox
        $cbFont.Location = New-Object System.Drawing.Point(110, 12)
        $cbFont.Size = New-Object System.Drawing.Size(260, 26)
        $cbFont.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbFont.DropDownStyle = 'DropDown'
        foreach ($f in $FontList) { [void]$cbFont.Items.Add($f) }
        $cbFont.Text = [string]$sp.font
        $panel.Controls.Add($cbFont)

        Add-Label2 $panel "字号:" 20 48
        $cbSize = New-Object System.Windows.Forms.ComboBox
        $cbSize.Location = New-Object System.Drawing.Point(110, 45)
        $cbSize.Size = New-Object System.Drawing.Size(110, 26)
        $cbSize.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbSize.DropDownStyle = 'DropDown'
        foreach ($s in $SizeList) { [void]$cbSize.Items.Add($s) }
        $cbSize.Text = (Pt-To-SizeName $sp.size_pt)
        $panel.Controls.Add($cbSize)
        Add-Label2 $panel "(可键入磅数, 如 16.5)" 230 48 200

        $cbBold = New-Object System.Windows.Forms.CheckBox
        $cbBold.Text = "加粗"
        $cbBold.Location = New-Object System.Drawing.Point(110, 80)
        $cbBold.Size = New-Object System.Drawing.Size(80, 24)
        $cbBold.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbBold.Checked = [bool]$sp.bold
        $panel.Controls.Add($cbBold)

        $cbItalic = New-Object System.Windows.Forms.CheckBox
        $cbItalic.Text = "斜体"
        $cbItalic.Location = New-Object System.Drawing.Point(200, 80)
        $cbItalic.Size = New-Object System.Drawing.Size(80, 24)
        $cbItalic.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbItalic.Checked = [bool]$sp.italic
        $panel.Controls.Add($cbItalic)

        Add-Label2 $panel "对齐:" 20 115
        $cbAlign = New-Object System.Windows.Forms.ComboBox
        $cbAlign.Location = New-Object System.Drawing.Point(110, 112)
        $cbAlign.Size = New-Object System.Drawing.Size(130, 26)
        $cbAlign.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbAlign.DropDownStyle = 'DropDownList'
        foreach ($a in $AlignList) { [void]$cbAlign.Items.Add($a) }
        $idxA = [Array]::IndexOf($AlignValues, [string]$sp.alignment)
        if ($idxA -lt 0) { $idxA = 0 }
        $cbAlign.SelectedIndex = $idxA
        $panel.Controls.Add($cbAlign)

        Add-Label2 $panel "行距:" 20 148
        $cbLineMode = New-Object System.Windows.Forms.ComboBox
        $cbLineMode.Location = New-Object System.Drawing.Point(110, 145)
        $cbLineMode.Size = New-Object System.Drawing.Size(100, 26)
        $cbLineMode.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $cbLineMode.DropDownStyle = 'DropDownList'
        foreach ($m in $LineModeList) { [void]$cbLineMode.Items.Add($m) }
        $idxLm = [Array]::IndexOf($LineModeValues, [string]$sp.line_spacing_mode)
        if ($idxLm -lt 0) { $idxLm = 0 }
        $cbLineMode.SelectedIndex = $idxLm
        $panel.Controls.Add($cbLineMode)

        $nudLine = New-Object System.Windows.Forms.NumericUpDown
        $nudLine.Location = New-Object System.Drawing.Point(220, 145)
        $nudLine.Size = New-Object System.Drawing.Size(80, 26)
        $nudLine.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $nudLine.DecimalPlaces = 2
        $nudLine.Minimum = 0.5; $nudLine.Maximum = 200.0; $nudLine.Increment = 0.25
        $nudLine.Value = [decimal]$sp.line_spacing_value
        $panel.Controls.Add($nudLine)

        $lblLineUnit = New-Object System.Windows.Forms.Label
        $lblLineUnit.Location = New-Object System.Drawing.Point(310, 148)
        $lblLineUnit.Size = New-Object System.Drawing.Size(130, 22)
        $lblLineUnit.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
        $lblLineUnit.ForeColor = [System.Drawing.Color]::DimGray
        if ($cbLineMode.SelectedIndex -eq 1) { $lblLineUnit.Text = "磅（固定值）" } else { $lblLineUnit.Text = "倍（1.0=单倍）" }
        $panel.Controls.Add($lblLineUnit)

        $cbLineMode_ref = $cbLineMode
        $lblLineUnit_ref = $lblLineUnit
        $cbLineMode.Add_SelectedIndexChanged({
            if ($cbLineMode_ref.SelectedIndex -eq 1) { $lblLineUnit_ref.Text = "磅（固定值）" }
            else { $lblLineUnit_ref.Text = "倍（1.0=单倍）" }
        }.GetNewClosure())

        # 段前 / 段后：之前段后 label(x=250,w=80) 和 NUD(x=310) 重叠 20px，label 盖在 NUD 上
        # 使点击无效。缩窄 label + 重排坐标：段前 20-195，gap，段后 230-405。
        Add-Label2 $panel "段前:" 20 182 60
        $nudBefore = New-Object System.Windows.Forms.NumericUpDown
        $nudBefore.Location = New-Object System.Drawing.Point(85, 180)
        $nudBefore.Size = New-Object System.Drawing.Size(75, 26)
        $nudBefore.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $nudBefore.DecimalPlaces = 1
        $nudBefore.Minimum = 0; $nudBefore.Maximum = 144; $nudBefore.Increment = 0.5
        $nudBefore.Value = [decimal]$sp.spacing_before_pt
        $panel.Controls.Add($nudBefore)
        Add-Label2 $panel "磅" 165 182 25

        Add-Label2 $panel "段后:" 230 182 60
        $nudAfter = New-Object System.Windows.Forms.NumericUpDown
        $nudAfter.Location = New-Object System.Drawing.Point(295, 180)
        $nudAfter.Size = New-Object System.Drawing.Size(75, 26)
        $nudAfter.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $nudAfter.DecimalPlaces = 1
        $nudAfter.Minimum = 0; $nudAfter.Maximum = 144; $nudAfter.Increment = 0.5
        $nudAfter.Value = [decimal]$sp.spacing_after_pt
        $panel.Controls.Add($nudAfter)
        Add-Label2 $panel "磅" 375 182 25

        Add-Label2 $panel "首行缩进:" 20 215 90
        $nudIndent = New-Object System.Windows.Forms.NumericUpDown
        $nudIndent.Location = New-Object System.Drawing.Point(120, 213)
        $nudIndent.Size = New-Object System.Drawing.Size(80, 26)
        $nudIndent.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $nudIndent.Minimum = 0; $nudIndent.Maximum = 10; $nudIndent.Increment = 1
        $nudIndent.Value = [decimal]$sp.first_line_indent_chars
        $panel.Controls.Add($nudIndent)
        Add-Label2 $panel "字符" 205 215 40

        $ctrlBag[$k] = @{
            cbFont=$cbFont; cbSize=$cbSize; cbBold=$cbBold; cbItalic=$cbItalic;
            cbAlign=$cbAlign; cbLineMode=$cbLineMode; nudLine=$nudLine;
            nudBefore=$nudBefore; nudAfter=$nudAfter; nudIndent=$nudIndent
        }
    }

    # 按钮
    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "恢复公文默认"
    $btnReset.Location = New-Object System.Drawing.Point(10, 385)
    $btnReset.Size = New-Object System.Drawing.Size(130, 32)
    $btnReset.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $form.Controls.Add($btnReset)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "确定"
    $btnOK.Location = New-Object System.Drawing.Point(370, 385)
    $btnOK.Size = New-Object System.Drawing.Size(100, 32)
    $btnOK.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $btnOK.DialogResult = 'OK'
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消"
    $btnCancel.Location = New-Object System.Drawing.Point(480, 385)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $btnCancel.DialogResult = 'Cancel'
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    # 恢复默认：把所有 tab 的控件值重置
    $ctrlBag_ref = $ctrlBag
    $btnReset.Add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show("所有页恢复为公文规范默认值？", "确认", 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
        foreach ($rk in @('title','h1','h2','h3','body','footnote')) {
            $d = New-DefaultSpec $rk
            $c = $ctrlBag_ref[$rk]
            $c.cbFont.Text = [string]$d.font
            $c.cbSize.Text = (Pt-To-SizeName $d.size_pt)
            $c.cbBold.Checked = [bool]$d.bold
            $c.cbItalic.Checked = [bool]$d.italic
            $alignVals = @('','left','center','right','justify')
            $ia = [Array]::IndexOf($alignVals, [string]$d.alignment)
            if ($ia -lt 0) { $ia = 0 }
            $c.cbAlign.SelectedIndex = $ia
            $lmVals = @('multiple','exact')
            $ilm = [Array]::IndexOf($lmVals, [string]$d.line_spacing_mode)
            if ($ilm -lt 0) { $ilm = 0 }
            $c.cbLineMode.SelectedIndex = $ilm
            $c.nudLine.Value = [decimal]$d.line_spacing_value
            $c.nudBefore.Value = [decimal]$d.spacing_before_pt
            $c.nudAfter.Value = [decimal]$d.spacing_after_pt
            $c.nudIndent.Value = [decimal]$d.first_line_indent_chars
        }
    }.GetNewClosure())

    # 带 Owner 打开模态，避免与主窗口 Z-order 争抢 / 被发到任务栏
    if ($Owner) { $result = $form.ShowDialog($Owner) } else { $result = $form.ShowDialog() }
    if ($result -ne 'OK') { return $null }

    # 读回
    $alignVals = @('','left','center','right','justify')
    $lmVals = @('multiple','exact')
    $out = @{}
    foreach ($rk in @('title','h1','h2','h3','body','footnote')) {
        $c = $ctrlBag[$rk]
        $pt = SizeName-To-Pt $c.cbSize.Text
        if ($pt -lt 0) { $pt = $initSpecs[$rk].size_pt }
        $out[$rk] = @{
            font                    = $c.cbFont.Text
            size_pt                 = [double]$pt
            bold                    = [bool]$c.cbBold.Checked
            italic                  = [bool]$c.cbItalic.Checked
            alignment               = [string]$alignVals[$c.cbAlign.SelectedIndex]
            line_spacing_mode       = [string]$lmVals[$c.cbLineMode.SelectedIndex]
            line_spacing_value      = [double]$c.nudLine.Value
            spacing_before_pt       = [double]$c.nudBefore.Value
            spacing_after_pt        = [double]$c.nudAfter.Value
            first_line_indent_chars = [int]$c.nudIndent.Value
        }
    }
    return $out
}

# ==========================================================
#   写 JSON 配置
# ==========================================================
function Write-ConfigJson($specs, $source, $h1m, $h2m, $h3m, $path) {
    $o = [ordered]@{
        source = $source
        h1_marker = if ($h1m) { $h1m } else { $null }
        h2_marker = if ($h2m) { $h2m } else { $null }
        h3_marker = if ($h3m) { $h3m } else { $null }
        title = $specs.title; h1 = $specs.h1; h2 = $specs.h2
        h3 = $specs.h3; body = $specs.body; footnote = $specs.footnote
    }
    $json = $o | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding $false))
}

# ==========================================================
#   主流程：单窗口
# ==========================================================
try {
    Log "format.ps1 v4 start.  PS=$($PSVersionTable.PSVersion)"
    Log "scriptDir='$scriptDir'  InputDocx='$InputDocx'"

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Log "WinForms loaded"

    if ([string]::IsNullOrEmpty($InputDocx)) { Safe-ShowError "ps1 未收到输入文件。"; exit 1 }
    if (-not (Test-Path -LiteralPath $InputDocx)) { Safe-ShowError "输入文件不存在：`n$InputDocx"; exit 1 }
    $ext = [System.IO.Path]::GetExtension($InputDocx)
    if ($ext.ToLower() -ne ".docx") { Safe-ShowError "只支持 .docx。如果是 .doc 请先在 Word/WPS 里另存为 .docx。"; exit 1 }
    if (-not (Test-Path -LiteralPath $exe)) { Safe-ShowError "找不到 gongwen-paiban.exe`n期望在：$exe"; exit 1 }

    # ========== 主窗口 ==========
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "公文排版工具"
    $form.Size = New-Object System.Drawing.Size(640, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    # 关键：不能 TopMost。TopMost 会和模态子窗口（样式编辑器）争 Z-order，
    # 导致子窗口一闪而过被丢进任务栏。下面用 BringToFront + Activate 做兜底。
    $form.TopMost = $false
    $form.Add_Shown({ $form.Activate() }.GetNewClosure())

    # 输入文件显示
    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Text = "输入文件：  $InputDocx"
    $lblFile.Location = New-Object System.Drawing.Point(15, 12)
    $lblFile.Size = New-Object System.Drawing.Size(600, 20)
    $lblFile.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $lblFile.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblFile)

    # ----- Group 1: 文档来源 -----
    $grp1 = New-Object System.Windows.Forms.GroupBox
    $grp1.Text = "① 文档来源"
    $grp1.Location = New-Object System.Drawing.Point(15, 40)
    $grp1.Size = New-Object System.Drawing.Size(600, 75)
    $grp1.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $form.Controls.Add($grp1)

    $rbWps = New-Object System.Windows.Forms.RadioButton
    $rbWps.Text = "WPS"
    $rbWps.Location = New-Object System.Drawing.Point(20, 30)
    $rbWps.Size = New-Object System.Drawing.Size(100, 28)
    $rbWps.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $grp1.Controls.Add($rbWps)

    $rbOffice = New-Object System.Windows.Forms.RadioButton
    $rbOffice.Text = "Microsoft Office"
    $rbOffice.Location = New-Object System.Drawing.Point(140, 30)
    $rbOffice.Size = New-Object System.Drawing.Size(180, 28)
    $rbOffice.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $grp1.Controls.Add($rbOffice)

    $rbAuto = New-Object System.Windows.Forms.RadioButton
    $rbAuto.Text = "不确定 / 自动"
    $rbAuto.Location = New-Object System.Drawing.Point(340, 30)
    $rbAuto.Size = New-Object System.Drawing.Size(180, 28)
    $rbAuto.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $rbAuto.Checked = $true   # 默认
    $grp1.Controls.Add($rbAuto)

    # ----- Group 2: 标题编号 -----
    $grp2 = New-Object System.Windows.Forms.GroupBox
    $grp2.Text = "② 标题编号识别"
    $grp2.Location = New-Object System.Drawing.Point(15, 125)
    $grp2.Size = New-Object System.Drawing.Size(600, 190)
    $grp2.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $form.Controls.Add($grp2)

    # 提示行（说明留空即自动识别）
    $lblInstr2 = New-Object System.Windows.Forms.Label
    $lblInstr2.Text = "如果知道各级标题的编号，填进去会更准；留空则由程序自动识别。"
    $lblInstr2.Location = New-Object System.Drawing.Point(15, 28)
    $lblInstr2.Size = New-Object System.Drawing.Size(575, 22)
    $lblInstr2.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $grp2.Controls.Add($lblInstr2)

    function Make-MarkerRow($grp, $label, $hint, $y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $label
        $l.Location = New-Object System.Drawing.Point(30, $y)
        $l.Size = New-Object System.Drawing.Size(100, 22)
        $l.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $grp.Controls.Add($l)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(135, ($y - 2))
        $tb.Size = New-Object System.Drawing.Size(150, 26)
        $tb.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        # 默认就启用，不再需要 checkbox 激活——去掉这个开关可避免用户忘勾选
        $grp.Controls.Add($tb)
        $h = New-Object System.Windows.Forms.Label
        $h.Text = $hint
        $h.Location = New-Object System.Drawing.Point(295, $y)
        $h.Size = New-Object System.Drawing.Size(285, 22)
        $h.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $h.ForeColor = [System.Drawing.Color]::DimGray
        $grp.Controls.Add($h)
        return $tb
    }
    $tbH1 = Make-MarkerRow $grp2 "一级标题:" "例：一、  /  1.  /  第一章     留空=自动" 60
    $tbH2 = Make-MarkerRow $grp2 "二级标题:" "例：（一）  /  1.1  /  第一节     留空=自动" 95
    $tbH3 = Make-MarkerRow $grp2 "三级标题:" "例：1.  /  (1)  /  1.1.1     留空=自动" 130

    $lblTip2 = New-Object System.Windows.Forms.Label
    $lblTip2.Text = "容错：不带标点也行，如 `"一`" = `"一、`"；括号中英文可混用"
    $lblTip2.Location = New-Object System.Drawing.Point(30, 162)
    $lblTip2.Size = New-Object System.Drawing.Size(550, 20)
    $lblTip2.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $lblTip2.ForeColor = [System.Drawing.Color]::DimGray
    $grp2.Controls.Add($lblTip2)

    # ----- Group 3: 排版样式 -----
    $grp3 = New-Object System.Windows.Forms.GroupBox
    $grp3.Text = "③ 排版样式"
    $grp3.Location = New-Object System.Drawing.Point(15, 325)
    $grp3.Size = New-Object System.Drawing.Size(600, 145)
    $grp3.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $form.Controls.Add($grp3)

    $cbCustom = New-Object System.Windows.Forms.CheckBox
    $cbCustom.Text = "自定义各级字体字号 / 行距 / 缩进（不勾选 = 使用内置公文规范）"
    $cbCustom.Location = New-Object System.Drawing.Point(15, 28)
    $cbCustom.Size = New-Object System.Drawing.Size(570, 26)
    $cbCustom.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $grp3.Controls.Add($cbCustom)

    $btnEditor = New-Object System.Windows.Forms.Button
    $btnEditor.Text = "打开样式编辑器..."
    $btnEditor.Location = New-Object System.Drawing.Point(30, 62)
    $btnEditor.Size = New-Object System.Drawing.Size(180, 34)
    $btnEditor.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $btnEditor.Enabled = $false
    $grp3.Controls.Add($btnEditor)

    $lblEditorStatus = New-Object System.Windows.Forms.Label
    $lblEditorStatus.Text = "（未自定义，使用公文默认）"
    $lblEditorStatus.Location = New-Object System.Drawing.Point(220, 70)
    $lblEditorStatus.Size = New-Object System.Drawing.Size(360, 22)
    $lblEditorStatus.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $lblEditorStatus.ForeColor = [System.Drawing.Color]::DimGray
    $grp3.Controls.Add($lblEditorStatus)

    $lblTip3 = New-Object System.Windows.Forms.Label
    $lblTip3.Text = "内置规范：标题方正小标宋二号居中；一级黑体三号；二级楷体三号加粗；正文仿宋三号 1.5 倍首行 2 字符"
    $lblTip3.Location = New-Object System.Drawing.Point(30, 110)
    $lblTip3.Size = New-Object System.Drawing.Size(555, 22)
    $lblTip3.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $lblTip3.ForeColor = [System.Drawing.Color]::DimGray
    $grp3.Controls.Add($lblTip3)

    # 样式编辑器状态：关键修复——
    # 用共享 hashtable 而不是 $script:customSpecs。PowerShell 的 .GetNewClosure()
    # 会把 $script: 作用域重绑到闭包内部快照，赋值不会穿出闭包；而对 hashtable 的
    # 键赋值（$state.customSpecs = $r）只是在共享引用上改值，外层能正确读到。
    $state = @{ customSpecs = $null }

    $cbCustom_ref = $cbCustom
    $btnEditor_ref = $btnEditor
    $lblEditorStatus_ref = $lblEditorStatus
    $state_ref = $state
    $cbCustom.Add_CheckedChanged({
        $btnEditor_ref.Enabled = $cbCustom_ref.Checked
        if (-not $cbCustom_ref.Checked) {
            $state_ref.customSpecs = $null
            $lblEditorStatus_ref.Text = "（未自定义，使用公文默认）"
            $lblEditorStatus_ref.ForeColor = [System.Drawing.Color]::DimGray
        } else {
            if ($null -eq $state_ref.customSpecs) {
                $lblEditorStatus_ref.Text = "（请点左侧按钮打开编辑器设置参数）"
                $lblEditorStatus_ref.ForeColor = [System.Drawing.Color]::OrangeRed
            }
        }
    }.GetNewClosure())

    $mainForm_ref = $form
    $btnEditor.Add_Click({
        $init = if ($state_ref.customSpecs) { $state_ref.customSpecs } else { Get-AllDefaultSpecs }
        # 把主窗口作为 Owner 传进去，模态对话框才能正确与父建立 Z-order
        $r = Show-StyleEditor -initSpecs $init -Owner $mainForm_ref
        if ($null -ne $r) {
            $state_ref.customSpecs = $r
            $lblEditorStatus_ref.Text = "√ 已保存自定义样式"
            $lblEditorStatus_ref.ForeColor = [System.Drawing.Color]::Green
            Log "style editor returned custom specs  (body.font=$($r.body.font)  body.size_pt=$($r.body.size_pt))"
        }
    }.GetNewClosure())

    # ----- 底部按钮 -----
    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = "开始排版"
    $btnGo.Location = New-Object System.Drawing.Point(370, 500)
    $btnGo.Size = New-Object System.Drawing.Size(115, 38)
    $btnGo.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnGo.DialogResult = 'OK'
    $form.AcceptButton = $btnGo
    $form.Controls.Add($btnGo)

    $btnCancelMain = New-Object System.Windows.Forms.Button
    $btnCancelMain.Text = "取消"
    $btnCancelMain.Location = New-Object System.Drawing.Point(495, 500)
    $btnCancelMain.Size = New-Object System.Drawing.Size(115, 38)
    $btnCancelMain.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnCancelMain.DialogResult = 'Cancel'
    $form.CancelButton = $btnCancelMain
    $form.Controls.Add($btnCancelMain)

    # 显示
    $mainResult = $form.ShowDialog()
    Log "main form closed, DialogResult=$mainResult"
    if ($mainResult -ne 'OK') { Log "user cancelled"; exit 1 }

    # 读取选择
    if     ($rbWps.Checked)    { $source = 'wps' }
    elseif ($rbOffice.Checked) { $source = 'office' }
    else                       { $source = 'auto' }

    # 去掉 checkbox 开关，直接读三个文本框；空字符串 = 不指定（exe 自动识别）
    $h1m = $tbH1.Text.Trim()
    $h2m = $tbH2.Text.Trim()
    $h3m = $tbH3.Text.Trim()
    Log "source=$source  h1='$h1m'  h2='$h2m'  h3='$h3m'  customSpecs=$($null -ne $state.customSpecs)"

    # ========== 调 exe ==========
    $dir  = [System.IO.Path]::GetDirectoryName($InputDocx)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputDocx)
    $out  = Join-Path $dir "${name}_formatted.docx"

    $exeArgs = @("format", $InputDocx, "-o", $out, "--source", $source)

    if ($state.customSpecs) {
        $cfgPath = Join-Path $env:TEMP "gongwen-paiban-config-$((Get-Date).Ticks).json"
        Write-ConfigJson $state.customSpecs $source $h1m $h2m $h3m $cfgPath
        Log "wrote custom config: $cfgPath"
        # 把用户看到的每级 font/size 也列出来，直接在 log 里就能看出是否符合预期
        foreach ($rk in @('title','h1','h2','h3','body','footnote')) {
            $s = $state.customSpecs[$rk]
            if ($s) { Log ("  " + $rk + "  font=" + $s.font + "  size_pt=" + $s.size_pt + "  bold=" + $s.bold + "  italic=" + $s.italic + "  line=" + $s.line_spacing_mode + ":" + $s.line_spacing_value + "  before=" + $s.spacing_before_pt + "  after=" + $s.spacing_after_pt + "  indent=" + $s.first_line_indent_chars) }
        }
        # 再把 JSON 文件内容直接 dump 到 log，用户 one-stop 看到真实传给 exe 的数据
        Log "--- JSON begin ---"
        foreach ($ln in (Get-Content -LiteralPath $cfgPath -Encoding UTF8)) { Log ("  " + $ln) }
        Log "--- JSON end ---"
        $exeArgs += @("--config", $cfgPath)
    }
    if ($h1m) { $exeArgs += @("--h1-marker", $h1m) }
    if ($h2m) { $exeArgs += @("--h2-marker", $h2m) }
    if ($h3m) { $exeArgs += @("--h3-marker", $h3m) }

    Log "about to invoke exe: '$exe' $($exeArgs -join ' | ')"
    $exeOutput = & $exe @exeArgs 2>&1 | Out-String
    $exitCode = $LASTEXITCODE

    Log "exe exit=$exitCode"
    foreach ($ln in ($exeOutput -split "`r?`n")) {
        if ($null -ne $ln -and $ln.Length -gt 0) { Log "  | $ln" }
    }

    if ($exitCode -ne 0) {
        Safe-ShowError "排版失败（退出码 $exitCode）。`n`n$exeOutput`n`n详情见 paiban-log.txt"
        exit $exitCode
    }

    $done = [System.Windows.Forms.MessageBox]::Show(
        "完成！`n`n输出：$out`n`n要打开所在文件夹吗？",
        "公文排版", 'YesNo', 'Information')
    if ($done -eq 'Yes') { Start-Process "explorer.exe" -ArgumentList "/select,`"$out`"" }
    Log "done ok"
    exit 0

} catch {
    $err = $_
    Log "UNHANDLED PS EXCEPTION at line $($err.InvocationInfo.ScriptLineNumber)"
    Log ($err | Out-String)
    try {
        Safe-ShowError ("ps1 发生异常（行 $($err.InvocationInfo.ScriptLineNumber)）：`n`n" + $err.Exception.Message + "`n`n详情见 paiban-log.txt")
    } catch {}
    exit 99
}
