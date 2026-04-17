# format.ps1 v3 -- gongwen-paiban interactive launcher with style editor
# UTF-8 BOM + CRLF. Requires Windows PowerShell 5.1 or later.

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
#   字号名 ↔ pt 映射（用户常用）
# ==========================================================
$SizeNameToPt = [ordered]@{
    "初号"  = 42.0
    "小初"  = 36.0
    "一号"  = 26.0
    "小一"  = 24.0
    "二号"  = 22.0
    "小二"  = 18.0
    "三号"  = 16.0
    "小三"  = 15.0
    "四号"  = 14.0
    "小四"  = 12.0
    "五号"  = 10.5
    "小五"  = 9.0
    "六号"  = 7.5
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
    # 去掉 "pt" / " pt" 后缀
    $t = $t -replace '\s*pt\s*$',''
    $t = $t -replace '\s+',''
    $val = 0.0
    if ([double]::TryParse($t, [ref]$val) -and $val -gt 0) { return $val }
    return -1.0
}

# ==========================================================
#   各角色的默认 spec（与 C# 侧 PresetDefault 一致）
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

# ==========================================================
#   构造一个"角色样式编辑面板"（TabControl 的一个 Tab）
#   返回 @{ Panel=...; Refresh=...; Read=... }
# ==========================================================
function Build-RoleTab([hashtable]$initSpec, [string]$role) {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'

    $FontList = @('方正小标宋简体','宋体','仿宋','仿宋_GB2312','黑体','楷体','楷体_GB2312',
                  '微软雅黑','华文仿宋','华文中宋','华文楷体','华文宋体','思源宋体','思源黑体',
                  'Times New Roman','Arial','Calibri')
    $SizeList = @('初号','小初','一号','小一','二号','小二','三号','小三','四号','小四','五号','小五','六号')
    $AlignList = @('(不设置)','居左','居中','居右','两端对齐')
    $AlignValues = @('','left','center','right','justify')
    $LineModeList = @('倍数','固定磅值')
    $LineModeValues = @('multiple','exact')

    function Add-Label($text, $x, $y, $w=80) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $text
        $lbl.Location = New-Object System.Drawing.Point($x, $y)
        $lbl.Size = New-Object System.Drawing.Size($w, 22)
        $lbl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $lbl.TextAlign = 'MiddleLeft'
        $panel.Controls.Add($lbl)
        return $lbl
    }

    # 字体
    Add-Label "字体:" 20 15 | Out-Null
    $cbFont = New-Object System.Windows.Forms.ComboBox
    $cbFont.Location = New-Object System.Drawing.Point(105, 12)
    $cbFont.Size = New-Object System.Drawing.Size(260, 26)
    $cbFont.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbFont.DropDownStyle = 'DropDown'  # editable
    $cbFont.Items.AddRange($FontList)
    $cbFont.Text = [string]$initSpec.font
    $panel.Controls.Add($cbFont)

    # 字号
    Add-Label "字号:" 20 48 | Out-Null
    $cbSize = New-Object System.Windows.Forms.ComboBox
    $cbSize.Location = New-Object System.Drawing.Point(105, 45)
    $cbSize.Size = New-Object System.Drawing.Size(110, 26)
    $cbSize.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbSize.DropDownStyle = 'DropDown'
    $cbSize.Items.AddRange($SizeList)
    $cbSize.Text = (Pt-To-SizeName $initSpec.size_pt)
    $panel.Controls.Add($cbSize)

    $lblSizeHint = New-Object System.Windows.Forms.Label
    $lblSizeHint.Location = New-Object System.Drawing.Point(225, 48)
    $lblSizeHint.Size = New-Object System.Drawing.Size(200, 22)
    $lblSizeHint.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $lblSizeHint.ForeColor = [System.Drawing.Color]::DimGray
    $lblSizeHint.Text = "(可直接输入磅数，如 16.5)"
    $panel.Controls.Add($lblSizeHint)

    # 加粗 / 斜体
    $cbBold = New-Object System.Windows.Forms.CheckBox
    $cbBold.Text = "加粗"
    $cbBold.Location = New-Object System.Drawing.Point(105, 80)
    $cbBold.Size = New-Object System.Drawing.Size(80, 24)
    $cbBold.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbBold.Checked = [bool]$initSpec.bold
    $panel.Controls.Add($cbBold)

    $cbItalic = New-Object System.Windows.Forms.CheckBox
    $cbItalic.Text = "斜体"
    $cbItalic.Location = New-Object System.Drawing.Point(200, 80)
    $cbItalic.Size = New-Object System.Drawing.Size(80, 24)
    $cbItalic.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbItalic.Checked = [bool]$initSpec.italic
    $panel.Controls.Add($cbItalic)

    # 对齐
    Add-Label "对齐:" 20 115 | Out-Null
    $cbAlign = New-Object System.Windows.Forms.ComboBox
    $cbAlign.Location = New-Object System.Drawing.Point(105, 112)
    $cbAlign.Size = New-Object System.Drawing.Size(120, 26)
    $cbAlign.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbAlign.DropDownStyle = 'DropDownList'
    $cbAlign.Items.AddRange($AlignList)
    $idxA = [Array]::IndexOf($AlignValues, [string]$initSpec.alignment)
    if ($idxA -lt 0) { $idxA = 0 }
    $cbAlign.SelectedIndex = $idxA
    $panel.Controls.Add($cbAlign)

    # 行距模式 + 值
    Add-Label "行距:" 20 148 | Out-Null
    $cbLineMode = New-Object System.Windows.Forms.ComboBox
    $cbLineMode.Location = New-Object System.Drawing.Point(105, 145)
    $cbLineMode.Size = New-Object System.Drawing.Size(90, 26)
    $cbLineMode.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $cbLineMode.DropDownStyle = 'DropDownList'
    $cbLineMode.Items.AddRange($LineModeList)
    $idxLm = [Array]::IndexOf($LineModeValues, [string]$initSpec.line_spacing_mode)
    if ($idxLm -lt 0) { $idxLm = 0 }
    $cbLineMode.SelectedIndex = $idxLm
    $panel.Controls.Add($cbLineMode)

    $nudLine = New-Object System.Windows.Forms.NumericUpDown
    $nudLine.Location = New-Object System.Drawing.Point(205, 145)
    $nudLine.Size = New-Object System.Drawing.Size(80, 26)
    $nudLine.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $nudLine.DecimalPlaces = 2
    $nudLine.Minimum = 0.5
    $nudLine.Maximum = 200.0
    $nudLine.Increment = 0.25
    $nudLine.Value = [decimal]$initSpec.line_spacing_value
    $panel.Controls.Add($nudLine)

    $lblLineUnit = New-Object System.Windows.Forms.Label
    $lblLineUnit.Location = New-Object System.Drawing.Point(295, 148)
    $lblLineUnit.Size = New-Object System.Drawing.Size(100, 22)
    $lblLineUnit.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $lblLineUnit.ForeColor = [System.Drawing.Color]::DimGray
    $panel.Controls.Add($lblLineUnit)

    $updateUnit = {
        if ($cbLineMode.SelectedIndex -eq 1) { $lblLineUnit.Text = "磅（固定值）" }
        else { $lblLineUnit.Text = "倍（1.0=单倍）" }
    }
    & $updateUnit
    $cbLineMode.Add_SelectedIndexChanged($updateUnit)

    # 段前
    Add-Label "段前:" 20 182 | Out-Null
    $nudBefore = New-Object System.Windows.Forms.NumericUpDown
    $nudBefore.Location = New-Object System.Drawing.Point(105, 180)
    $nudBefore.Size = New-Object System.Drawing.Size(80, 26)
    $nudBefore.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $nudBefore.DecimalPlaces = 1
    $nudBefore.Minimum = 0; $nudBefore.Maximum = 144; $nudBefore.Increment = 0.5
    $nudBefore.Value = [decimal]$initSpec.spacing_before_pt
    $panel.Controls.Add($nudBefore)
    Add-Label "磅" 190 182 30 | Out-Null

    # 段后
    Add-Label "段后:" 250 182 | Out-Null
    $nudAfter = New-Object System.Windows.Forms.NumericUpDown
    $nudAfter.Location = New-Object System.Drawing.Point(305, 180)
    $nudAfter.Size = New-Object System.Drawing.Size(80, 26)
    $nudAfter.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $nudAfter.DecimalPlaces = 1
    $nudAfter.Minimum = 0; $nudAfter.Maximum = 144; $nudAfter.Increment = 0.5
    $nudAfter.Value = [decimal]$initSpec.spacing_after_pt
    $panel.Controls.Add($nudAfter)
    Add-Label "磅" 390 182 30 | Out-Null

    # 首行缩进
    Add-Label "首行缩进:" 20 215 90 | Out-Null
    $nudIndent = New-Object System.Windows.Forms.NumericUpDown
    $nudIndent.Location = New-Object System.Drawing.Point(115, 213)
    $nudIndent.Size = New-Object System.Drawing.Size(80, 26)
    $nudIndent.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $nudIndent.Minimum = 0; $nudIndent.Maximum = 10; $nudIndent.Increment = 1
    $nudIndent.Value = [decimal]$initSpec.first_line_indent_chars
    $panel.Controls.Add($nudIndent)
    Add-Label "字符" 200 215 40 | Out-Null

    # Read 函数：把控件值读回成 hashtable
    $readFn = {
        $pt = SizeName-To-Pt $cbSize.Text
        if ($pt -lt 0) { $pt = $initSpec.size_pt }
        return @{
            font                    = $cbFont.Text
            size_pt                 = [double]$pt
            bold                    = [bool]$cbBold.Checked
            italic                  = [bool]$cbItalic.Checked
            alignment               = [string]$AlignValues[$cbAlign.SelectedIndex]
            line_spacing_mode       = [string]$LineModeValues[$cbLineMode.SelectedIndex]
            line_spacing_value      = [double]$nudLine.Value
            spacing_before_pt       = [double]$nudBefore.Value
            spacing_after_pt        = [double]$nudAfter.Value
            first_line_indent_chars = [int]$nudIndent.Value
        }
    }.GetNewClosure()

    return @{ Panel = $panel; Read = $readFn }
}

# ==========================================================
#   打开一个 TabControl 让用户编辑六个角色
# ==========================================================
function Open-StyleEditor {
    param([hashtable]$initSpecs)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "公文排版 - 自定义各级样式"
    $form.Size = New-Object System.Drawing.Size(580, 430)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $form.TopMost = $true

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(10, 10)
    $tabs.Size = New-Object System.Drawing.Size(550, 320)
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

    $tabHandles = @{}
    foreach ($ro in $roleOrder) {
        $tp = New-Object System.Windows.Forms.TabPage
        $tp.Text = $ro.label
        $tabs.TabPages.Add($tp)
        $h = Build-RoleTab $initSpecs[$ro.k] $ro.k
        $tp.Controls.Add($h.Panel)
        $tabHandles[$ro.k] = $h
    }

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "恢复默认"
    $btnReset.Location = New-Object System.Drawing.Point(10, 345)
    $btnReset.Size = New-Object System.Drawing.Size(110, 32)
    $btnReset.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $btnReset.Add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show("所有页恢复为公文规范默认值？", "确认", 'YesNo', 'Question')
        if ($r -eq 'Yes') {
            foreach ($ro in $roleOrder) {
                $tabs.TabPages[[Array]::IndexOf($roleOrder.k, $ro.k)].Controls.Clear()
            }
            # 简单化：关窗后用默认重开
            $form.Tag = 'reset'
            $form.Close()
        }
    })
    $form.Controls.Add($btnReset)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "确定"
    $btnOK.Location = New-Object System.Drawing.Point(350, 345)
    $btnOK.Size = New-Object System.Drawing.Size(95, 32)
    $btnOK.DialogResult = 'OK'
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消"
    $btnCancel.Location = New-Object System.Drawing.Point(455, 345)
    $btnCancel.Size = New-Object System.Drawing.Size(95, 32)
    $btnCancel.DialogResult = 'Cancel'
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($form.Tag -eq 'reset') {
        $defaults = @{}
        foreach ($ro in $roleOrder) { $defaults[$ro.k] = (New-DefaultSpec $ro.k) }
        return Open-StyleEditor -initSpecs $defaults
    }
    if ($result -ne 'OK') { return $null }

    $out = @{}
    foreach ($ro in $roleOrder) {
        $out[$ro.k] = & $tabHandles[$ro.k].Read
    }
    return $out
}

# ==========================================================
#   写 JSON 配置（用户自定义时）
# ==========================================================
function Write-ConfigJson($specs, $source, $h1m, $h2m, $h3m, $path) {
    $o = [ordered]@{
        source = $source
        h1_marker = if ($h1m) { $h1m } else { $null }
        h2_marker = if ($h2m) { $h2m } else { $null }
        h3_marker = if ($h3m) { $h3m } else { $null }
        title = $specs.title
        h1 = $specs.h1
        h2 = $specs.h2
        h3 = $specs.h3
        body = $specs.body
        footnote = $specs.footnote
    }
    $json = $o | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding $false))
}

# ==========================================================
#   主流程
# ==========================================================
try {
    Log "format.ps1 v3 start.  PS=$($PSVersionTable.PSVersion)"
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

    # ---- Step 1: 来源 ----
    $srcForm = New-Object System.Windows.Forms.Form
    $srcForm.Text = "公文排版 - 第 1 步：文档来源"
    $srcForm.Size = New-Object System.Drawing.Size(440, 220)
    $srcForm.StartPosition = "CenterScreen"; $srcForm.FormBorderStyle = "FixedDialog"
    $srcForm.MaximizeBox = $false; $srcForm.MinimizeBox = $false; $srcForm.TopMost = $true
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "这份文档是用什么软件做的？"; $lbl.Location = New-Object System.Drawing.Point(20, 25)
    $lbl.Size = New-Object System.Drawing.Size(400, 35); $lbl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $srcForm.Controls.Add($lbl)
    $script:sourceChoice = $null
    foreach ($btn in @(
        @{t='WPS';          x=20;  v='wps'},
        @{t='Microsoft Office'; x=150; v='office'},
        @{t='不确定';        x=305; v='auto'})) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $btn.t; $b.Location = New-Object System.Drawing.Point([int]$btn.x, 85)
        $b.Size = New-Object System.Drawing.Size(($(if($btn.v -eq 'office'){140}else{115})), 50)
        $b.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
        $b.Tag = $btn.v
        $b.Add_Click({ $script:sourceChoice = $this.Tag; $srcForm.Close() }.GetNewClosure())
        $srcForm.Controls.Add($b)
    }
    [void]$srcForm.ShowDialog()
    if ([string]::IsNullOrEmpty($script:sourceChoice)) { Log "step 1 cancelled"; exit 1 }
    $source = $script:sourceChoice
    Log "step 1 done, source='$source'"

    # ---- Step 2: 编号方式（原有逻辑） ----
    $msg = "您知道这份文档的一级/二级/三级标题前的编号格式吗？`n`n・「是」：自己指定各级的编号样例`n・「否」：让程序自动判断"
    $knowResult = [System.Windows.Forms.MessageBox]::Show($msg, "公文排版 - 第 2 步：编号方式", 'YesNoCancel', 'Question')
    if ($knowResult -eq 'Cancel') { Log "step 2 cancelled"; exit 1 }
    $h1m = ""; $h2m = ""; $h3m = ""
    if ($knowResult -eq 'Yes') {
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "公文排版 - 第 3 步：指定编号样例"
        $inputForm.Size = New-Object System.Drawing.Size(560, 360)
        $inputForm.StartPosition = "CenterScreen"; $inputForm.FormBorderStyle = "FixedDialog"
        $inputForm.MaximizeBox = $false; $inputForm.MinimizeBox = $false; $inputForm.TopMost = $true
        $tip = New-Object System.Windows.Forms.Label
        $tip.Text = "只需填一个样例，程序智能补齐。留空 = 该级交给自动判断。"
        $tip.Location = New-Object System.Drawing.Point(20, 15); $tip.Size = New-Object System.Drawing.Size(520, 30)
        $tip.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9); $tip.ForeColor = [System.Drawing.Color]::DimGray
        $inputForm.Controls.Add($tip)
        $tbl = @{}; $i = 0
        foreach ($row in @(
            @{ k='H1'; label='一级标题：'; hint='例：一、  /  1.  /  第一章' },
            @{ k='H2'; label='二级标题：'; hint='例：（一）  /  1.1  /  第一节' },
            @{ k='H3'; label='三级标题：'; hint='例：1.  /  (1)  /  1.1.1' })) {
            $y = 60 + $i * 42
            $l = New-Object System.Windows.Forms.Label; $l.Text = $row.label
            $l.Location = New-Object System.Drawing.Point(20, $y); $l.Size = New-Object System.Drawing.Size(100, 24)
            $l.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Location = New-Object System.Drawing.Point(125, ($y - 2)); $tb.Size = New-Object System.Drawing.Size(160, 28)
            $tb.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
            $h = New-Object System.Windows.Forms.Label; $h.Text = $row.hint
            $h.Location = New-Object System.Drawing.Point(295, $y); $h.Size = New-Object System.Drawing.Size(245, 24)
            $h.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9); $h.ForeColor = [System.Drawing.Color]::DimGray
            $inputForm.Controls.AddRange(@($l, $tb, $h))
            $tbl[$row.k] = $tb; $i++
        }
        $btnOK2 = New-Object System.Windows.Forms.Button
        $btnOK2.Text = "确定"; $btnOK2.Location = New-Object System.Drawing.Point(330, 275)
        $btnOK2.Size = New-Object System.Drawing.Size(95, 34); $btnOK2.DialogResult = 'OK'
        $inputForm.AcceptButton = $btnOK2
        $btnCancel2 = New-Object System.Windows.Forms.Button
        $btnCancel2.Text = "取消"; $btnCancel2.Location = New-Object System.Drawing.Point(435, 275)
        $btnCancel2.Size = New-Object System.Drawing.Size(95, 34); $btnCancel2.DialogResult = 'Cancel'
        $inputForm.CancelButton = $btnCancel2
        $inputForm.Controls.AddRange(@($btnOK2, $btnCancel2))
        $r = $inputForm.ShowDialog()
        if ($r -eq 'Cancel') { Log "step 3 cancelled"; exit 1 }
        $h1m = $tbl['H1'].Text.Trim(); $h2m = $tbl['H2'].Text.Trim(); $h3m = $tbl['H3'].Text.Trim()
        Log "markers: H1='$h1m' H2='$h2m' H3='$h3m'"
    }

    # ---- Step 4: 排版样式（默认 vs 自定义） ----
    $styleMsg = "排版格式要使用:`n`n・「是」：使用内置公文规范`n  （标题方正小标宋二号，正文仿宋_GB2312三号 1.5 倍行距 首行缩进 2 字符 等）`n`n・「否」：我要自定义各级字体字号行距等"
    $useDefault = [System.Windows.Forms.MessageBox]::Show($styleMsg, "公文排版 - 第 4 步：排版样式", 'YesNoCancel', 'Question')
    if ($useDefault -eq 'Cancel') { Log "step 4 cancelled"; exit 1 }
    Log "step 4 answer=$useDefault"

    $customSpecs = $null
    if ($useDefault -eq 'No') {
        $initSpecs = @{
            title = (New-DefaultSpec 'title');  h1 = (New-DefaultSpec 'h1');
            h2    = (New-DefaultSpec 'h2');     h3 = (New-DefaultSpec 'h3');
            body  = (New-DefaultSpec 'body');   footnote = (New-DefaultSpec 'footnote')
        }
        $customSpecs = Open-StyleEditor -initSpecs $initSpecs
        if ($null -eq $customSpecs) { Log "style editor cancelled"; exit 1 }
        Log "style editor returned custom specs"
    }

    # ---- Run exe ----
    $dir  = [System.IO.Path]::GetDirectoryName($InputDocx)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputDocx)
    $out  = Join-Path $dir "${name}_formatted.docx"

    $exeArgs = @("format", $InputDocx, "-o", $out, "--source", $source)
    if ($h1m) { $exeArgs += @("--h1-marker", $h1m) }
    if ($h2m) { $exeArgs += @("--h2-marker", $h2m) }
    if ($h3m) { $exeArgs += @("--h3-marker", $h3m) }

    if ($customSpecs) {
        $cfgPath = Join-Path $env:TEMP "gongwen-paiban-config-$((Get-Date).Ticks).json"
        Write-ConfigJson $customSpecs $source $h1m $h2m $h3m $cfgPath
        Log "wrote custom config: $cfgPath"
        $exeArgs = @("format", $InputDocx, "-o", $out, "--config", $cfgPath)
        if ($h1m) { $exeArgs += @("--h1-marker", $h1m) }
        if ($h2m) { $exeArgs += @("--h2-marker", $h2m) }
        if ($h3m) { $exeArgs += @("--h3-marker", $h3m) }
    }

    Log "about to invoke exe: '$exe' $($exeArgs -join ' | ')"
    $exeOutput = & $exe @exeArgs 2>&1 | Out-String
    $exitCode = $LASTEXITCODE

    Log "exe exit=$exitCode"
    foreach ($line in ($exeOutput -split "`r?`n")) {
        if ($null -ne $line -and $line.Length -gt 0) { Log "  | $line" }
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
