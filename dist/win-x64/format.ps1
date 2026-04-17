# format.ps1 -- gongwen-paiban interactive launcher
# Saved as UTF-8 with BOM. Windows PowerShell 5.1 requires the BOM to
# correctly read non-ASCII characters in the script.

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$InputDocx
)

$ErrorActionPreference = 'Stop'

# --- basic paths (safe computation) -------------------------------------
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) {
    try { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path } catch {}
}
if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = (Get-Location).Path }

$exe     = Join-Path $scriptDir "gongwen-paiban.exe"
$logFile = Join-Path $scriptDir "paiban-log.txt"

# Make Add-Content-emitted bytes match what bat wrote (UTF-8).
# PS 5.1's "UTF8" encoding writes with BOM on first write, but since
# bat already created the file, Add-Content just appends raw UTF-8 bytes.
$PSDefaultParameterValues = @{
    'Add-Content:Encoding' = 'UTF8'
}

function Log([string]$msg) {
    $line = "[ps1 $(Get-Date -Format 'HH:mm:ss')] $msg"
    # Try the log file; if it somehow fails, also emit to console so bat can still capture.
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 }
    catch {
        try { [Console]::Out.WriteLine($line) } catch {}
    }
}

function Safe-ShowError([string]$msg) {
    Log "ERROR: $msg"
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show($msg, "公文排版", 'OK', 'Error')
    } catch {
        Log "Safe-ShowError failed to show MsgBox: $($_.Exception.Message)"
    }
}

try {
    Log "format.ps1 start.  PS version=$($PSVersionTable.PSVersion)"
    Log "scriptDir='$scriptDir'"
    Log "InputDocx='$InputDocx'"
    Log "exe='$exe'"

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Log "WinForms loaded"

    if ([string]::IsNullOrEmpty($InputDocx)) {
        Safe-ShowError "ps1 未收到输入文件。应由 format.bat 以 -InputDocx 形式传入。"
        exit 1
    }
    if (-not (Test-Path -LiteralPath $InputDocx)) {
        Safe-ShowError "输入文件不存在：`n$InputDocx"
        exit 1
    }
    $ext = [System.IO.Path]::GetExtension($InputDocx)
    if ($ext.ToLower() -ne ".docx") {
        Safe-ShowError "只支持 .docx 格式。当前：$ext`n如果是 .doc，请先用 Word/WPS 另存为 .docx。"
        exit 1
    }
    if (-not (Test-Path -LiteralPath $exe)) {
        Safe-ShowError "找不到 gongwen-paiban.exe。`n应和 format.bat / format.ps1 放在同一目录。`n期望路径：$exe"
        exit 1
    }

    # ==================== Step 1: source ====================
    Log "step 1 start (source picker)"
    $srcForm = New-Object System.Windows.Forms.Form
    $srcForm.Text = "公文排版 - 第 1 步：文档来源"
    $srcForm.Size = New-Object System.Drawing.Size(440, 220)
    $srcForm.StartPosition = "CenterScreen"
    $srcForm.FormBorderStyle = "FixedDialog"
    $srcForm.MaximizeBox = $false
    $srcForm.MinimizeBox = $false
    $srcForm.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "这份文档是用什么软件做的？"
    $lbl.Location = New-Object System.Drawing.Point(20, 25)
    $lbl.Size = New-Object System.Drawing.Size(400, 35)
    $lbl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $srcForm.Controls.Add($lbl)

    $script:sourceChoice = $null

    $btnWps = New-Object System.Windows.Forms.Button
    $btnWps.Text = "WPS"
    $btnWps.Location = New-Object System.Drawing.Point(20, 85)
    $btnWps.Size = New-Object System.Drawing.Size(115, 50)
    $btnWps.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnWps.Add_Click({ $script:sourceChoice = "wps"; $srcForm.Close() })

    $btnOffice = New-Object System.Windows.Forms.Button
    $btnOffice.Text = "Microsoft Office"
    $btnOffice.Location = New-Object System.Drawing.Point(150, 85)
    $btnOffice.Size = New-Object System.Drawing.Size(140, 50)
    $btnOffice.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnOffice.Add_Click({ $script:sourceChoice = "office"; $srcForm.Close() })

    $btnAuto = New-Object System.Windows.Forms.Button
    $btnAuto.Text = "不确定"
    $btnAuto.Location = New-Object System.Drawing.Point(305, 85)
    $btnAuto.Size = New-Object System.Drawing.Size(105, 50)
    $btnAuto.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnAuto.Add_Click({ $script:sourceChoice = "auto"; $srcForm.Close() })

    $srcForm.Controls.AddRange(@($btnWps, $btnOffice, $btnAuto))
    [void]$srcForm.ShowDialog()

    if ([string]::IsNullOrEmpty($script:sourceChoice)) {
        Log "step 1 cancelled by user"
        exit 1
    }
    $source = $script:sourceChoice
    Log "step 1 done, source='$source'"

    # ==================== Step 2: know markers ====================
    Log "step 2 start"
    $msg = @"
您知道这份文档的一级/二级/三级标题前的编号格式吗？

・「是」：自己指定各级的编号样例，程序按您的意图精准识别
・「否」：让程序按通用规则自动判断
"@
    $knowResult = [System.Windows.Forms.MessageBox]::Show($msg, "公文排版 - 第 2 步：编号方式", 'YesNoCancel', 'Question')
    if ($knowResult -eq 'Cancel') { Log "step 2 cancelled"; exit 1 }
    Log "step 2 answer=$knowResult"

    $h1m = ""; $h2m = ""; $h3m = ""

    if ($knowResult -eq 'Yes') {
        Log "step 3 start (marker inputs)"
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "公文排版 - 第 3 步：指定编号样例"
        $inputForm.Size = New-Object System.Drawing.Size(560, 400)
        $inputForm.StartPosition = "CenterScreen"
        $inputForm.FormBorderStyle = "FixedDialog"
        $inputForm.MaximizeBox = $false
        $inputForm.MinimizeBox = $false
        $inputForm.TopMost = $true

        $tip = New-Object System.Windows.Forms.Label
        $tip.Text = "只需填一个样例，程序智能补齐。留空 = 该级交给自动判断。"
        $tip.Location = New-Object System.Drawing.Point(20, 15)
        $tip.Size = New-Object System.Drawing.Size(520, 30)
        $tip.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $tip.ForeColor = [System.Drawing.Color]::DimGray
        $inputForm.Controls.Add($tip)

        $tbl = @{}
        $i = 0
        foreach ($row in @(
            @{ k='H1'; label='一级标题：'; hint='例：一、  /  1.  /  第一章' },
            @{ k='H2'; label='二级标题：'; hint='例：（一）  /  1.1  /  第一节' },
            @{ k='H3'; label='三级标题：'; hint='例：1.  /  (1)  /  1.1.1' }
        )) {
            $y = 60 + $i * 42
            $lbl2 = New-Object System.Windows.Forms.Label
            $lbl2.Text = $row.label
            $lbl2.Location = New-Object System.Drawing.Point(20, $y)
            $lbl2.Size = New-Object System.Drawing.Size(100, 24)
            $lbl2.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Location = New-Object System.Drawing.Point(125, ($y - 2))
            $tb.Size = New-Object System.Drawing.Size(160, 28)
            $tb.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

            $hintLbl = New-Object System.Windows.Forms.Label
            $hintLbl.Text = $row.hint
            $hintLbl.Location = New-Object System.Drawing.Point(295, $y)
            $hintLbl.Size = New-Object System.Drawing.Size(245, 24)
            $hintLbl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
            $hintLbl.ForeColor = [System.Drawing.Color]::DimGray

            $inputForm.Controls.AddRange(@($lbl2, $tb, $hintLbl))
            $tbl[$row.k] = $tb
            $i++
        }

        $note = New-Object System.Windows.Forms.Label
        $note.Text = "容错提示：" + [Environment]::NewLine +
                     "・括号中英文宽度均可，开闭括号可错配" + [Environment]::NewLine +
                     "・标点后缀可省" + [Environment]::NewLine +
                     "・分级不会错吃：1. 不会匹配 1.1；1.1 不会匹配 1.1.1"
        $note.Location = New-Object System.Drawing.Point(20, 195)
        $note.Size = New-Object System.Drawing.Size(520, 110)
        $note.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $note.ForeColor = [System.Drawing.Color]::DimGray
        $inputForm.Controls.Add($note)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "确定"
        $btnOK.Location = New-Object System.Drawing.Point(330, 315)
        $btnOK.Size = New-Object System.Drawing.Size(95, 34)
        $btnOK.DialogResult = 'OK'
        $inputForm.AcceptButton = $btnOK

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "取消"
        $btnCancel.Location = New-Object System.Drawing.Point(435, 315)
        $btnCancel.Size = New-Object System.Drawing.Size(95, 34)
        $btnCancel.DialogResult = 'Cancel'
        $inputForm.CancelButton = $btnCancel

        $inputForm.Controls.AddRange(@($btnOK, $btnCancel))
        $r = $inputForm.ShowDialog()
        if ($r -eq 'Cancel') { Log "step 3 cancelled"; exit 1 }

        $h1m = $tbl['H1'].Text.Trim()
        $h2m = $tbl['H2'].Text.Trim()
        $h3m = $tbl['H3'].Text.Trim()
        Log "step 3 done; H1='$h1m' H2='$h2m' H3='$h3m'"
    } else {
        Log "user chose auto-detection"
    }

    # ==================== Run exe ====================
    $dir  = [System.IO.Path]::GetDirectoryName($InputDocx)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputDocx)
    $out  = Join-Path $dir "${name}_formatted.docx"

    $exeArgs = @("format", $InputDocx, "-o", $out, "--source", $source)
    if ($h1m) { $exeArgs += @("--h1-marker", $h1m) }
    if ($h2m) { $exeArgs += @("--h2-marker", $h2m) }
    if ($h3m) { $exeArgs += @("--h3-marker", $h3m) }

    Log "about to invoke exe: '$exe' $($exeArgs -join ' | ')"

    # Use the call operator (&) with splatting — works on both PS 5.1 (.NET Fx 4.x)
    # and PS 7 (.NET 5+). Capturing both streams via 2>&1 into one string.
    $exeOutput = & $exe @exeArgs 2>&1 | Out-String
    $exitCode = $LASTEXITCODE

    Log "exe exit code = $exitCode"
    Log "exe output (merged stdout+stderr) -----"
    foreach ($line in ($exeOutput -split "`r?`n")) {
        if ($null -ne $line -and $line.Length -gt 0) { Log "  | $line" }
    }
    Log "---- end exe output"

    if ($exitCode -ne 0) {
        Safe-ShowError "排版失败（退出码 $exitCode）。`n`n$exeOutput`n`n详情见 paiban-log.txt"
        exit $exitCode
    }

    $done = [System.Windows.Forms.MessageBox]::Show(
        "完成！`n`n输出：$out`n`n要打开所在文件夹吗？",
        "公文排版",
        'YesNo', 'Information')
    if ($done -eq 'Yes') {
        Start-Process "explorer.exe" -ArgumentList "/select,`"$out`""
    }
    Log "all done, exit 0"
    exit 0

} catch {
    $err = $_
    $errStr = $err | Out-String
    Log "UNHANDLED PS EXCEPTION at line $($err.InvocationInfo.ScriptLineNumber)"
    Log $errStr
    try {
        $briefMsg = "ps1 脚本发生异常（行 $($err.InvocationInfo.ScriptLineNumber)）：`n`n" +
                    $err.Exception.Message +
                    "`n`n详情见 paiban-log.txt"
        Safe-ShowError $briefMsg
    } catch {}
    exit 99
}
