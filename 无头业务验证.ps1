# 无头业务验证：对生产工作簿《退货入库分配系统.xlsm》做对抗性端到端审查。
# 每个场景在临时副本上运行，不改动正式文件；全部通过时退出码为 0。
#
# 场景清单：
#   S1 正常业务流（SF0000：干跑无异常，分配结果与预期一致，历史追加 2 行）
#   S2 重跑覆盖（第二次分配覆盖输出而不是追加，历史 3 行）
#   S3 E01 拦截（行号数值型 → 无法分配 + 异常明细，无分配明细）
#   S4 空输入（不崩溃，输出为空，历史仍追加）
#   S5 调试日志（级别=详细 → 调试日志表 19 列且有数据行）
#   S6 性能基线（500 物流单号 / 1500 退单行 / 2500 库存行，全部批量导入）
#   S7 E02 原始值前导零（行号 "00001" 重复 → 异常明细原始值按文本保留 "00001"）
#   S8 E06 进异常明细（物流单号仅在退单表 → 汇总与异常明细同时出现 E06）
#   （非法配置场景由 RunSingleTest 17 的 TC-54a/b 覆盖；COM 触发未捕获 VBA 错误会挂死无头进程）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File .\无头业务验证.ps1

param(
    [string]$ProdPath = (Join-Path $PSScriptRoot "退货入库分配系统.xlsm")
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ProdPath = [System.IO.Path]::GetFullPath($ProdPath)

if (-not (Test-Path -LiteralPath $ProdPath)) { throw "生产工作簿不存在：$ProdPath" }
try {
    $stream = [System.IO.File]::Open($ProdPath, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
} catch { throw "生产工作簿正在使用中，请先关闭后重试：$ProdPath" }

$tempDir = Join-Path $env:TEMP ("ria_e2e_" + [System.Guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$prevSec = $excel.AutomationSecurity
$excel.AutomationSecurity = 1

$script:passCount = 0
$script:failCount = 0

function Report-Case([string]$id, [bool]$ok, [string]$detail) {
    if ($ok) { $script:passCount++; Write-Output "PASS $id $detail" }
    else { $script:failCount++; Write-Output "FAIL $id $detail" }
}

function New-CaseWorkbook([string]$caseName) {
    $copy = Join-Path $tempDir ($caseName + ".xlsm")
    Copy-Item -LiteralPath $ProdPath -Destination $copy -Force
    return $excel.Workbooks.Open($copy, 0, $false)
}

function Run-Silent($wb, [string]$macro) {
    $excel.Run("'" + $wb.Name + "'!" + $macro, $wb) | Out-Null
}

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        # PowerShell 对 [object] 装箱的 Int32 走 COM Value2 赋值会误转 String，
        # 这里按值类型显式还原：数值写 Double，其余写 String，空值写空串。
        if ($null -eq $v) { $ws.Cells($row, $c + 1).Value2 = "" }
        elseif ($v -is [ValueType]) { $ws.Cells($row, $c + 1).Value2 = [double]$v }
        else { $ws.Cells($row, $c + 1).Value2 = [string]$v }
    }
}

function Get-DataRows($wb, [string]$sheetName) {
    $ws = $wb.Worksheets.Item($sheetName)
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row  # xlUp
    $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column
    $rows = @()
    for ($r = 2; $r -le $lastRow; $r++) {
        $vals = @()
        for ($c = 1; $c -le $lastCol; $c++) { $vals += [string]$ws.Cells($r, $c).Text }
        $rows += , ($vals -join "|")
    }
    # 单元素数组在 return 时会被 PowerShell 展开成字符串，必须包一层防展开。
    return , $rows
}

function Add-Sf0000Inputs($wb) {
    Set-Row $wb.Worksheets.Item("输入_退单表") 2 @("SF3190000000000", "TK00000001", "H000000001", "00001", 2)
    Set-Row $wb.Worksheets.Item("输入_质检库存表") 2 @("SF3190000000000", "H000000001", "ZP", "LA01", "2029/01/01", 2)
}

$failNote = ""
try {
    # ---------- S1 正常业务流 ----------
    $wb = New-CaseWorkbook "s1"
    Add-Sf0000Inputs $wb
    Run-Silent $wb "RunValidationOnlySilent"
    $anomalies = Get-DataRows $wb "数据异常明细表"
    Run-Silent $wb "RunFullAllocationSilent"
    $summary = Get-DataRows $wb "分配状态汇总表"
    $detail = Get-DataRows $wb "成功分配明细表"
    $history = Get-DataRows $wb "运行历史记录表"
    $ok = ($anomalies.Count -eq 0) -and ($summary.Count -eq 1) -and ($detail.Count -eq 1) -and ($history.Count -eq 2)
    if ($ok) {
        $s = $summary[0]
        $ok = $s -like "SF3190000000000|TK00000001|批量导入|*"
    }
    if ($ok) {
        $d = $detail[0]
        $ok = $d -eq "SF3190000000000|TK00000001|H000000001|00001|2|ZP|LA01|2029/01/01|2|批量导入|批量导入"
    }
    if ($ok) {
        $ok = ($history[0] -match "\|Dry Run\|") -and ($history[1] -match "\|Full Run\|")
    }
    Report-Case "S1" $ok ("汇总=" + ($summary -join ";") + " 明细=" + ($detail -join ";") + " 历史行=" + $history.Count + " 异常行=" + $anomalies.Count)

    # ---------- S2 重跑覆盖（沿用 S1 工作簿再跑一次完整分配） ----------
    Run-Silent $wb "RunFullAllocationSilent"
    $summary2 = Get-DataRows $wb "分配状态汇总表"
    $detail2 = Get-DataRows $wb "成功分配明细表"
    $history2 = Get-DataRows $wb "运行历史记录表"
    $ok = ($summary2.Count -eq 1) -and ($detail2.Count -eq 1) -and ($history2.Count -eq 3) `
        -and ($summary2[0] -eq $summary[0]) -and ($detail2[0] -eq $detail[0])
    Report-Case "S2" $ok ("重跑后 汇总行=" + $summary2.Count + " 明细行=" + $detail2.Count + " 历史行=" + $history2.Count)
    $wb.Close($false)

    # ---------- S3 E01 拦截（行号数值型） ----------
    $wb = New-CaseWorkbook "s3"
    Set-Row $wb.Worksheets.Item("输入_退单表") 2 @("SF3190000099001", "TK00009901", "H000000001", 1, 2)
    Set-Row $wb.Worksheets.Item("输入_质检库存表") 2 @("SF3190000099001", "H000000001", "ZP", "LA01", "2029/01/01", 2)
    Run-Silent $wb "RunFullAllocationSilent"
    $summary = Get-DataRows $wb "分配状态汇总表"
    $detail = Get-DataRows $wb "成功分配明细表"
    $anomalies = Get-DataRows $wb "数据异常明细表"
    $hasE01 = $false; foreach ($a in $anomalies) { if ($a -match "E01") { $hasE01 = $true } }
    $blocked = $false; foreach ($s in $summary) { if ($s -match "无法分配") { $blocked = $true } }
    $ok = ($detail.Count -eq 0) -and $hasE01 -and $blocked
    Report-Case "S3" $ok ("汇总=" + ($summary -join ";") + " 异常=" + ($anomalies -join ";"))
    $wb.Close($false)

    # ---------- S4 空输入 ----------
    $wb = New-CaseWorkbook "s4"
    $threw = $false
    try { Run-Silent $wb "RunFullAllocationSilent" } catch { $threw = $true; $failNote = $_.Exception.Message }
    $summary = Get-DataRows $wb "分配状态汇总表"
    $history = Get-DataRows $wb "运行历史记录表"
    $ok = (-not $threw) -and ($summary.Count -eq 0) -and ($history.Count -eq 1)
    Report-Case "S4" $ok ("异常抛出=" + $threw + " " + $failNote + " 汇总行=" + $summary.Count + " 历史行=" + $history.Count)
    $wb.Close($false)

    # ---------- S5 调试日志（详细级别） ----------
    # 注：非法配置（E13/配置读取错误）的覆盖由 RunSingleTest 17 的 TC-54a/b 承担；
    # 经 COM 触发未捕获 VBA 错误会弹“Microsoft Visual Basic”对话框挂死无头进程，不适合在此断言。
    $wb = New-CaseWorkbook "s5"
    Add-Sf0000Inputs $wb
    $configWs = $wb.Worksheets.Item("输入_配置")
    $lastCfgRow = $configWs.Cells($configWs.Rows.Count, 1).End(-4162).Row
    for ($r = 2; $r -le $lastCfgRow; $r++) {
        if ([string]$configWs.Cells($r, 1).Value2 -eq "调试日志级别") { $configWs.Cells($r, 2).Value2 = "详细" }
    }
    Run-Silent $wb "RunFullAllocationSilent"
    $debugWs = $wb.Worksheets.Item("调试日志")
    $headerCols = $debugWs.Cells(1, $debugWs.Columns.Count).End(-4159).Column
    $debugRows = (Get-DataRows $wb "调试日志").Count
    $ok = ($headerCols -eq 19) -and ($debugRows -ge 1)
    Report-Case "S5" $ok ("日志列数=" + $headerCols + " 日志行=" + $debugRows)
    $wb.Close($false)

    # ---------- S6 性能基线（500 单 / 1500 退单行 / 1500 库存行） ----------
    # 库存按 SKU 与退单量逐 SKU 守恒（E08）：A=2、B=3、C=1 各一行 ZP，保证 100% 策略一批量导入。
    $wb = New-CaseWorkbook "s6"
    $returnWs = $wb.Worksheets.Item("输入_退单表")
    $invWs = $wb.Worksheets.Item("输入_质检库存表")
    $n = 500
    $rr = 2; $ir = 2
    for ($i = 1; $i -le $n; $i++) {
        $ship = "SF3190001{0:D6}" -f $i
        $wms = "TK{0:D8}" -f $i
        Set-Row $returnWs $rr   @($ship, $wms, "H0000000A", "00001", 2); $rr++
        Set-Row $returnWs $rr   @($ship, $wms, "H0000000B", "00002", 3); $rr++
        Set-Row $returnWs $rr   @($ship, $wms, "H0000000C", "00003", 1); $rr++
        Set-Row $invWs $ir @($ship, "H0000000A", "ZP", "L01", "2029/01/01", 2); $ir++
        Set-Row $invWs $ir @($ship, "H0000000B", "ZP", "L01", "2029/01/01", 3); $ir++
        Set-Row $invWs $ir @($ship, "H0000000C", "ZP", "L01", "2029/01/01", 1); $ir++
    }
    $t0 = Get-Date
    Run-Silent $wb "RunFullAllocationSilent"
    $elapsed = [int]((Get-Date) - $t0).TotalSeconds
    $summary = Get-DataRows $wb "分配状态汇总表"
    $detail = Get-DataRows $wb "成功分配明细表"
    $allOk = $true; foreach ($s in $summary) { if ($s -notmatch "批量导入") { $allOk = $false } }
    $ok = ($summary.Count -eq $n) -and ($detail.Count -eq (3 * $n)) -and $allOk
    Report-Case "S6" $ok ("耗时=" + $elapsed + "s 汇总行=" + $summary.Count + " 明细行=" + $detail.Count)
    $wb.Close($false)

    # ---------- S7 E02 原始值保留前导零（2026-07-19 用户报告缺陷） ----------
    $wb = New-CaseWorkbook "s7"
    $wsR = $wb.Worksheets.Item("输入_退单表")
    Set-Row $wsR 2 @("SF3190000099002", "TK00009902", "H000000001", "00001", 1)
    Set-Row $wsR 3 @("SF3190000099002", "TK00009902", "H000000001", "00001", 1)
    Set-Row $wb.Worksheets.Item("输入_质检库存表") 2 @("SF3190000099002", "H000000001", "ZP", "LA01", "2029/01/01", 2)
    Run-Silent $wb "RunFullAllocationSilent"
    $wsA = $wb.Worksheets.Item("数据异常明细表")
    $anomalies = Get-DataRows $wb "数据异常明细表"
    $rawCell = $wsA.Cells(2, 7)
    $ok = ($anomalies.Count -eq 2) -and ($rawCell.Text -eq "00001") -and ($rawCell.Value2.GetType().Name -eq "String")
    Report-Case "S7" $ok ("原始值=" + $rawCell.Text + " 类型=" + $rawCell.Value2.GetType().Name + " 异常行=" + $anomalies.Count)
    $wb.Close($false)

    # ---------- S8 E06 进入异常明细表（2026-07-20 规则变更） ----------
    $wb = New-CaseWorkbook "s8"
    Set-Row $wb.Worksheets.Item("输入_退单表") 2 @("SF3190000099003", "TK00009903", "H000000001", "00001", 5)
    # 库存表留空：该物流单号仅在退单表 → E06
    Run-Silent $wb "RunFullAllocationSilent"
    $summary = Get-DataRows $wb "分配状态汇总表"
    $wsA = $wb.Worksheets.Item("数据异常明细表")
    $anomalies = Get-DataRows $wb "数据异常明细表"
    $ok = ($summary.Count -eq 1) -and ($summary[0] -match "无法分配") -and ($summary[0] -match "E06") `
        -and ($anomalies.Count -eq 1) `
        -and ([string]$wsA.Cells(2, 1).Text -eq "退单表") `
        -and ([string]$wsA.Cells(2, 3).Text -eq "SF3190000099003") `
        -and ([string]$wsA.Cells(2, 6).Text -eq "物流单号") `
        -and ([string]$wsA.Cells(2, 7).Text -eq "SF3190000099003") `
        -and ([string]$wsA.Cells(2, 8).Text -eq "E06")
    Report-Case "S8" $ok ("汇总=" + ($summary -join ";") + " 异常=" + ($anomalies -join ";"))
    $wb.Close($false)

    Write-Output ("VERDICT=" + $(if ($script:failCount -eq 0) { "PASS" } else { "FAIL" }) +
        "（通过=" + $script:passCount + "，失败=" + $script:failCount + "）")
} finally {
    try { $excel.AutomationSecurity = $prevSec } catch {}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    # 两次 GC 确保工作簿/工作表 RCW 终结，避免隐藏 EXCEL.EXE 进程残留
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    if (Test-Path -LiteralPath $tempDir) { Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($script:failCount -gt 0) { exit 1 }
