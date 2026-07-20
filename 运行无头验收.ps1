# 无头验收：通过 Excel COM 在隐藏实例中运行全部测试入口，无需人工打开 VBA 编辑器。
# 覆盖 README“如何运行测试”中的四个入口：RunAllTests、RunSingleTest 17、RunSingleTest 18、RunBatchTestPlan。
# 运行前自动检查文件占用并在上级目录留备份；默认不保存工作簿（结果只读到控制台），加 -SaveResults 才落盘。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File .\运行无头验收.ps1
#   powershell -ExecutionPolicy Bypass -File .\运行无头验收.ps1 -SaveResults

param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "测试用例部分汇总.xlsm"),
    [switch]$SaveResults
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$WorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "测试汇总工作簿不存在：$WorkbookPath"
}

# 占用检测：目标工作簿若已被打开则中止，避免两实例互相覆盖。
try {
    $stream = [System.IO.File]::Open(
        $WorkbookPath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None)
    $stream.Close()
} catch {
    throw "测试汇总工作簿正在使用中，请先关闭后重试：$WorkbookPath"
}

# 备份到项目上级目录，与 同步测试工作簿VBA.ps1 的备份位置约定一致。
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Split-Path (Split-Path $WorkbookPath -Parent) -Parent
$backupPath = Join-Path $backupDir ("测试用例部分汇总_无头验收前_{0}.xlsm" -f $stamp)
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

# 允许经 COM 打开的工作簿执行宏；结束后恢复原设置。
$prevAutomationSecurity = $excel.AutomationSecurity
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow

$wb = $null
$failures = 0
try {
    $wb = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $macroPrefix = "'" + $wb.Name + "'!"

    Write-Output "== 1/4 RunAllTests（内存单元+端到端） =="
    $r1 = [string]$excel.Run($macroPrefix + "RunAllTestsSilent")
    Write-Output $r1
    if ($r1 -notmatch "失败=0") {
        $failures++
        $log1 = [string]$excel.Run($macroPrefix + "GetFailLog")
        if ($log1) { Write-Output ("失败明细：`n" + $log1) }
    }

    Write-Output "== 2/4 RunSingleTest 17（文件集成） =="
    $r2 = [string]$excel.Run($macroPrefix + "RunSingleTestSilent", 17)
    Write-Output $r2
    if ($r2 -notmatch "失败=0") {
        $failures++
        $log2 = [string]$excel.Run($macroPrefix + "GetFailLog")
        if ($log2) { Write-Output ("失败明细：`n" + $log2) }
    }

    Write-Output "== 3/4 RunSingleTest 18（批量测试器冒烟） =="
    $r3 = [string]$excel.Run($macroPrefix + "RunSingleTestSilent", 18)
    Write-Output $r3
    if ($r3 -notmatch "失败=0") {
        $failures++
        $log3 = [string]$excel.Run($macroPrefix + "GetFailLog")
        if ($log3) { Write-Output ("失败明细：`n" + $log3) }
    }

    Write-Output "== 4/4 RunBatchTestPlan（Excel 批量回归） =="

    # 防呆：没有启用批次时 RunBatchTestPlan 静默返回，结果表仍是上次运行的残留。
    # 先确认存在启用批次，再用“开始时间 >= 本次调用时刻”过滤出真正的新结果。
    $planWs = $wb.Worksheets.Item("批量测试计划")
    $planLast = $planWs.Cells($planWs.Rows.Count, 1).End(-4162).Row
    $enabledPlans = @()
    for ($row = 2; $row -le $planLast; $row++) {
        if ([string]$planWs.Cells($row, 2).Value2 -eq "是") {
            $enabledPlans += [string]$planWs.Cells($row, 1).Value2
        }
    }
    if ($enabledPlans.Count -eq 0) {
        Write-Output "批量回归：批量测试计划中没有启用=是的批次，未实际执行。"
        $failures++
    } else {
        Write-Output ("启用批次：" + ($enabledPlans -join "、"))
        $batchStart = Get-Date
        $excel.Run($macroPrefix + "RunBatchTestPlan", $wb, $false) | Out-Null

        $resultWs = $wb.Worksheets.Item("批量测试结果")
        $lastRow = $resultWs.Cells($resultWs.Rows.Count, 1).End(-4162).Row  # xlUp
        $totalPass = 0
        $totalFail = 0
        $freshRows = 0
        $badRows = @()
        for ($row = 2; $row -le $lastRow; $row++) {
            $rawStart = $resultWs.Cells($row, 7).Value2
            $startedAt = [DateTime]::MinValue
            if ($rawStart -is [DateTime]) { $startedAt = $rawStart }
            elseif ($null -ne $rawStart) { try { $startedAt = [DateTime]::FromOADate([double]$rawStart) } catch {} }
            if ($startedAt -lt $batchStart.AddSeconds(-2)) { continue }  # 历史残留行，不计入本次验收
            $freshRows++
            $status = [string]$resultWs.Cells($row, 6).Value2
            $pass = [int]($resultWs.Cells($row, 9).Value2)
            $fail = [int]($resultWs.Cells($row, 10).Value2)
            $totalPass += $pass
            $totalFail += $fail
            if ($status -ne "成功" -and $status -ne "期望错误通过") {
                $badRows += ("行{0}: 批次={1} 子序号={2} 状态={3} 备注={4}" -f `
                    $row, $resultWs.Cells($row, 1).Value2, $resultWs.Cells($row, 2).Value2, `
                    $status, $resultWs.Cells($row, 11).Value2)
            }
        }
        Write-Output ("批量回归：本次子批次={0}，断言通过={1}，断言失败={2}" -f $freshRows, $totalPass, $totalFail)
        if ($freshRows -eq 0) {
            Write-Output "批量回归：结果表没有本次运行的新行，判定为未实际执行。"
            $failures++
        }
        if ($badRows.Count -gt 0) {
            Write-Output "异常子批次："
            $badRows | ForEach-Object { Write-Output $_ }
            $failures++
        }
        if ($totalFail -gt 0) { $failures++ }
    }

    if ($SaveResults) {
        $wb.Save()
    }
    $wb.Close($SaveResults -eq $true)
    $wb = $null

    Write-Output "BACKUP=$backupPath"
    if ($failures -eq 0) {
        Write-Output "VERDICT=PASS"
    } else {
        Write-Output ("VERDICT=FAIL（{0} 个阶段未通过）" -f $failures)
    }
} finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    try { $excel.AutomationSecurity = $prevAutomationSecurity } catch {}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    # 两次 GC 确保工作表等 RCW 终结，避免隐藏 EXCEL.EXE 进程残留
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}

if ($failures -gt 0) { exit 1 }
