# 一次性补充 TC-45 自动化回归资产（数据层）：
#   批量测试计划 追加 BATCH-TC45A / BATCH-TC45B1 / BATCH-TC45B2 三个 DryRun 批次（启用）；
#   预期_断言 追加对应的结构化断言行（依赖 modBatchTestRunner 的扩展断言项）。
# 断言口径来自 TC-45_R128_DryRun测试用例.md §3~§5。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm"
try {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
} catch { throw "工作簿正在使用中，请先关闭：$path" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = "D:\cursor_practice\测试用例部分汇总_TC45资产前_$stamp.xlsm"
Copy-Item -LiteralPath $path -Destination $backup -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path, 0, $false)

    # ── 批量测试计划 ──
    $planWs = $wb.Worksheets.Item("批量测试计划")
    $next = $planWs.Cells($planWs.Rows.Count, 1).End(-4162).Row + 1
    $plans = @(
        @("BATCH-TC45A",  "是", "DryRun", "SF3190000000000", "按物流单号读取", "", "关闭", "TC-45a 干跑全通过：输出全空，历史 Dry Run"),
        @("BATCH-TC45B1", "是", "DryRun", "SF3190000000013", "按物流单号读取", "", "关闭", "TC-45b-1 E01 字段级异常进异常明细"),
        @("BATCH-TC45B2", "是", "DryRun", "SF3190000000036", "按物流单号读取", "", "关闭", "TC-45b-2 E11 只进汇总不进异常明细")
    )
    foreach ($p in $plans) {
        for ($c = 0; $c -lt $p.Count; $c++) {
            $planWs.Cells($next, $c + 1).Value2 = [string]$p[$c]
        }
        Write-Output ("计划 r" + $next + ": " + $p[0])
        $next++
    }

    # ── 预期_断言 ──
    $assertWs = $wb.Worksheets.Item("预期_断言")
    $nextA = $assertWs.Cells($assertWs.Rows.Count, 3).End(-4162).Row + 1
    $asserts = @(
        @("BATCH-TC45A", "SF3190000000000", "汇总表行数", "0"),
        @("BATCH-TC45A", "SF3190000000000", "明细表行数", "0"),
        @("BATCH-TC45A", "SF3190000000000", "异常表行数", "0"),
        @("BATCH-TC45A", "SF3190000000000", "历史行数", "1"),
        @("BATCH-TC45A", "SF3190000000000", "运行历史.退单表行数", "1"),
        @("BATCH-TC45A", "SF3190000000000", "运行历史.库存表行数", "1"),
        @("BATCH-TC45A", "SF3190000000000", "运行历史.校验失败数", "0"),
        @("BATCH-TC45A", "SF3190000000000", "运行历史.总回溯次数", "0"),
        @("BATCH-TC45B1", "SF3190000000013", "汇总表行数", "2"),
        @("BATCH-TC45B1", "SF3190000000013", "明细表行数", "0"),
        @("BATCH-TC45B1", "SF3190000000013", "异常表行数", "2"),
        @("BATCH-TC45B1", "SF3190000000013", "汇总错误码", "E01"),
        @("BATCH-TC45B1", "SF3190000000013", "运行历史.校验失败数", "1"),
        @("BATCH-TC45B2", "SF3190000000036", "汇总表行数", "1"),
        @("BATCH-TC45B2", "SF3190000000036", "明细表行数", "0"),
        @("BATCH-TC45B2", "SF3190000000036", "异常表行数", "0"),
        @("BATCH-TC45B2", "SF3190000000036", "汇总错误码", "E11"),
        @("BATCH-TC45B2", "SF3190000000036", "运行历史.校验失败数", "1")
    )
    foreach ($a in $asserts) {
        for ($c = 0; $c -lt 4; $c++) {
            $assertWs.Cells($nextA, $c + 1).NumberFormat = "@"
            $assertWs.Cells($nextA, $c + 1).Value2 = [string]$a[$c]
        }
        $nextA++
    }
    Write-Output ("断言追加=" + $asserts.Count + " 行")

    $wb.Save()
    $wb.Close($false)
    Write-Output "BACKUP=$backup"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
