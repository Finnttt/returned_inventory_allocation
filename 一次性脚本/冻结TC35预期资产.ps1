# 一次性冻结 TC-35 预期资产：预期_汇总表 2 行 + 预期_成功分配明细 9 行 + 预期_断言 5 行。
# 内容来自 TC-35 文档 §4，已与 2026-07-19 实际运行捕获逐行一致。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm"
try {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
} catch { throw "工作簿正在使用中，请先关闭：$path" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = "D:\cursor_practice\测试用例部分汇总_TC35预期前_$stamp.xlsm"
Copy-Item -LiteralPath $path -Destination $backup -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        $cell = $ws.Cells($row, $c + 1)
        if ($null -eq $v -or $v -eq "") { $cell.NumberFormat = "@"; $cell.Value2 = "" }
        elseif ($v -is [ValueType]) { $cell.Value2 = [double]$v }
        else { $cell.NumberFormat = "@"; $cell.Value2 = [string]$v }
    }
}

$B = "BATCH-TC35"
$S = "SF3190000000035"
try {
    $wb = $excel.Workbooks.Open($path, 0, $false)

    $ws = $wb.Worksheets.Item("预期_汇总表")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    $rows = @(
        @($B, $S, "TK10000351", "批量导入", ""),
        @($B, $S, "TK10000352", "手工操作", "")
    )
    foreach ($r in $rows) { Set-Row $ws $next $r; $next++ }
    Write-Output ("预期汇总 +" + $rows.Count)

    $ws = $wb.Worksheets.Item("预期_成功分配明细")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    $rows = @(
        @($B, $S, "TK10000351", "H00000035", "00001", 15, "QC", "L03", "2029/01/01", 15, "批量导入", "批量导入"),
        @($B, $S, "TK10000351", "H00000035", "00002", 6, "QC", "L04", "2029/01/01", 6, "批量导入", "批量导入"),
        @($B, $S, "TK10000351", "H00000035", "00003", 6, "QC", "L05", "2029/01/01", 6, "批量导入", "批量导入"),
        @($B, $S, "TK10000351", "H00000035", "00004", 6, "ZP", "L02", "2029/01/01", 6, "批量导入", "批量导入"),
        @($B, $S, "TK10000352", "H00000035", "00001", 6, "ZP", "L01", "2029/01/01", 6, "批量导入", "手工操作"),
        @($B, $S, "TK10000352", "H00000035", "00002", 6, "ZP", "L01", "2029/01/01", 6, "批量导入", "手工操作"),
        @($B, $S, "TK10000352", "H00000035", "00003", 6, "ZP", "L01", "2029/01/01", 3, "手工操作", "手工操作"),
        @($B, $S, "TK10000352", "H00000035", "00003", 6, "ZP", "L02", "2029/01/01", 3, "手工操作", "手工操作"),
        @($B, $S, "TK10000352", "H00000036", "00004", 2, "NG", "L01", "2029/01/01", 2, "批量导入", "手工操作")
    )
    foreach ($r in $rows) { Set-Row $ws $next $r; $next++ }
    Write-Output ("预期明细 +" + $rows.Count)

    $ws = $wb.Worksheets.Item("预期_断言")
    $next = $ws.Cells($ws.Rows.Count, 3).End(-4162).Row + 1
    $rows = @(
        @($B, $S, "汇总表行数", "2"),
        @($B, $S, "明细表行数", "9"),
        @($B, $S, "异常表行数", "0"),
        @($B, $S, "运行历史.分配成功数", "1"),
        @($B, $S, "运行历史.总回溯次数", "4")
    )
    foreach ($r in $rows) { Set-Row $ws $next $r; $next++ }
    Write-Output ("预期断言 +" + $rows.Count)

    $wb.Save()
    $wb.Close($false)
    Write-Output "BACKUP=$backup"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
