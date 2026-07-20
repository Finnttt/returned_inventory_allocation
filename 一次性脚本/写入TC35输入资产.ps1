# 一次性写入 TC-35 输入侧资产（退单/库存/配置 8+6+1 行 + BATCH-TC35 计划）。
# 预期表先不写：先跑系统捕获实际输出，与 TC-35 文档 §3 的手工推导比对一致后再补预期。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm"
try {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
} catch { throw "工作簿正在使用中，请先关闭：$path" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = "D:\cursor_practice\测试用例部分汇总_TC35输入前_$stamp.xlsm"
Copy-Item -LiteralPath $path -Destination $backup -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        $cell = $ws.Cells($row, $c + 1)
        if ($null -eq $v) { $cell.Value2 = "" }
        elseif ($v -is [ValueType]) { $cell.Value2 = [double]$v }
        else { $cell.NumberFormat = "@"; $cell.Value2 = [string]$v }
    }
}

$ship = "SF3190000000035"
try {
    $wb = $excel.Workbooks.Open($path, 0, $false)

    $ret = @(
        @($ship, "TK10000351", "H00000035", "00001", 15),
        @($ship, "TK10000351", "H00000035", "00002", 6),
        @($ship, "TK10000351", "H00000035", "00003", 6),
        @($ship, "TK10000351", "H00000035", "00004", 6),
        @($ship, "TK10000352", "H00000035", "00001", 6),
        @($ship, "TK10000352", "H00000035", "00002", 6),
        @($ship, "TK10000352", "H00000035", "00003", 6),
        @($ship, "TK10000352", "H00000036", "00004", 2)
    )
    $ws = $wb.Worksheets.Item("输入_退单表")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    foreach ($r in $ret) { Set-Row $ws $next $r; $next++ }
    Write-Output ("退单 +" + $ret.Count)

    $inv = @(
        @($ship, "H00000035", "ZP", "L01", "2029/01/01", 15),
        @($ship, "H00000035", "ZP", "L02", "2029/01/01", 9),
        @($ship, "H00000035", "QC", "L03", "2029/01/01", 15),
        @($ship, "H00000035", "QC", "L04", "2029/01/01", 6),
        @($ship, "H00000035", "QC", "L05", "2029/01/01", 6),
        @($ship, "H00000036", "NG", "L01", "2029/01/01", 2)
    )
    $ws = $wb.Worksheets.Item("输入_质检库存表")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    foreach ($r in $inv) { Set-Row $ws $next $r; $next++ }
    Write-Output ("库存 +" + $inv.Count)

    $ws = $wb.Worksheets.Item("输入_配置")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    Set-Row $ws $next @($ship, "TC-35", 200, "关闭", "不敏感", "2099/01/01", "大小批量混合含回溯")
    Write-Output "配置 +1"

    $ws = $wb.Worksheets.Item("批量测试计划")
    $next = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
    Set-Row $ws $next @("BATCH-TC35", "是", "FullRun", $ship, "按物流单号读取", "", "关闭", "TC-35 大小批量混合含回溯")
    Write-Output "计划 +BATCH-TC35"

    $wb.Save()
    $wb.Close($false)
    Write-Output "BACKUP=$backup"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
