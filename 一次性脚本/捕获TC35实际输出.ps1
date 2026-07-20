# 捕获 TC-35 数据集在实际系统中的输出（用于与手工推导比对）
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$prod = "D:\cursor_practice\returned_inventory_allocation\退货入库分配系统.xlsm"
$tmp = Join-Path $env:TEMP ("tc35_capture_" + [System.Guid]::NewGuid().ToString("N") + ".xlsm")
Copy-Item $prod $tmp -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        $cell = $ws.Cells($row, $c + 1)
        if ($null -eq $v) { $cell.Value2 = "" }
        elseif ($v -is [ValueType]) { $cell.Value2 = [double]$v }
        else { $cell.NumberFormat = "@"; $cell.Value2 = [string]$v }
    }
}
function Dump($wb, [string]$name) {
    $ws = $wb.Worksheets.Item($name)
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column
    Write-Output ("== " + $name + " rows=" + ($lastRow - 1))
    for ($r = 2; $r -le $lastRow; $r++) {
        $vals = @()
        for ($c = 1; $c -le $lastCol; $c++) { $vals += [string]$ws.Cells($r, $c).Text }
        Write-Output ($vals -join "|")
    }
}

$ship = "SF3190000000035"
try {
    $wb = $excel.Workbooks.Open($tmp, 0, $false)
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
    for ($i = 0; $i -lt $ret.Count; $i++) { Set-Row $ws ($i + 2) $ret[$i] }
    $inv = @(
        @($ship, "H00000035", "ZP", "L01", "2029/01/01", 15),
        @($ship, "H00000035", "ZP", "L02", "2029/01/01", 9),
        @($ship, "H00000035", "QC", "L03", "2029/01/01", 15),
        @($ship, "H00000035", "QC", "L04", "2029/01/01", 6),
        @($ship, "H00000035", "QC", "L05", "2029/01/01", 6),
        @($ship, "H00000036", "NG", "L01", "2029/01/01", 2)
    )
    $ws = $wb.Worksheets.Item("输入_质检库存表")
    for ($i = 0; $i -lt $inv.Count; $i++) { Set-Row $ws ($i + 2) $inv[$i] }

    $excel.Run("'" + $wb.Name + "'!RunFullAllocationSilent", $wb) | Out-Null

    Dump $wb "分配状态汇总表"
    Dump $wb "成功分配明细表"
    Dump $wb "数据异常明细表"
    Dump $wb "运行历史记录表"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}
