# 捕获 5 套 DataSet 的简版调试日志实际输出（SF0003/16/27/28/49），
# 输入取自汇总工作簿，配置取各自配置行（SF0028 最大回溯次数=10），统一 调试日志级别=简版。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$testPath = "D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm"
$prodPath = "D:\cursor_practice\returned_inventory_allocation\退货入库分配系统.xlsm"
$ships = @(
    @{ No = "SF3190000000003"; MaxBT = 200 },
    @{ No = "SF3190000000016"; MaxBT = 200 },
    @{ No = "SF3190000000027"; MaxBT = 200 },
    @{ No = "SF3190000000028"; MaxBT = 10 },
    @{ No = "SF3190000000049"; MaxBT = 200 }
)

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
function Get-RowsByShip($ws, [int]$keyCol, [string]$ship, [int]$maxCol, [int]$qtyCol) {
    $lastRow = $ws.Cells($ws.Rows.Count, $keyCol).End(-4162).Row
    $rows = @()
    for ($r = 2; $r -le $lastRow; $r++) {
        if ([string]$ws.Cells($r, $keyCol).Text -eq $ship) {
            $vals = @()
            for ($c = 1; $c -le $maxCol; $c++) {
                if ($c -eq $qtyCol) { $vals += [double]$ws.Cells($r, $c).Value2 }
                else { $vals += [string]$ws.Cells($r, $c).Text }
            }
            $rows += , $vals
        }
    }
    return , $rows
}

try {
    $twb = $excel.Workbooks.Open($testPath, 0, $true)
    $retWs = $twb.Worksheets.Item("输入_退单表")
    $invWs = $twb.Worksheets.Item("输入_质检库存表")

    foreach ($s in $ships) {
        $ship = $s.No
        $retRows = Get-RowsByShip $retWs 1 $ship 5 5
        $invRows = Get-RowsByShip $invWs 1 $ship 6 6

        $tmp = Join-Path $env:TEMP ("dbgcap_" + [System.Guid]::NewGuid().ToString("N") + ".xlsm")
        Copy-Item $prodPath $tmp -Force
        $wb = $excel.Workbooks.Open($tmp, 0, $false)
        $wsR = $wb.Worksheets.Item("输入_退单表")
        for ($i = 0; $i -lt $retRows.Count; $i++) { Set-Row $wsR ($i + 2) $retRows[$i] }
        $wsI = $wb.Worksheets.Item("输入_质检库存表")
        for ($i = 0; $i -lt $invRows.Count; $i++) { Set-Row $wsI ($i + 2) $invRows[$i] }

        $cfgWs = $wb.Worksheets.Item("输入_配置")
        $lastCfg = $cfgWs.Cells($cfgWs.Rows.Count, 1).End(-4162).Row
        for ($r = 2; $r -le $lastCfg; $r++) {
            $pname = [string]$cfgWs.Cells($r, 1).Value2
            if ($pname -eq "最大回溯次数") { $cfgWs.Cells($r, 2).Value2 = [string]$s.MaxBT }
            if ($pname -eq "调试日志级别") { $cfgWs.Cells($r, 2).Value2 = "简版" }
        }

        $excel.Run("'" + $wb.Name + "'!RunFullAllocationSilent", $wb) | Out-Null

        Write-Output ("########## " + $ship + " (MaxBT=" + $s.MaxBT + ") 退单行=" + $retRows.Count + " 库存行=" + $invRows.Count)
        foreach ($sheetName in @("分配状态汇总表", "调试日志", "运行历史记录表")) {
            $ws = $wb.Worksheets.Item($sheetName)
            $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
            $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column
            Write-Output ("== " + $sheetName + " rows=" + ($lastRow - 1) + " cols=" + $lastCol)
            for ($r = 2; $r -le $lastRow; $r++) {
                $vals = @()
                for ($c = 1; $c -le $lastCol; $c++) { $vals += [string]$ws.Cells($r, $c).Text }
                Write-Output ($vals -join "|")
            }
        }
        $wb.Close($false)
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
    $twb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
