# 一次性修正：启用 5 个标准回归批次（BATCH-ERR-* 两个独立文件专用批次保持禁用）。
# 依据：批量测试计划备注列自述“汇总工作簿可执行”的批次集合；启用后 RunBatchTestPlan 才真正执行。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm"
try {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
} catch { throw "工作簿正在使用中，请先关闭：$path" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = "D:\cursor_practice\测试用例部分汇总_启用批次前_$stamp.xlsm"
Copy-Item -LiteralPath $path -Destination $backup -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path, 0, $false)
    $ws = $wb.Worksheets.Item("批量测试计划")
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    for ($r = 2; $r -le $lastRow; $r++) {
        $id = [string]$ws.Cells($r, 1).Value2
        if ($id -like "BATCH-ERR-*") {
            $ws.Cells($r, 2).Value2 = "否"
        } else {
            $ws.Cells($r, 2).Value2 = "是"
        }
        Write-Output ($id + " -> " + [string]$ws.Cells($r, 2).Value2)
    }
    $wb.Save()
    $wb.Close($false)
    Write-Output "BACKUP=$backup"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
