# =============================================================
# 升级预期调试日志19列.ps1
# 功能：将工作簿中「预期_调试日志*」工作表的表头统一为 19 列（见 调试日志19列规格说明.md）
# 用法：
#   powershell -File 升级预期调试日志19列.ps1
#   powershell -File 升级预期调试日志19列.ps1 -TargetFile "测试用例部分汇总.xlsm"
# 说明：
#   - 仅更新表头行；已有数据行需按 TC 文档手工/脚本补全后对照验收
#   - 旧 11 列表头会被整行替换，不会自动迁移旧列数据到新列
# =============================================================

param(
    [string]$WorkDir = $PSScriptRoot,
    [string]$TargetFile = ""
)

$debugHeaders = @(
    "物流单号", "SKU", "WMS退单号", "行号", "D", "处理序", "动态nextMinQty",
    "候选QC数", "被排除QC列表", "策略", "分配QC", "分配前QC剩余", "分配后QC剩余",
    "批号/效期组合数", "是否回溯重试", "实际回溯次数", "行状态", "错误码", "分配失败子类型"
)

$debugSheetNames = @("预期_调试日志", "预期_调试日志表")

function Update-DebugLogHeader($ws) {
    for ($c = 1; $c -le $debugHeaders.Count; $c++) {
        $cell = $ws.Cells(1, $c)
        $cell.NumberFormat = "@"
        $cell.Value2 = $debugHeaders[$c - 1]
        $cell.Font.Bold = $true
    }
    # 清除旧表头多余列（最多清到第 30 列）
    for ($c = $debugHeaders.Count + 1; $c -le 30; $c++) {
        if ($ws.Cells(1, $c).Value2 -ne $null) {
            $ws.Cells(1, $c).ClearContents()
        }
    }
    $ws.Columns.AutoFit() | Out-Null
}

function Process-Workbook($excel, $path) {
    Write-Host "处理: $path"
    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($path, 0, $false)
        $updated = 0
        foreach ($sheet in $wb.Worksheets) {
            if ($debugSheetNames -contains $sheet.Name) {
                Update-DebugLogHeader $sheet
                Write-Host "  [$($sheet.Name)] 表头已更新为 19 列"
                $updated++
            }
        }
        if ($updated -gt 0) {
            $wb.Save()
            Write-Host "  已保存 ($updated 张调试日志预期表)" -ForegroundColor Green
        } else {
            Write-Host "  未找到调试日志预期表，跳过保存"
        }
        $wb.Close($false)
    } catch {
        Write-Host "  错误: $_" -ForegroundColor Red
        if ($wb -ne $null) { $wb.Close($false) }
    }
}

Write-Host "=== 升级预期调试日志表头为 19 列 ===" -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false

try {
    if ($TargetFile -ne "") {
        $fullPath = if ([System.IO.Path]::IsPathRooted($TargetFile)) { $TargetFile } else { Join-Path $WorkDir $TargetFile }
        Process-Workbook $excel $fullPath
    } else {
        $patterns = @("SF*.xlsx", "SF*.xlsm", "测试用例*.xlsm", "测试用例*.xlsx")
        $files = @()
        foreach ($pat in $patterns) {
            $files += Get-ChildItem -Path $WorkDir -Filter $pat -ErrorAction SilentlyContinue
        }
        $files = $files | Sort-Object Name -Unique
        if ($files.Count -eq 0) {
            Write-Host "目录内未找到 Excel 工作簿，请将 SF*.xlsx 或汇总 xlsm 放入: $WorkDir"
        }
        foreach ($f in $files) {
            Process-Workbook $excel $f.FullName
        }
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    Write-Host "=== 脚本结束 ===" -ForegroundColor Cyan
}
