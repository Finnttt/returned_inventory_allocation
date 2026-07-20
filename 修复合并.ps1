# =============================================================
# 修复合并.ps1
# 功能：
#   1. 修复 预期_数据异常明细（重新合并9列数据）
#   2. 追加 批次1 独立xlsx 的 输入_退单表/输入_质检库存表/输入_配置
# 前置条件：请先关闭 测试用例部分汇总.xlsm 再运行！
# =============================================================

$workDir  = $PSScriptRoot
$xlsmPath = Join-Path $workDir "测试用例部分汇总.xlsm"

# 批次1 的独立文件（不含已在 xlsm 中的 SF0032/36/46/47/48）
# 也不含断言格式文件（SF0052/53/54/54A/54B/55）
$batch1Files = @(
    "SF0013_测试数据.xlsx",
    "SF0014_测试数据.xlsx",
    "SF0015_测试数据.xlsx",
    "SF0037_测试数据.xlsx",
    "SF0051_测试数据.xlsx",
    "SF0056_测试数据.xlsx",
    "SF0057_测试数据.xlsx",
    "SF0058_测试数据.xlsx",
    "SF0059_测试数据.xlsx",
    "SF0060_测试数据.xlsx"
)

# 运行批次映射（按物流单号）
$batchByShip = @{
    "SF3190000000013" = "批次1"; "SF3190000000014" = "批次1"
    "SF3190000000015" = "批次1"; "SF3190000000037" = "批次1"
    "SF3190000000051" = "批次1"; "SF3190000000056" = "批次1"
    "SF3190000000057" = "批次1"; "SF3190000000058" = "批次1"
    "SF3190000000059" = "批次1"; "SF3190000000060" = "批次1"
}

# ── 工具函数 ─────────────────────────────────────────────────

# 获取工作表第1行实际最后一列列号（16384 = Excel 最大列数 XFD）
function Get-HeaderColCount($ws) {
    return $ws.Cells(1, 16384).End(-4159).Column  # -4159 = xlToLeft
}

# 获取指定列的最后一行
function Get-LastRow($ws, $colIdx) {
    return $ws.Cells($ws.Rows.Count, $colIdx).End(-4162).Row  # -4162 = xlUp
}

# 从源表追加所有数据行到目标表
# 目标表已有"运行批次"首列（原数据从第2列起）
# batchName：该批次名称
function Append-AllCols($srcWs, $dstWs, $batchName) {
    $srcColCount = Get-HeaderColCount $srcWs
    $srcLastRow  = Get-LastRow $srcWs 1
    if ($srcLastRow -lt 2) {
        Write-Host "      (无数据行，跳过)"
        return
    }

    # 目标表最后一行（用第2列定位，第2列是原第1列数据）
    $dstLastRow = Get-LastRow $dstWs 2
    if ($dstLastRow -lt 1) { $dstLastRow = 1 }

    $count = 0
    for ($r = 2; $r -le $srcLastRow; $r++) {
        $dstRow = $dstLastRow + ($r - 1)  # 追加位置
        $dstWs.Cells($dstRow, 1).Value2 = $batchName  # 运行批次列
        for ($c = 1; $c -le $srcColCount; $c++) {
            $val = $srcWs.Cells($r, $c).Value2
            if ($val -ne $null -and -not ($val -is [System.DBNull])) {
                $dstWs.Cells($dstRow, $c + 1).Value2 = $val
            }
        }
        $count++
    }
    Write-Host "      追加 $count 行（$($srcWs.Name) → $($dstWs.Name)，列数=$srcColCount）"
}

# 直接追加行（不加运行批次列，用于输入表）
function Append-InputSheet($srcWs, $dstWs) {
    $srcColCount = Get-HeaderColCount $srcWs
    $srcLastRow  = Get-LastRow $srcWs 1
    if ($srcLastRow -lt 2) { return }

    # 目标表最后一行（用第1列定位）
    $dstLastRow = Get-LastRow $dstWs 1
    if ($dstLastRow -lt 1) { $dstLastRow = 1 }

    $count = 0
    for ($r = 2; $r -le $srcLastRow; $r++) {
        $dstRow = $dstLastRow + ($r - 1)
        for ($c = 1; $c -le $srcColCount; $c++) {
            $val = $srcWs.Cells($r, $c).Value2
            if ($val -ne $null -and -not ($val -is [System.DBNull])) {
                $dstWs.Cells($dstRow, $c).Value2 = $val
            }
        }
        $count++
    }
    Write-Host "      追加 $count 行（$($srcWs.Name) → $($dstWs.Name)）"
}

# ================================================================
# 主程序
# ================================================================
Write-Host "=== 开始修复合并 ===" -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false
$excel.AutomationSecurity = 1

try {
    Write-Host "`n[1] 打开 xlsm..."
    $wb = $excel.Workbooks.Open($xlsmPath, 0, $false)

    # ── 步骤1：清空 预期_数据异常明细 的数据行并重新合并 ──────────
    Write-Host "`n[2] 修复 预期_数据异常明细..."
    $dstAnomal = $wb.Worksheets.Item("预期_数据异常明细")

    # 清除第2行起的所有数据
    $anomLastRow = Get-LastRow $dstAnomal 2
    if ($anomLastRow -ge 2) {
        $dstAnomal.Range("A2:Z$anomLastRow").ClearContents()
        Write-Host "  清除旧数据：第2~$anomLastRow 行"
    }

    # 重新从各独立 xlsx 追加
    foreach ($fname in $batch1Files) {
        $fpath = Join-Path $workDir $fname
        if (-not (Test-Path $fpath)) { continue }
        try {
            $srcWb = $excel.Workbooks.Open($fpath, 0, $true)
            $shipNo = [string]($srcWb.Worksheets.Item("输入_配置").Cells(2,1).Value2)
            $batchName = if ($batchByShip.ContainsKey($shipNo)) { $batchByShip[$shipNo] } else { "批次1" }

            # 尝试两种可能的表名
            $anomSheetName = $null
            foreach ($s in $srcWb.Worksheets) {
                if ($s.Name -like "*数据异常明细*") { $anomSheetName = $s.Name }
            }
            if ($anomSheetName) {
                Write-Host "  处理 $fname（$shipNo）[$batchName]..."
                $srcAnom = $srcWb.Worksheets.Item($anomSheetName)
                Append-AllCols $srcAnom $dstAnomal $batchName
            } else {
                Write-Host "  跳过 $fname（无异常明细表）"
            }
            $srcWb.Close($false)
        } catch {
            Write-Host "  打开失败：$fname — $_"
        }
    }

    # ── 步骤2：追加输入数据 ──────────────────────────────────────
    Write-Host "`n[3] 追加 输入_退单表 / 输入_质检库存表 / 输入_配置..."
    $dstReturn    = $wb.Worksheets.Item("输入_退单表")
    $dstInventory = $wb.Worksheets.Item("输入_质检库存表")
    $dstConfig    = $wb.Worksheets.Item("输入_配置")

    foreach ($fname in $batch1Files) {
        $fpath = Join-Path $workDir $fname
        if (-not (Test-Path $fpath)) { continue }
        try {
            $srcWb = $excel.Workbooks.Open($fpath, 0, $true)
            $shipNo = [string]($srcWb.Worksheets.Item("输入_配置").Cells(2,1).Value2)
            Write-Host "  处理 $fname（$shipNo）..."

            # 输入_退单表
            Append-InputSheet ($srcWb.Worksheets.Item("输入_退单表"))    $dstReturn
            # 输入_质检库存表
            Append-InputSheet ($srcWb.Worksheets.Item("输入_质检库存表")) $dstInventory
            # 输入_配置（每个文件只有1行配置数据）
            Append-InputSheet ($srcWb.Worksheets.Item("输入_配置"))       $dstConfig

            $srcWb.Close($false)
        } catch {
            Write-Host "  打开失败：$fname — $_"
        }
    }

    # ── 保存 ──────────────────────────────────────────────────
    Write-Host "`n保存 xlsm..."
    $wb.Save()
    $wb.Close($false)
    Write-Host "✓ 全部完成！" -ForegroundColor Green

} catch {
    Write-Host "错误：$_" -ForegroundColor Red
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    Write-Host "=== 脚本结束 ===" -ForegroundColor Cyan
}
