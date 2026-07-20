# =============================================================
# 合并预期结果.ps1
# 功能：
#   1. 给 xlsm 的预期表补充"运行批次"首列，并填写批次值
#   2. 把独立 SF*.xlsx 的汇总、异常明细和调试日志合并到 xlsm
#   3. 在 xlsm 中维护"预期_断言"表，按物流单号覆盖旧数据
#   4. 把 SF0062/63/64 的标准输入与配置合并到批量测试源数据
#   5. 支持重复执行：不重复插列、不重复追加同一物流单号
# 前置条件：
#   请先在 Excel 中关闭 测试用例部分汇总.xlsm，再运行本脚本！
# =============================================================

$workDir  = $PSScriptRoot
$xlsmPath = Join-Path $workDir "测试用例部分汇总.xlsm"
$scriptFailed = $false

# ── 占用检测与自动备份（2026-07-19 补强） ─────────────────────
# 目标工作簿被 Excel 占用时立即中止，避免合并写入与人工编辑互相覆盖；
# 每次运行先在上级目录留一份时间戳备份，与其他工具脚本口径一致。
if (-not (Test-Path -LiteralPath $xlsmPath)) {
    throw "目标工作簿不存在：$xlsmPath"
}
try {
    $stream = [System.IO.File]::Open(
        $xlsmPath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None)
    $stream.Close()
} catch {
    throw "目标工作簿正在使用中，请先关闭后重试：$xlsmPath"
}
$mergeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$mergeBackup = Join-Path (Split-Path $workDir -Parent) ("测试用例部分汇总_合并预期前_{0}.xlsm" -f $mergeStamp)
Copy-Item -LiteralPath $xlsmPath -Destination $mergeBackup -Force
Write-Host "已备份：$mergeBackup"

# ── 运行批次映射（物流单号 → 批次名） ──────────────────────────
# 批次1: 所有标准运行, 批次2: SF0028(特殊maxBT=10)
# 批次3-8: 断言格式（单独跑），由文件名决定（见断言合并段）
$batchByShip = @{
    "SF3190000000000" = "批次1"
    "SF3190000000001" = "批次1"
    "SF3190000000002" = "批次1"
    "SF3190000000003" = "批次1"
    "SF3190000000005" = "批次1"
    "SF3190000000013" = "批次1"
    "SF3190000000014" = "批次1"
    "SF3190000000015" = "批次1"
    "SF3190000000016" = "批次1"
    "SF3190000000017" = "批次1"
    "SF3190000000018" = "批次1"
    "SF3190000000019" = "批次1"
    "SF3190000000020" = "批次1"
    "SF3190000000026" = "批次1"
    "SF3190000000027" = "批次1"
    "SF3190000000028" = "批次2"  # 特殊：最大回溯=10
    "SF3190000000032" = "批次1"
    "SF3190000000036" = "批次1"
    "SF3190000000037" = "批次1"
    "SF3190000000046" = "批次1"
    "SF3190000000047" = "批次1"
    "SF3190000000048" = "批次1"
    "SF3190000000049" = "批次1"
    "SF3190000000051" = "批次1"
    "SF3190000000056" = "批次1"
    "SF3190000000057" = "批次1"
    "SF3190000000058" = "批次1"
    "SF3190000000059" = "批次1"
    "SF3190000000060" = "批次1"
    "SF3190000000062" = "批次9"
    "SF3190000000063" = "批次9"
    "SF3190000000064" = "批次9"
}

# ── 断言格式文件 → 批次名（按文件名） ───────────────────────────
$assertBatch = @{
    "SF0052" = "批次3"
    "SF0053" = "批次4"
    "SF0054"  = "批次5"
    "SF0054A" = "批次6"
    "SF0054B" = "批次7"
    "SF0055"  = "批次8"
}

# ── 独立 xlsx 不需要合并的（已在 xlsm 里，会重复） ──────────────
$skipFiles = @("SF0032","SF0036","SF0046","SF0047","SF0048")

# 这三套是后补的标准 DataSet，需要同时进入汇总工作簿的输入区。
# 故意损坏表头/配置的 DataSet 保持独立，不能污染标准批量测试源数据。
$inputMergeShips = @{
    "SF3190000000062" = $true
    "SF3190000000063" = $true
    "SF3190000000064" = $true
}

# ================================================================
# 工具函数
# ================================================================

# 在工作表第1行找到指定列名的列号
function Get-ColIndex($ws, $headerName) {
    $lastCol = $ws.UsedRange.Columns.Count
    for ($c = 1; $c -le ($lastCol + 5); $c++) {
        if ($ws.Cells(1, $c).Value2 -eq $headerName) { return $c }
    }
    return -1
}

function Rename-SheetIfNeeded($wb, $oldName, $newName) {
    $newExists = $false
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq $newName) { $newExists = $true }
    }
    if ($newExists) { return }

    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq $oldName) {
            $s.Name = $newName
            Write-Host "  Sheet 重命名: [$oldName] -> [$newName]"
            return
        }
    }
}

# 取工作表最后一行（从指定列往上找）
function Get-LastRow($ws, $colIdx) {
    return $ws.Cells($ws.Rows.Count, $colIdx).End(-4162).Row  # xlUp
}

# 给工作表首列插入"运行批次"，并按物流单号填写批次值
function Add-BatchColumn($ws, $batchByShip) {
    # 跳过空表（只有1行或0行数据）
    $used = $ws.UsedRange.Rows.Count
    if ($used -lt 1) { return }

    if ($ws.Cells(1, 1).Value2 -ne "运行批次") {
        # 在 A 列前插入新列
        $ws.Columns("A:A").Insert() | Out-Null
        $ws.Cells(1, 1).Value2 = "运行批次"
    }

    # 找物流单号所在列（插入后它已向右移了1列）
    $shipCol = Get-ColIndex $ws "物流单号"
    if ($shipCol -lt 1) {
        Write-Host "  警告: [$($ws.Name)] 未找到'物流单号'列，跳过批次填写"
        return
    }

    $lastRow = Get-LastRow $ws $shipCol
    for ($r = 2; $r -le $lastRow; $r++) {
        $shipNo = $ws.Cells($r, $shipCol).Value2
        if ($shipNo -and $batchByShip.ContainsKey($shipNo)) {
            $ws.Cells($r, 1).Value2 = $batchByShip[$shipNo]
        } else {
            $ws.Cells($r, 1).Value2 = "批次1"  # 默认
        }
    }
    Write-Host "  [$($ws.Name)] 完成：$($lastRow - 1) 行 × 批次列已填写"
}

function Remove-ExistingShipRows($dstWs, $shipNo) {
    $shipCol = Get-ColIndex $dstWs "物流单号"
    if ($shipCol -lt 1) { return }

    $lastRow = Get-LastRow $dstWs $shipCol
    for ($r = $lastRow; $r -ge 2; $r--) {
        if ($dstWs.Cells($r, $shipCol).Value2 -eq $shipNo) {
            $dstWs.Rows($r).Delete() | Out-Null
        }
    }
}

# 清理旧版本合并时遗留的“物流单号为空、WMS退单号有值”行。
# 这类行无法仅按物流单号删除，会导致脚本每运行一次就重复追加。
function Remove-ExistingSourceRows($srcWs, $dstWs, $shipNo) {
    # 一个源工作簿可能包含多个物流单号（例如 SF0037 文件同时含 SF0037/SF0038）。
    # 必须覆盖源表中出现的全部物流单号，不能只使用配置表第 2 行的默认物流单号。
    $srcShipCol = Get-ColIndex $srcWs "物流单号"
    $sourceShips = @{}
    if ($srcShipCol -gt 0) {
        $srcShipLastRow = Get-LastRow $srcWs $srcShipCol
        for ($r = 2; $r -le $srcShipLastRow; $r++) {
            $sourceShipNo = [string]$srcWs.Cells($r, $srcShipCol).Value2
            if ($sourceShipNo -eq "") { $sourceShipNo = [string]$shipNo }
            if ($sourceShipNo -ne "") { $sourceShips[$sourceShipNo] = $true }
        }
    }
    if ($sourceShips.Count -eq 0 -and [string]$shipNo -ne "") {
        $sourceShips[[string]$shipNo] = $true
    }
    foreach ($sourceShipNo in $sourceShips.Keys) {
        Remove-ExistingShipRows $dstWs $sourceShipNo
    }

    $srcWmsCol = Get-ColIndex $srcWs "WMS退单号"
    $dstWmsCol = Get-ColIndex $dstWs "WMS退单号"
    $dstShipCol = Get-ColIndex $dstWs "物流单号"
    if ($srcWmsCol -lt 1 -or $dstWmsCol -lt 1 -or $dstShipCol -lt 1) { return }

    $sourceWms = @{}
    $srcLastRow = Get-LastRow $srcWs $srcWmsCol
    for ($r = 2; $r -le $srcLastRow; $r++) {
        $wmsNo = [string]$srcWs.Cells($r, $srcWmsCol).Value2
        if ($wmsNo -ne "") { $sourceWms[$wmsNo] = $true }
    }

    $dstLastRow = Get-LastRow $dstWs $dstWmsCol
    for ($r = $dstLastRow; $r -ge 2; $r--) {
        $wmsNo = [string]$dstWs.Cells($r, $dstWmsCol).Value2
        $existingShipNo = [string]$dstWs.Cells($r, $dstShipCol).Value2
        if ($existingShipNo -eq "" -and $sourceWms.ContainsKey($wmsNo)) {
            $dstWs.Rows($r).Delete() | Out-Null
        }
    }
}

# 从源工作表追加数据行到目标工作表（目标表已有"运行批次"首列）
# srcShipNo: 该文件的物流单号
# batchName: 该文件对应的批次名
function Append-Sheet($srcWs, $dstWs, $shipNo, $batchName) {
    if ($srcWs -eq $null) { return }

    Remove-ExistingSourceRows $srcWs $dstWs $shipNo

    # 目标表的列顺序（插入运行批次后）：运行批次 | 原col1 | 原col2 ...
    # 源表的列顺序：原col1 | 原col2 ...
    # 先找目标表最后一行
    $dstLastRow = Get-LastRow $dstWs 2  # 用第2列（原第1列）定位
    if ($dstLastRow -lt 1) { $dstLastRow = 1 }

    # 源表数据行（跳过表头行1）
    $srcLastRow = Get-LastRow $srcWs 1
    $srcColCount = $srcWs.UsedRange.Columns.Count
    $srcShipCol = Get-ColIndex $srcWs "物流单号"

    for ($r = 2; $r -le $srcLastRow; $r++) {
        $dstRow = $dstLastRow + $r - 1  # 追加行号
        # 写批次
        $dstWs.Cells($dstRow, 1).NumberFormat = "@"
        $dstWs.Cells($dstRow, 1).Value2 = [string]$batchName
        # 拷贝源数据各列到目标表（目标从第2列开始）
        for ($c = 1; $c -le $srcColCount; $c++) {
            $val = $srcWs.Cells($r, $c).Value2
            if ($c -eq $srcShipCol -and [string]$val -eq "") {
                $val = $shipNo
            }
            if ($val -ne $null) {
                $dstWs.Cells($dstRow, $c + 1).NumberFormat = "@"
                $dstWs.Cells($dstRow, $c + 1).Value2 = [string]$val
            }
        }
    }
    $rowsAdded = $srcLastRow - 1
    if ($rowsAdded -gt 0) {
        Write-Host "    [$($srcWs.Name)] → [$($dstWs.Name)] 追加 $rowsAdded 行"
    }
}

# 按表头名称合并标准输入或配置行，并按物流单号覆盖旧数据。
# 保留源单元格格式和值类型，确保行号仍是文本、数量仍是数值。
function Append-InputSheet($srcWs, $dstWs, $shipNo) {
    if ($srcWs -eq $null -or $dstWs -eq $null) { return }

    Remove-ExistingShipRows $dstWs $shipNo

    $srcLastRow = Get-LastRow $srcWs 1
    if ($srcLastRow -lt 2) { return }

    $dstLastRow = Get-LastRow $dstWs 1
    if ($dstLastRow -lt 1) { $dstLastRow = 1 }
    $srcColCount = $srcWs.UsedRange.Columns.Count
    $added = 0

    for ($r = 2; $r -le $srcLastRow; $r++) {
        $dstRow = $dstLastRow + $added + 1
        for ($c = 1; $c -le $srcColCount; $c++) {
            $header = [string]$srcWs.Cells(1, $c).Value2
            $dstCol = Get-ColIndex $dstWs $header
            if ($dstCol -lt 1) { continue }

            $srcCell = $srcWs.Cells($r, $c)
            $dstCell = $dstWs.Cells($dstRow, $dstCol)
            $srcCell.Copy($dstCell) | Out-Null
        }
        $added++
    }

    Write-Host "    [$($srcWs.Name)] → [$($dstWs.Name)] 覆盖合并 $added 行"
}

# ================================================================
# 主程序开始
# ================================================================

Write-Host "=== 开始合并预期结果 ===" -ForegroundColor Cyan
Write-Host "目标文件: $xlsmPath"

# 启动 Excel COM 实例
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false
$excel.AutomationSecurity = 1  # 允许所有宏（仅用于打开文件）

try {
    # 打开 xlsm（false=ReadWrite）
    Write-Host "`n[1/4] 打开 xlsm..."
    $wb = $excel.Workbooks.Open($xlsmPath, 0, $false)
    Rename-SheetIfNeeded $wb "预期_数据异常明细表" "预期_数据异常明细"
    Rename-SheetIfNeeded $wb "预期_调试日志表" "预期_调试日志"
    Rename-SheetIfNeeded $wb "运行历史记录" "运行历史记录表"

    # ── 步骤1：给 xlsm 的 5 张预期_* 表加"运行批次"首列 ─────────────
    Write-Host "`n[2/4] 给 xlsm 预期表插入'运行批次'列..."
    $preqiSheets = @("预期_汇总表","预期_成功分配明细","预期_数据异常明细","预期_调试日志")
    foreach ($sn in $preqiSheets) {
        try {
            $ws = $wb.Worksheets.Item($sn)
            Add-BatchColumn $ws $batchByShip
        } catch {
            Write-Host "  跳过 [$sn]：未找到该工作表"
        }
    }
    # 预期_运行历史记录（只允许插入一次）
    try {
        $wsHist = $wb.Worksheets.Item("预期_运行历史记录")
        if ($wsHist.UsedRange.Rows.Count -gt 1 -and $wsHist.Cells(1,1).Value2 -ne "运行批次") {
            $wsHist.Columns("A:A").Insert() | Out-Null
            $wsHist.Cells(1,1).Value2 = "运行批次"
            Write-Host "  [预期_运行历史记录] 已插入批次列（请手工补填批次值）"
        }
    } catch {}

    # ── 步骤2：获取目标工作表引用 ────────────────────────────────────
    $dstSummary = $wb.Worksheets.Item("预期_汇总表")
    $dstAnomal  = $wb.Worksheets.Item("预期_数据异常明细")
    $dstDebug   = $wb.Worksheets.Item("预期_调试日志")
    $dstReturn  = $wb.Worksheets.Item("输入_退单表")
    $dstInventory = $wb.Worksheets.Item("输入_质检库存表")
    $dstConfig  = $wb.Worksheets.Item("输入_配置")

    # ── 步骤3：新建"预期_断言"表 ─────────────────────────────────────
    Write-Host "`n[3/4] 新建'预期_断言'表..."
    $assertSheetExists = $false
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "预期_断言") { $assertSheetExists = $true } }
    if (-not $assertSheetExists) {
        # 在最后一张表之后新建
        $wsAssert = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $wsAssert.Name = "预期_断言"
        # 写表头
        $wsAssert.Cells(1,1).Value2 = "运行批次"
        $wsAssert.Cells(1,2).Value2 = "物流单号"
        $wsAssert.Cells(1,3).Value2 = "断言项"
        $wsAssert.Cells(1,4).Value2 = "预期"
        Write-Host "  '预期_断言' 表已创建"
    } else {
        $wsAssert = $wb.Worksheets.Item("预期_断言")
        Write-Host "  '预期_断言' 表已存在，追加数据"
    }

    # ── 步骤4：逐一处理独立 xlsx 文件 ──────────────────────────────
    Write-Host "`n[4/4] 合并独立 xlsx 数据..."

    $allXlsx = Get-ChildItem $workDir -Filter "SF*.xlsx" | Sort-Object Name
    foreach ($f in $allXlsx) {
        # 提取文件关键字（如 SF0013, SF0054A）
        $key = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) -replace "_测试数据",""

        # 跳过已在 xlsm 里的重复文件
        $isSkip = $false
        foreach ($sk in $skipFiles) { if ($key -eq $sk) { $isSkip = $true } }
        if ($isSkip) { Write-Host "  跳过 $key（已在 xlsm）"; continue }

        Write-Host "  处理 $key ..."

        # 打开 xlsx（只读）
        try {
            $srcWb = $excel.Workbooks.Open($f.FullName, 0, $true)  # ReadOnly=true
        } catch {
            Write-Host "    打开失败（可能被占用）：$($f.Name)"
            continue
        }

        # 读该文件物流单号
        try {
            $cfgWs = $srcWb.Worksheets.Item("输入_配置")
            $shipNo = $cfgWs.Cells(2, 1).Value2
        } catch { $shipNo = "" }

        # 判断文件类型：断言格式 or 标准格式
        $isAssertFormat = $false
        foreach ($s in $srcWb.Worksheets) {
            if ($s.Name -eq "预期_断言") { $isAssertFormat = $true }
        }

        if ($isAssertFormat) {
            # ── 断言格式处理 ──────────────────────────────────────
            $batchName = if ($assertBatch.ContainsKey($key)) { $assertBatch[$key] } else { "批次?" }
            try {
                $srcAssert = $srcWb.Worksheets.Item("预期_断言")
                Remove-ExistingShipRows $wsAssert $shipNo
                $srcLastRow = Get-LastRow $srcAssert 1
                $dstAssertLastRow = Get-LastRow $wsAssert 3  # 用第3列"断言项"定位
                if ($dstAssertLastRow -lt 1) { $dstAssertLastRow = 1 }

                for ($r = 2; $r -le $srcLastRow; $r++) {
                    $dstR = $dstAssertLastRow + $r - 1
                    for ($c = 1; $c -le 4; $c++) {
                        $wsAssert.Cells($dstR, $c).NumberFormat = "@"
                    }
                    $wsAssert.Cells($dstR, 1).Value2 = [string]$batchName
                    $wsAssert.Cells($dstR, 2).Value2 = [string]$shipNo
                    $wsAssert.Cells($dstR, 3).Value2 = [string]$srcAssert.Cells($r, 1).Value2  # 断言项
                    # 第4列可能是"预期"或"期望值"，都映射到目标第4列
                    $wsAssert.Cells($dstR, 4).Value2 = [string]$srcAssert.Cells($r, 2).Value2
                }
                Write-Host "    断言 → 预期_断言：$($srcLastRow-1) 行 [$batchName]"
            } catch { Write-Host "    读取预期_断言失败" }

        } else {
            # ── 标准格式处理 ──────────────────────────────────────
            $batchName = if ($batchByShip.ContainsKey($shipNo)) { $batchByShip[$shipNo] } else { "批次1" }

            # 新增标准 DataSet 同时合并输入和配置，供 RunBatchTestPlan 实际执行。
            if ($inputMergeShips.ContainsKey([string]$shipNo)) {
                Append-InputSheet $srcWb.Worksheets.Item("输入_退单表") $dstReturn $shipNo
                Append-InputSheet $srcWb.Worksheets.Item("输入_质检库存表") $dstInventory $shipNo
                Append-InputSheet $srcWb.Worksheets.Item("输入_配置") $dstConfig $shipNo
            }

            # 合并 预期_汇总表
            try {
                $srcSum = $srcWb.Worksheets.Item("预期_汇总表")
                Append-Sheet $srcSum $dstSummary $shipNo $batchName
            } catch { Write-Host "    无预期_汇总表，跳过" }

            # 合并 预期_数据异常明细
            try {
                $srcAnom = $srcWb.Worksheets.Item("预期_数据异常明细")
                Append-Sheet $srcAnom $dstAnomal $shipNo $batchName
            } catch {
                # 可能表名有"表"后缀
                try {
                    $srcAnom = $srcWb.Worksheets.Item("预期_数据异常明细表")
                    Append-Sheet $srcAnom $dstAnomal $shipNo $batchName
                } catch { }
            }

            # 合并 预期_调试日志（仅有该表且有数据的 DataSet 会追加）
            try {
                $srcDebug = $srcWb.Worksheets.Item("预期_调试日志")
                Append-Sheet $srcDebug $dstDebug $shipNo $batchName
            } catch {
                try {
                    $srcDebug = $srcWb.Worksheets.Item("预期_调试日志表")
                    Append-Sheet $srcDebug $dstDebug $shipNo $batchName
                } catch { }
            }
        }

        $srcWb.Close($false)
    }

    # ── 保存 xlsm ────────────────────────────────────────────────
    Write-Host "`n保存 xlsm..."
    $wb.Save()
    $wb.Close($false)
    Write-Host "✓ 保存完成！" -ForegroundColor Green

} catch {
    $scriptFailed = $true
    Write-Host "错误：$_" -ForegroundColor Red
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    Write-Host "=== 脚本结束 ===" -ForegroundColor Cyan
}

if ($scriptFailed) { exit 1 }
exit 0
