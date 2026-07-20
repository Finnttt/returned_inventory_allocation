# 生成测试 DataSet 的 Excel 工作簿
# 覆盖：M04 标准化边界、M05 分配前校验、M03/E12 结构防御等测试数据。
# 用法示例：
#   powershell -File 生成测试数据Excel.ps1 -Target OnlyM05     （仅生成 TC-13~20、TC-37）
#   powershell -File 生成测试数据Excel.ps1 -Target OnlyPendingTC（仅生成 TC-28/40/41）
#   powershell -File 生成测试数据Excel.ps1 -Target OnlySF0055  （仅生成 SF0055）

param(
    [string]$Target = "All"
)

$outputDir = $PSScriptRoot

# ---------- 辅助：写文本单元格 ----------
function wt($sheet, $r, $c, $v) {
    $cell = $sheet.Cells($r, $c)
    $cell.NumberFormat = "@"
    $cell.Value2 = [string]$v
}

# ---------- 辅助：写数值单元格（不改格式，直接赋 double）----------
function wn($sheet, $r, $c, $v) {
    $sheet.Cells($r, $c).Value2 = [double]$v
}

# ---------- 辅助：标题行样式 ----------
function hdr($sheet, $row, $cols, $bg) {
    for ($c = 1; $c -le $cols; $c++) {
        $cell = $sheet.Cells($row, $c)
        $cell.Font.Bold = $true
        $cell.Interior.Color = $bg
        $cell.Borders.LineStyle = 1
    }
}

# ---------- 辅助：数据行边框 ----------
function bdr($sheet, $row, $cols) {
    for ($c = 1; $c -le $cols; $c++) { $sheet.Cells($row, $c).Borders.LineStyle = 1 }
}

# ---------- 辅助：写配置 sheet（横向格式，与 modConfig.bas 一致）----------
function config($wb, $shipmentNo, $tcNo, $maxBt, $dbgLv, $batchMode, $sentinel, $note = "") {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "输入_配置"
    $headers = @("物流单号","TC编号","最大回溯次数","调试日志级别","批号比较模式","无保质期哨兵值","备注")
    for ($c = 1; $c -le $headers.Count; $c++) { wt $ws 1 $c ($headers[$c-1]) }
    hdr $ws 1 $headers.Count 0xFFF2CC
    wt $ws 2 1 $shipmentNo; wt $ws 2 2 $tcNo
    wn $ws 2 3 $maxBt; wt $ws 2 4 $dbgLv
    wt $ws 2 5 $batchMode; wt $ws 2 6 $sentinel; wt $ws 2 7 $note
    bdr $ws 2 $headers.Count
    $ws.Columns.AutoFit() | Out-Null
}

# ---------- 辅助：写预期_断言 sheet（E12 等结构级异常用例）----------
function assertionSheet($wb, $expectedLabel, $snippet1, $snippet2 = "") {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_断言"
    wt $ws 1 1 "断言项"; wt $ws 1 2 "期望值"
    hdr $ws 1 2 0xFFF2CC
    wt $ws 2 1 "期望错误"; wt $ws 2 2 $expectedLabel
    bdr $ws 2 2
    wt $ws 3 1 "关键报错片段1"; wt $ws 3 2 $snippet1
    bdr $ws 3 2
    if ($snippet2 -ne "") {
        wt $ws 4 1 "关键报错片段2"; wt $ws 4 2 $snippet2
        bdr $ws 4 2
    }
    $ws.Columns.AutoFit() | Out-Null
}

# ============================================================
# 辅助：写预期_汇总表（支持多行）
# 用法：先调 summaryHeader，再对每行调 summaryRow
# ============================================================
function summaryHeader($wb) {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_汇总表"
    $cols = @("物流单号","WMS退单号","退单号状态","原因")
    for ($c = 1; $c -le $cols.Count; $c++) { wt $ws 1 $c ($cols[$c-1]) }
    hdr $ws 1 $cols.Count 0xE2EFDA
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}
function summaryRow($ws, $row, $sfNum, $wmsNum, $status, $reason) {
    wt $ws $row 1 $sfNum; wt $ws $row 2 $wmsNum; wt $ws $row 3 $status; wt $ws $row 4 $reason
    bdr $ws $row 4
}

# ============================================================
# 辅助：写预期_数据异常明细 sheet
# ============================================================
function anomalyHeader($wb) {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_数据异常明细"
    $h = @("来源表","原始行号","物流单号","WMS退单号","SKU","异常字段名","原始值","错误码","原因说明")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xFCE4D6
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}
function anomalyRow($ws, $row, $src, $rowNum, $sf, $wms, $sku, $field, $raw, $code, $reason) {
    wt $ws $row 1 $src; wn $ws $row 2 $rowNum; wt $ws $row 3 $sf; wt $ws $row 4 $wms
    wt $ws $row 5 $sku; wt $ws $row 6 $field; wt $ws $row 7 $raw
    wt $ws $row 8 $code; wt $ws $row 9 $reason
    bdr $ws $row 9
}

# ============================================================
# 辅助：写预期_调试日志 sheet（19 列，与 调试日志19列规格说明.md 一致）
# ============================================================
function debugLogHeader($wb) {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_调试日志"
    $h = @(
        "物流单号","SKU","WMS退单号","行号","D","处理序","动态nextMinQty",
        "候选QC数","被排除QC列表","策略","分配QC","分配前QC剩余","分配后QC剩余",
        "批号/效期组合数","是否回溯重试","实际回溯次数","行状态","错误码","分配失败子类型"
    )
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xDDEBF7
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}
function debugLogRow($ws, $row, $sf, $sku, $wms, $lineNo, $d, $procOrder, $nextMin, $candQc, $exclQc,
    $strategy, $usedQc, $qcBefore, $qcAfter, $comboCount, $isRetry, $btCount, $lineStatus, $errCode, $failSub) {
    wt $ws $row 1 $sf; wt $ws $row 2 $sku; wt $ws $row 3 $wms; wt $ws $row 4 $lineNo
    wn $ws $row 5 $d; wt $ws $row 6 $procOrder; wt $ws $row 7 $nextMin
    wt $ws $row 8 $candQc; wt $ws $row 9 $exclQc; wt $ws $row 10 $strategy; wt $ws $row 11 $usedQc
    wt $ws $row 12 $qcBefore; wt $ws $row 13 $qcAfter; wt $ws $row 14 $comboCount
    wt $ws $row 15 $isRetry; wn $ws $row 16 $btCount; wt $ws $row 17 $lineStatus
    wt $ws $row 18 $errCode; wt $ws $row 19 $failSub
    bdr $ws $row 19
}

# ============================================================
# SF0013 — TC-13 / R041  E01 关键字段为空或格式异常
# ============================================================
function Generate-SF0013Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0013_测试数据.xlsx (TC-13 E01)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：3行，行2行号含字母，行3 WMS退单号为空
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000013"; wt $ws 2 2 "TK10000130"; wt $ws 2 3 "H000000013"; wt $ws 2 4 "00001"; wn $ws 2 5 3; bdr $ws 2 $h.Count
    wt $ws 3 1 "SF3190000000013"; wt $ws 3 2 "TK10000130"; wt $ws 3 3 "H000000013"; wt $ws 3 4 "1AB23"; wn $ws 3 5 2; bdr $ws 3 $h.Count
    wt $ws 4 1 "SF3190000000013"; wt $ws 4 2 ""; wt $ws 4 3 "H000000013"; wt $ws 4 4 "00003"; wn $ws 4 5 5; bdr $ws 4 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：1行，数量10（=退单合计，E08不触发）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000013"; wt $ws2 2 2 "H000000013"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 10
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000013" "TC-13" 200 "关闭" "不敏感" "2099/01/01"

    # 预期_汇总表：2行（TK10000130 和 [N/A]）
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000013" "TK10000130" "无法分配" "E01 - 关键字段为空或格式异常"
    summaryRow $wsSumm 3 "SF3190000000013" "[N/A]" "无法分配" "E01 - 关键字段为空或格式异常"

    # 预期_数据异常明细：2行
    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "退单表" 3 "SF3190000000013" "TK10000130" "H000000013" "行号" "1AB23" "E01" "行号格式异常：须为五位纯数字文本，当前含非数字字符"
    anomalyRow $wsAnomaly 3 "退单表" 4 "SF3190000000013" "[N/A]" "H000000013" "WMS退单号" "[N/A]" "E01" "关键字段为空"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0013_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0014 — TC-14 / R042  E02 行号重复
# ============================================================
function Generate-SF0014Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0014_测试数据.xlsx (TC-14 E02重复)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：2行，同一退单号行号均为 00001
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000014"; wt $ws 2 2 "TK10000140"; wt $ws 2 3 "H000000014"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    wt $ws 3 1 "SF3190000000014"; wt $ws 3 2 "TK10000140"; wt $ws 3 3 "H000000014"; wt $ws 3 4 "00001"; wn $ws 3 5 3; bdr $ws 3 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：1行，数量8（=退单合计）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000014"; wt $ws2 2 2 "H000000014"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 8
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000014" "TC-14" 200 "关闭" "不敏感" "2099/01/01"

    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000014" "TK10000140" "无法分配" "E02 - 退单表行号重复或不连续（行号 00001 在同一退单号下出现 2 次）"

    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "退单表" 2 "SF3190000000014" "TK10000140" "H000000014" "行号" "00001" "E02" "行号重复：00001 在同一退单号下出现 2 次"
    anomalyRow $wsAnomaly 3 "退单表" 3 "SF3190000000014" "TK10000140" "H000000014" "行号" "00001" "E02" "行号重复：00001 在同一退单号下出现 2 次"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0014_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0015 — TC-15 / R043  E03 QC情况非法
# ============================================================
function Generate-SF0015Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0015_测试数据.xlsx (TC-15 E03)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：1行
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000015"; wt $ws 2 2 "TK10000150"; wt $ws 2 3 "H000000015"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：QC情况 = XP（非法）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000015"; wt $ws2 2 2 "H000000015"; wt $ws2 2 3 "XP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000015" "TC-15" 200 "关闭" "不敏感" "2099/01/01"

    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000015" "TK10000150" "无法分配" "E03 - QC情况非法（须为 ZP/QC/NG）"

    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "质检库存表" 2 "SF3190000000015" "[N/A]" "H000000015" "QC情况" "XP" "E03" "QC情况非法：XP 不在合法值集合 {ZP, QC, NG} 内"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0015_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0056 — TC-16 / R044  E04 数量非法
# ============================================================
function Generate-SF0056Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0056_测试数据.xlsx (TC-16 E04)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：2行，行2数量=-1（负数）
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000056"; wt $ws 2 2 "TK10000560"; wt $ws 2 3 "H000000056"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    wt $ws 3 1 "SF3190000000056"; wt $ws 3 2 "TK10000560"; wt $ws 3 3 "H000000056"; wt $ws 3 4 "00002"; wn $ws 3 5 (-1); bdr $ws 3 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：数量5（=有效退单行合计，保持语义一致）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000056"; wt $ws2 2 2 "H000000056"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000056" "TC-16" 200 "关闭" "不敏感" "2099/01/01"

    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000056" "TK10000560" "无法分配" "E04 - 数量非法（须为正整数）"

    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "退单表" 3 "SF3190000000056" "TK10000560" "H000000056" "数量" "-1" "E04" "数量非法：须为正整数（>=1），当前值为 -1"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0056_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0057 — TC-17 / R045  E05 效期格式非法
# ============================================================
function Generate-SF0057Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0057_测试数据.xlsx (TC-17 E05)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：1行
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000057"; wt $ws 2 2 "TK10000570"; wt $ws 2 3 "H000000057"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：效期=2029/13/01（月份13非法）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000057"; wt $ws2 2 2 "H000000057"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/13/01"; wn $ws2 2 6 5
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000057" "TC-17" 200 "关闭" "不敏感" "2099/01/01"

    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000057" "TK10000570" "无法分配" "E05 - 效期格式非法（须为有效日期 YYYY/MM/DD）"

    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "质检库存表" 2 "SF3190000000057" "[N/A]" "H000000057" "效期" "2029/13/01" "E05" "效期格式非法：月份 13 超出范围（须为 1~12）"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0057_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0058 — TC-18 / R046  E06 物流单号仅在退单表
# ============================================================
function Generate-SF0058Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0058_测试数据.xlsx (TC-18 E06)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：1行
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000058"; wt $ws 2 2 "TK10000580"; wt $ws 2 3 "H000000058"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：空（仅有表头）—— SF0058 不在库存表，触发 E06
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000058" "TC-18" 200 "关闭" "不敏感" "2099/01/01"

    # 预期_汇总表
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000058" "TK10000580" "无法分配" "E06 - 物流单号仅存在于退单表（库存表无对应记录）"

    # E06 不产生数据异常明细，不加该 sheet

    while ($wb.Sheets.Count -gt 4) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0058_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0059 — TC-19 / R047  E07 物流单号仅在库存表
# ============================================================
function Generate-SF0059Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0059_测试数据.xlsx (TC-19 E07)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：空（仅有表头）—— SF0059 不在退单表，触发 E07
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：1行
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000059"; wt $ws2 2 2 "H000000059"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000059" "TC-19" 200 "关闭" "不敏感" "2099/01/01"

    # 预期_汇总表：WMS退单号 = [N/A]
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000059" "[N/A]" "无法分配" "E07 - 物流单号仅存在于质检库存表（退单表无对应记录）"

    # E07 不产生数据异常明细

    while ($wb.Sheets.Count -gt 4) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0059_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0060 — TC-20 / R048  E08 两表数量不一致
# ============================================================
function Generate-SF0060Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0060_测试数据.xlsx (TC-20 E08)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：2行（不同退单号），合计=8
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000060"; wt $ws 2 2 "TK10000600"; wt $ws 2 3 "H000000060"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    wt $ws 3 1 "SF3190000000060"; wt $ws 3 2 "TK10000601"; wt $ws 3 3 "H000000060"; wt $ws 3 4 "00001"; wn $ws 3 5 3; bdr $ws 3 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：合计=5（≠8，触发 E08）
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000060"; wt $ws2 2 2 "H000000060"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5
    bdr $ws2 2 $h2.Count; $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000060" "TC-20" 200 "关闭" "不敏感" "2099/01/01"

    # 预期_汇总表：2行（两个退单号均失败）
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000060" "TK10000600" "无法分配" "E08 - 同物流单号+SKU数量不一致（退单合计=8，库存合计=5）"
    summaryRow $wsSumm 3 "SF3190000000060" "TK10000601" "无法分配" "E08 - 同物流单号+SKU数量不一致（退单合计=8，库存合计=5）"

    # E08 不产生数据异常明细

    while ($wb.Sheets.Count -gt 4) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0060_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

# ============================================================
# SF0037 — TC-37 / R040  多错误码并存（E01+E04 / E08）
# ============================================================
function Generate-SF0037Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0037_测试数据.xlsx (TC-37 多错误码)..."
    $wb = $excel.Workbooks.Add()

    # 退单表：SF0037(2行) + SF0038(2行)
    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    # SF0037 行1：行号=X1234（E01格式非法），数量=5（合法）
    wt $ws 2 1 "SF3190000000037"; wt $ws 2 2 "TK10000370"; wt $ws 2 3 "H000000037"; wt $ws 2 4 "X1234"; wn $ws 2 5 5; bdr $ws 2 $h.Count
    # SF0037 行2：行号=00002（合法），数量=-2（E04非法）
    wt $ws 3 1 "SF3190000000037"; wt $ws 3 2 "TK10000370"; wt $ws 3 3 "H000000037"; wt $ws 3 4 "00002"; wn $ws 3 5 (-2); bdr $ws 3 $h.Count
    # SF0038 行1
    wt $ws 4 1 "SF3190000000038"; wt $ws 4 2 "TK10000380"; wt $ws 4 3 "H000000038"; wt $ws 4 4 "00001"; wn $ws 4 5 6; bdr $ws 4 $h.Count
    # SF0038 行2
    wt $ws 5 1 "SF3190000000038"; wt $ws 5 2 "TK10000381"; wt $ws 5 3 "H000000038"; wt $ws 5 4 "00001"; wn $ws 5 5 4; bdr $ws 5 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    # 质检库存表：SF0037(数量5，E08因E04整单跳过) + SF0038(数量6，<退单合计10，触发E08)
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000037"; wt $ws2 2 2 "H000000037"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5; bdr $ws2 2 $h2.Count
    wt $ws2 3 1 "SF3190000000038"; wt $ws2 3 2 "H000000038"; wt $ws2 3 3 "ZP"
    wt $ws2 3 4 "LA01"; wt $ws2 3 5 "2029/06/15"; wn $ws2 3 6 6; bdr $ws2 3 $h2.Count
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000037" "TC-37" 200 "关闭" "不敏感" "2099/01/01"

    # 预期_汇总表：3行
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000037" "TK10000370" "无法分配" "E01 - 关键字段为空或格式异常; E02 - 退单表行号重复或不连续; E04 - 数量非法"
    summaryRow $wsSumm 3 "SF3190000000038" "TK10000380" "无法分配" "E08 - 同物流单号+SKU数量不一致（退单合计=10，库存合计=6）"
    summaryRow $wsSumm 4 "SF3190000000038" "TK10000381" "无法分配" "E08 - 同物流单号+SKU数量不一致（退单合计=10，库存合计=6）"

    # 预期_数据异常明细：E01+E02+E04，E08不产生
    $wsAnomaly = anomalyHeader $wb
    anomalyRow $wsAnomaly 2 "退单表" 2 "SF3190000000037" "TK10000370" "H000000037" "行号" "X1234" "E01" "行号格式异常：须为五位纯数字文本，当前含非数字字符"
    anomalyRow $wsAnomaly 3 "退单表" 3 "SF3190000000037" "TK10000370" "H000000037" "行号" "00002" "E02" "行号不从 00001 起：当前序列首行为 00002"
    anomalyRow $wsAnomaly 4 "退单表" 3 "SF3190000000037" "TK10000370" "H000000037" "数量" "-2" "E04" "数量非法：须为正整数（>=1），当前值为 -2"

    while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $path = "$outputDir\SF0037_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

function Generate-SF0036Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0036_测试数据.xlsx ..."
    $wb = $excel.Workbooks.Add()

    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000036"; wt $ws 2 2 "TK00000036"; wt $ws 2 3 "H000000001"; wt $ws 2 4 "00001"; wn $ws 2 5 2
    bdr $ws 2 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000036"; wt $ws2 2 2 "H000000001"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/01/01"; wn $ws2 2 6 1
    bdr $ws2 2 $h2.Count
    wt $ws2 3 1 "SF3190000000036"; wt $ws2 3 2 "H000000001"; wt $ws2 3 3 "QC"
    wt $ws2 3 4 "LA01"; wt $ws2 3 5 "2029/01/01"; wn $ws2 3 6 1
    bdr $ws2 3 $h2.Count
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000036" "TC-36" 200 "简版" "不敏感" "2099/01/01"
    summary $wb "SF3190000000036" "TK00000036" "无法分配"
    $ws3 = $wb.Sheets | Where-Object { $_.Name -eq "预期_汇总表" }
    wt $ws3 2 4 "E11 - QC库存碎片无法分配（0 < T < groupMinQty）"

    while ($wb.Sheets.Count -gt 4) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $p36 = "$outputDir\SF0036_测试数据.xlsx"
    $wb.SaveAs($p36, 51); $wb.Close($false)
    Write-Host "  已保存: $p36"
    return $p36
}

function Generate-SF0055Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0055_测试数据.xlsx ..."
    $wb = $excel.Workbooks.Add()

    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
    wt $ws 2 1 "SF3190000000055"; wt $ws 2 2 "TK10000550"; wt $ws 2 3 "H000000055"; wt $ws 2 4 "00001"; wn $ws 2 5 5
    bdr $ws 2 $h.Count
    wt $ws 3 1 "SF3190000000056"; wt $ws 3 2 "TK10000550"; wt $ws 3 3 "H000000055"; wt $ws 3 4 "00001"; wn $ws 3 5 3
    bdr $ws 3 $h.Count
    $ws.Columns.AutoFit() | Out-Null

    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    $h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
    hdr $ws2 1 $h2.Count 0xD9E1F2
    wt $ws2 2 1 "SF3190000000055"; wt $ws2 2 2 "H000000055"; wt $ws2 2 3 "ZP"
    wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 8
    bdr $ws2 2 $h2.Count
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000055" "TC-55" 200 "关闭" "不敏感" "2099/01/01" "E12-② 跨物流单号重复"
    assertionSheet $wb "E12" "TK10000550" "SF3190000000055"

    while ($wb.Sheets.Count -gt 4) { $wb.Sheets($wb.Sheets.Count).Delete() }
    $p55 = "$outputDir\SF0055_测试数据.xlsx"
    $wb.SaveAs($p55, 51); $wb.Close($false)
    Write-Host "  已保存: $p55"
    return $p55
}

# ---------- 辅助：写汇总 sheet ----------
function summary($wb, $sfNum, $wmsNum, $status) {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_汇总表"
    $cols = @("物流单号","WMS退单号","退单号状态","原因")
    for ($c = 1; $c -le $cols.Count; $c++) { wt $ws 1 $c ($cols[$c-1]) }
    hdr $ws 1 $cols.Count 0xE2EFDA
    wt $ws 2 1 $sfNum; wt $ws 2 2 $wmsNum; wt $ws 2 3 $status; wt $ws 2 4 ""
    bdr $ws 2 $cols.Count
    $ws.Columns.AutoFit() | Out-Null
}

function detailHeader($wb) {
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws.Name = "预期_成功分配明细"
    $h = @("物流单号","WMS退单号","SKU","行号","退单数量","QC情况","批号","效期","分配数量","行状态","退单号状态")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xE2EFDA
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}

function returnHeader($ws) {
    $h = @("物流单号","WMS退单号","SKU","行号","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
}

function inventoryHeader($ws) {
    $h = @("物流单号","SKU","QC情况","批号","效期","数量")
    for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
    hdr $ws 1 $h.Count 0xD9E1F2
}

function Generate-SF0062Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0062_测试数据.xlsx (TC-40 预检测B)..."
    $wb = $excel.Workbooks.Add()

    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    returnHeader $ws
    wt $ws 2 1 "SF3190000000062"; wt $ws 2 2 "TK10000620"; wt $ws 2 3 "H000000062"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 5
    wt $ws 3 1 "SF3190000000062"; wt $ws 3 2 "TK10000620"; wt $ws 3 3 "H000000062"; wt $ws 3 4 "00002"; wn $ws 3 5 6; bdr $ws 3 5
    wt $ws 4 1 "SF3190000000062"; wt $ws 4 2 "TK10000620"; wt $ws 4 3 "H000000062"; wt $ws 4 4 "00003"; wn $ws 4 5 1; bdr $ws 4 5
    $ws.Columns.AutoFit() | Out-Null

    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    inventoryHeader $ws2
    wt $ws2 2 1 "SF3190000000062"; wt $ws2 2 2 "H000000062"; wt $ws2 2 3 "ZP"; wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 7; bdr $ws2 2 6
    wt $ws2 3 1 "SF3190000000062"; wt $ws2 3 2 "H000000062"; wt $ws2 3 3 "QC"; wt $ws2 3 4 "LA01"; wt $ws2 3 5 "2029/06/15"; wn $ws2 3 6 4; bdr $ws2 3 6
    wt $ws2 4 1 "SF3190000000062"; wt $ws2 4 2 "H000000062"; wt $ws2 4 3 "NG"; wt $ws2 4 4 "LA01"; wt $ws2 4 5 "2029/06/15"; wn $ws2 4 6 1; bdr $ws2 4 6
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000062" "TC-40" 200 "简版" "不敏感" "2099/01/01"
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000062" "TK10000620" "无法分配" "E09 - 分配路径穷尽"
    detailHeader $wb | Out-Null
    $wsDbg = debugLogHeader $wb
    debugLogRow $wsDbg 2 "SF3190000000062" "H000000062" "TK10000620" "00001" 5 1 "-" 1 "" "-" "-" "-" "-" "-" "否" 0 "失败" "E09" "预检测B（强制竞争库存不足）"
    debugLogRow $wsDbg 3 "SF3190000000062" "H000000062" "TK10000620" "00002" 6 2 "-" 1 "" "-" "-" "-" "-" "-" "否" 0 "失败" "E09" "预检测B（强制竞争库存不足）"
    debugLogRow $wsDbg 4 "SF3190000000062" "H000000062" "TK10000620" "00003" 1 3 "-" 2 "" "-" "-" "-" "-" "-" "否" 0 "失败" "E09" "预检测B（强制竞争库存不足）"

    $path = "$outputDir\SF0062_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

function Generate-SF0063Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0063_测试数据.xlsx (TC-41 数据异常明细格式)..."
    $wb = $excel.Workbooks.Add()

    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    returnHeader $ws
    wt $ws 2 1 "SF3190000000063"; wt $ws 2 2 "TK10000630"; wt $ws 2 3 "H000000063"; wt $ws 2 4 "00001"; wn $ws 2 5 3; bdr $ws 2 5
    wt $ws 3 1 "SF3190000000063"; wt $ws 3 2 "TK10000630"; wt $ws 3 3 "H000000063"; wt $ws 3 4 "X9999"; wn $ws 3 5 2; bdr $ws 3 5
    wt $ws 4 1 "SF3190000000063"; wt $ws 4 2 ""; wt $ws 4 3 "H000000063"; wt $ws 4 4 "00003"; wn $ws 4 5 1; bdr $ws 4 5
    wt $ws 5 1 "SF3190000000063"; wt $ws 5 2 "TK10000630"; wt $ws 5 3 "H000000063"; wt $ws 5 4 "00004"; wn $ws 5 5 -1; bdr $ws 5 5
    $ws.Columns.AutoFit() | Out-Null

    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    inventoryHeader $ws2
    wt $ws2 2 1 "SF3190000000063"; wt $ws2 2 2 "H000000063"; wt $ws2 2 3 "XP"; wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5; bdr $ws2 2 6
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000063" "TC-41" 200 "关闭" "不敏感" "2099/01/01"
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000063" "TK10000630" "无法分配" "E01 - 关键字段为空或格式异常; E02 - 退单表行号重复或不连续; E03 - QC情况非法; E04 - 数量非法"
    summaryRow $wsSumm 3 "SF3190000000063" "[N/A]" "无法分配" "E01 - 关键字段为空或格式异常; E03 - QC情况非法"

    $wsAnom = anomalyHeader $wb
    anomalyRow $wsAnom 2 "退单表" 3 "SF3190000000063" "TK10000630" "H000000063" "行号" "X9999" "E01" "行号格式异常：须为五位纯数字文本"
    anomalyRow $wsAnom 3 "退单表" 4 "SF3190000000063" "[N/A]" "H000000063" "WMS退单号" "[N/A]" "E01" "关键字段为空"
    anomalyRow $wsAnom 4 "退单表" 5 "SF3190000000063" "TK10000630" "H000000063" "数量" "-1" "E04" "数量非法：须为正整数"
    anomalyRow $wsAnom 5 "质检库存表" 2 "SF3190000000063" "[N/A]" "H000000063" "QC情况" "XP" "E03" "QC情况非法：须为 ZP/QC/NG"
    anomalyRow $wsAnom 6 "退单表" 2 "SF3190000000063" "TK10000630" "H000000063" "行号" "00001" "E02" "行号不连续：当前序列为 00001、00004"
    anomalyRow $wsAnom 7 "退单表" 5 "SF3190000000063" "TK10000630" "H000000063" "行号" "00004" "E02" "行号不连续：当前序列为 00001、00004"

    detailHeader $wb | Out-Null
    $path = "$outputDir\SF0063_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

function Generate-SF0064Workbook($excel, $outputDir) {
    Write-Host "`n生成 SF0064_测试数据.xlsx (TC-28 多错误码原因格式)..."
    $wb = $excel.Workbooks.Add()

    $ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
    returnHeader $ws
    wt $ws 2 1 "SF3190000000064"; wt $ws 2 2 "TK10000640"; wt $ws 2 3 "H000000064"; wt $ws 2 4 "X1234"; wn $ws 2 5 5; bdr $ws 2 5
    wt $ws 3 1 "SF3190000000064"; wt $ws 3 2 "TK10000640"; wt $ws 3 3 "H000000064"; wt $ws 3 4 "00002"; wn $ws 3 5 -3; bdr $ws 3 5
    wt $ws 4 1 "SF3190000000064"; wt $ws 4 2 "TK10000640"; wt $ws 4 3 "H000000064"; wt $ws 4 4 "00003"; wn $ws 4 5 2; bdr $ws 4 5
    $ws.Columns.AutoFit() | Out-Null

    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
    $ws2.Name = "输入_质检库存表"
    inventoryHeader $ws2
    wt $ws2 2 1 "SF3190000000064"; wt $ws2 2 2 "H000000064"; wt $ws2 2 3 "XP"; wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 5; bdr $ws2 2 6
    $ws2.Columns.AutoFit() | Out-Null

    config $wb "SF3190000000064" "TC-28" 200 "关闭" "不敏感" "2099/01/01"
    $wsSumm = summaryHeader $wb
    summaryRow $wsSumm 2 "SF3190000000064" "TK10000640" "无法分配" "E01 - 关键字段为空或格式异常; E02 - 退单表行号重复或不连续; E03 - QC情况非法; E04 - 数量非法"

    $wsAnom = anomalyHeader $wb
    anomalyRow $wsAnom 2 "退单表" 2 "SF3190000000064" "TK10000640" "H000000064" "行号" "X1234" "E01" "行号格式异常：须为五位纯数字文本"
    anomalyRow $wsAnom 3 "退单表" 3 "SF3190000000064" "TK10000640" "H000000064" "数量" "-3" "E04" "数量非法：须为正整数"
    anomalyRow $wsAnom 4 "质检库存表" 2 "SF3190000000064" "[N/A]" "H000000064" "QC情况" "XP" "E03" "QC情况非法：须为 ZP/QC/NG"
    anomalyRow $wsAnom 5 "退单表" 3 "SF3190000000064" "TK10000640" "H000000064" "行号" "00002" "E02" "行号不从 00001 起：当前序列首行为 00002"
    anomalyRow $wsAnom 6 "退单表" 4 "SF3190000000064" "TK10000640" "H000000064" "行号" "00003" "E02" "行号不从 00001 起：当前序列首行为 00002"

    detailHeader $wb | Out-Null
    $path = "$outputDir\SF0064_测试数据.xlsx"
    $wb.SaveAs($path, 51); $wb.Close($false)
    Write-Host "  已保存: $path"
    return $path
}

function Rename-LegacyExpectedSheets {
    param([string]$Dir)
    $ex = New-Object -ComObject Excel.Application
    $ex.Visible = $false; $ex.DisplayAlerts = $false
    Write-Host "统一预期 Sheet 命名..."
    Get-ChildItem -Path $Dir -Filter "SF*_*.xlsx" | ForEach-Object {
        $wb = $ex.Workbooks.Open($_.FullName)
        foreach ($sheet in $wb.Worksheets) {
            if ($sheet.Name -match "分配状态") {
                $old = $sheet.Name
                $sheet.Name = $old -replace "分配状态汇总", "汇总表"
                Write-Host "  $($_.Name): $old -> $($sheet.Name)"
            }
        }
        $wb.Save()
        $wb.Close($false)
    }
    $ex.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ex) | Out-Null
    Write-Host "Sheet 命名统一完成。"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false

if ($Target -eq "RenameSheets") {
    Rename-LegacyExpectedSheets $outputDir
    exit 0
}

if ($Target -eq "OnlySF0055") {
    Write-Host "仅生成 SF0055_测试数据.xlsx ..."
    $p55 = Generate-SF0055Workbook $excel $outputDir
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "SF0055 生成完成: $p55"
    exit 0
}

if ($Target -eq "OnlySF0036") {
    Write-Host "仅生成 SF0036_测试数据.xlsx ..."
    $p36 = Generate-SF0036Workbook $excel $outputDir
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "SF0036 生成完成: $p36"
    exit 0
}

if ($Target -eq "OnlyPendingTC") {
    Write-Host "仅生成待自动化 TC DataSet（SF0062/SF0063/SF0064）..."
    $p62 = Generate-SF0062Workbook $excel $outputDir
    $p63 = Generate-SF0063Workbook $excel $outputDir
    $p64 = Generate-SF0064Workbook $excel $outputDir
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "`n待自动化 TC DataSet 生成完成，共 3 个工作簿"
    Write-Host "  $p62"; Write-Host "  $p63"; Write-Host "  $p64"
    exit 0
}

if ($Target -eq "OnlyM05") {
    Write-Host "仅生成 M05 缺口 DataSet（TC-13~20、TC-37）..."
    $p13 = Generate-SF0013Workbook $excel $outputDir
    $p14 = Generate-SF0014Workbook $excel $outputDir
    $p15 = Generate-SF0015Workbook $excel $outputDir
    $p56 = Generate-SF0056Workbook $excel $outputDir
    $p57 = Generate-SF0057Workbook $excel $outputDir
    $p58 = Generate-SF0058Workbook $excel $outputDir
    $p59 = Generate-SF0059Workbook $excel $outputDir
    $p60 = Generate-SF0060Workbook $excel $outputDir
    $p37 = Generate-SF0037Workbook $excel $outputDir
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "`nM05 DataSet 生成完成，共 9 个工作簿"
    Write-Host "  $p13"; Write-Host "  $p14"; Write-Host "  $p15"
    Write-Host "  $p56"; Write-Host "  $p57"; Write-Host "  $p58"
    Write-Host "  $p59"; Write-Host "  $p60"; Write-Host "  $p37"
    exit 0
}

Write-Host "Excel COM 已启动，开始生成工作簿..."

# ============================================================
# SF0032 — TC-32 / R014 无保质期哨兵效期
# ============================================================
Write-Host "`n[1/5] 生成 SF0032_测试数据.xlsx ..."
$wb = $excel.Workbooks.Add()

# 退单表
$ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
$h = @("物流单号","WMS退单号","SKU","行号","数量")
for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
hdr $ws 1 $h.Count 0xD9E1F2
wt $ws 2 1 "SF3190000000032"; wt $ws 2 2 "TK10000320"; wt $ws 2 3 "H000000032"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
wt $ws 3 1 "SF3190000000032"; wt $ws 3 2 "TK10000320"; wt $ws 3 3 "H000000032"; wt $ws 3 4 "00002"; wn $ws 3 5 5; bdr $ws 3 $h.Count
$ws.Columns.AutoFit() | Out-Null

# 质检库存表
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws2.Name = "输入_质检库存表"
$h2 = @("物流单号","SKU","QC情况","批号","效期","数量","DataSet说明")
for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
hdr $ws2 1 $h2.Count 0xD9E1F2
wt $ws2 2 1 "SF3190000000032"; wt $ws2 2 2 "H000000032"; wt $ws2 2 3 "ZP"
wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2099/01/01"; wn $ws2 2 6 10
wt $ws2 2 7 "该 SKU 为无保质期商品，库存统一使用哨兵效期 2099/01/01（R014）"
bdr $ws2 2 $h2.Count
$ws2.Columns.AutoFit() | Out-Null

# 配置
config $wb "SF3190000000032" "TC-32" 200 "关闭" "不敏感" "2099/01/01"

# 汇总表
summary $wb "SF3190000000032" "TK10000320" "批量导入"

# 分配明细
$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws5.Name = "预期_成功分配明细"
$h5 = @("物流单号","WMS退单号","SKU","行号","退单数量","QC情况","批号","效期","分配数量","行状态","退单号状态")
for ($c = 1; $c -le $h5.Count; $c++) { wt $ws5 1 $c ($h5[$c-1]) }
hdr $ws5 1 $h5.Count 0xE2EFDA
# 行1：哨兵效期 2099/01/01
wt $ws5 2 1 "SF3190000000032"; wt $ws5 2 2 "TK10000320"; wt $ws5 2 3 "H000000032"
wt $ws5 2 4 "00001"; wn $ws5 2 5 5; wt $ws5 2 6 "ZP"; wt $ws5 2 7 "LA01"
wt $ws5 2 8 "2099/01/01"; wn $ws5 2 9 5; wt $ws5 2 10 "批量导入"; wt $ws5 2 11 "批量导入"
bdr $ws5 2 $h5.Count
# 行2：同一无保质期商品继续使用哨兵效期
wt $ws5 3 1 "SF3190000000032"; wt $ws5 3 2 "TK10000320"; wt $ws5 3 3 "H000000032"
wt $ws5 3 4 "00002"; wn $ws5 3 5 5; wt $ws5 3 6 "ZP"; wt $ws5 3 7 "LA01"
wt $ws5 3 8 "2099/01/01"; wn $ws5 3 9 5; wt $ws5 3 10 "批量导入"; wt $ws5 3 11 "批量导入"
bdr $ws5 3 $h5.Count
$ws5.Columns.AutoFit() | Out-Null

while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
$p32 = "$outputDir\SF0032_测试数据.xlsx"
$wb.SaveAs($p32, 51); $wb.Close($false)
Write-Host "  已保存: $p32"

# ============================================================
# SF0046 — TC-46 / R021 批号大小写标准化
# ============================================================
Write-Host "`n[2/5] 生成 SF0046_测试数据.xlsx ..."
$wb = $excel.Workbooks.Add()

$ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
$h = @("物流单号","WMS退单号","SKU","行号","数量")
for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
hdr $ws 1 $h.Count 0xD9E1F2
wt $ws 2 1 "SF3190000000046"; wt $ws 2 2 "TK10000460"; wt $ws 2 3 "H000000046"; wt $ws 2 4 "00001"; wn $ws 2 5 5; bdr $ws 2 $h.Count
wt $ws 3 1 "SF3190000000046"; wt $ws 3 2 "TK10000460"; wt $ws 3 3 "H000000046"; wt $ws 3 4 "00002"; wn $ws 3 5 5; bdr $ws 3 $h.Count
$ws.Columns.AutoFit() | Out-Null

$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws2.Name = "输入_质检库存表"
$h2 = @("物流单号","SKU","QC情况","批号","效期","数量","DataSet说明")
for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
hdr $ws2 1 $h2.Count 0xD9E1F2
# 行1：小写批号 a01
wt $ws2 2 1 "SF3190000000046"; wt $ws2 2 2 "H000000046"; wt $ws2 2 3 "ZP"
wt $ws2 2 4 "a01"   # ← 小写批号，R021 UCase 后 = A01
wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 6
wt $ws2 2 7 "小写批号 a01，R021 UCase 后=A01，与下行合并为同一五元组 T=10"
bdr $ws2 2 $h2.Count
# 行2：大写批号 A01
wt $ws2 3 1 "SF3190000000046"; wt $ws2 3 2 "H000000046"; wt $ws2 3 3 "ZP"
wt $ws2 3 4 "A01"   # ← 大写批号
wt $ws2 3 5 "2029/06/15"; wn $ws2 3 6 4
wt $ws2 3 7 "大写批号 A01，与上行标准化后相同，合并 T=6+4=10"
bdr $ws2 3 $h2.Count
$ws2.Columns.AutoFit() | Out-Null

config $wb "SF3190000000046" "TC-46" 200 "关闭" "不敏感" "2099/01/01"
summary $wb "SF3190000000046" "TK10000460" "批量导入"

$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws5.Name = "预期_成功分配明细"
$h5 = @("物流单号","WMS退单号","SKU","行号","退单数量","QC情况","批号","效期","分配数量","行状态","退单号状态")
for ($c = 1; $c -le $h5.Count; $c++) { wt $ws5 1 $c ($h5[$c-1]) }
hdr $ws5 1 $h5.Count 0xE2EFDA
wt $ws5 2 1 "SF3190000000046"; wt $ws5 2 2 "TK10000460"; wt $ws5 2 3 "H000000046"
wt $ws5 2 4 "00001"; wn $ws5 2 5 5; wt $ws5 2 6 "ZP"; wt $ws5 2 7 "A01"
wt $ws5 2 8 "2029/06/15"; wn $ws5 2 9 5; wt $ws5 2 10 "批量导入"; wt $ws5 2 11 "批量导入"
bdr $ws5 2 $h5.Count
wt $ws5 3 1 "SF3190000000046"; wt $ws5 3 2 "TK10000460"; wt $ws5 3 3 "H000000046"
wt $ws5 3 4 "00002"; wn $ws5 3 5 5; wt $ws5 3 6 "ZP"; wt $ws5 3 7 "A01"
wt $ws5 3 8 "2029/06/15"; wn $ws5 3 9 5; wt $ws5 3 10 "批量导入"; wt $ws5 3 11 "批量导入"
bdr $ws5 3 $h5.Count
$ws5.Columns.AutoFit() | Out-Null

while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
$p46 = "$outputDir\SF0046_测试数据.xlsx"
$wb.SaveAs($p46, 51); $wb.Close($false)
Write-Host "  已保存: $p46"

# ============================================================
# SF0047 — TC-47 / R022 QC大小写标准化
# ============================================================
Write-Host "`n[3/5] 生成 SF0047_测试数据.xlsx ..."
$wb = $excel.Workbooks.Add()

$ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
$h = @("物流单号","WMS退单号","SKU","行号","数量")
for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
hdr $ws 1 $h.Count 0xD9E1F2
wt $ws 2 1 "SF3190000000047"; wt $ws 2 2 "TK10000470"; wt $ws 2 3 "H000000047"; wt $ws 2 4 "00001"; wn $ws 2 5 3; bdr $ws 2 $h.Count
wt $ws 3 1 "SF3190000000047"; wt $ws 3 2 "TK10000470"; wt $ws 3 3 "H000000047"; wt $ws 3 4 "00002"; wn $ws 3 5 2; bdr $ws 3 $h.Count
$ws.Columns.AutoFit() | Out-Null

$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws2.Name = "输入_质检库存表"
$h2 = @("物流单号","SKU","QC情况","批号","效期","数量","DataSet说明")
for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
hdr $ws2 1 $h2.Count 0xD9E1F2
# 行1：小写 QC zp
wt $ws2 2 1 "SF3190000000047"; wt $ws2 2 2 "H000000047"
wt $ws2 2 3 "zp"    # ← 全小写 QC，R022 标准化后=ZP
wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 3
wt $ws2 2 7 "小写 QC zp，R022 Trim->UCase->校验 -> ZP，E03 不触发"
bdr $ws2 2 $h2.Count
# 行2：混合大小写 QC Qc
wt $ws2 3 1 "SF3190000000047"; wt $ws2 3 2 "H000000047"
wt $ws2 3 3 "Qc"    # ← 混合大小写，R022 标准化后=QC
wt $ws2 3 4 "LA02"; wt $ws2 3 5 "2029/06/15"; wn $ws2 3 6 2
wt $ws2 3 7 "混合大小写 Qc，R022 Trim->UCase->校验 -> QC，E03 不触发"
bdr $ws2 3 $h2.Count
$ws2.Columns.AutoFit() | Out-Null

config $wb "SF3190000000047" "TC-47" 200 "关闭" "不敏感" "2099/01/01"
summary $wb "SF3190000000047" "TK10000470" "批量导入"

$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws5.Name = "预期_成功分配明细"
$h5 = @("物流单号","WMS退单号","SKU","行号","退单数量","QC情况","批号","效期","分配数量","行状态","退单号状态")
for ($c = 1; $c -le $h5.Count; $c++) { wt $ws5 1 $c ($h5[$c-1]) }
hdr $ws5 1 $h5.Count 0xE2EFDA
wt $ws5 2 1 "SF3190000000047"; wt $ws5 2 2 "TK10000470"; wt $ws5 2 3 "H000000047"
wt $ws5 2 4 "00001"; wn $ws5 2 5 3; wt $ws5 2 6 "ZP"; wt $ws5 2 7 "LA01"
wt $ws5 2 8 "2029/06/15"; wn $ws5 2 9 3; wt $ws5 2 10 "批量导入"; wt $ws5 2 11 "批量导入"
bdr $ws5 2 $h5.Count
wt $ws5 3 1 "SF3190000000047"; wt $ws5 3 2 "TK10000470"; wt $ws5 3 3 "H000000047"
wt $ws5 3 4 "00002"; wn $ws5 3 5 2; wt $ws5 3 6 "QC"; wt $ws5 3 7 "LA02"
wt $ws5 3 8 "2029/06/15"; wn $ws5 3 9 2; wt $ws5 3 10 "批量导入"; wt $ws5 3 11 "批量导入"
bdr $ws5 3 $h5.Count
$ws5.Columns.AutoFit() | Out-Null

while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
$p47 = "$outputDir\SF0047_测试数据.xlsx"
$wb.SaveAs($p47, 51); $wb.Close($false)
Write-Host "  已保存: $p47"

# ============================================================
# SF0048 — TC-48 / R011 行号文本格式合法性测试
# 验证：文本型五位前导零行号（00001/00002/00003）正确读入；连续性校验通过，E02不触发
# ============================================================
Write-Host "`n[4/5] 生成 SF0048_测试数据.xlsx ..."
$wb = $excel.Workbooks.Add()

# 退单表（关键：行号为文本型，直接填写五位前导零字符串）
$ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
$h = @("物流单号","WMS退单号","SKU","行号","数量")
for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
hdr $ws 1 $h.Count 0xD9E1F2
wt $ws 2 1 "SF3190000000048"; wt $ws 2 2 "TK10000480"; wt $ws 2 3 "H000000048"
wt $ws 2 4 "00001"   # ← 文本型，个位数（1）对应的五位前导零
wn $ws 2 5 5
bdr $ws 2 $h.Count
wt $ws 3 1 "SF3190000000048"; wt $ws 3 2 "TK10000480"; wt $ws 3 3 "H000000048"
wt $ws 3 4 "00002"   # ← 文本型，连续行号第2行
wn $ws 3 5 5
bdr $ws 3 $h.Count
wt $ws 4 1 "SF3190000000048"; wt $ws 4 2 "TK10000480"; wt $ws 4 3 "H000000048"
wt $ws 4 4 "00003"   # ← 文本型，连续行号第3行；三行00001/00002/00003从00001起连续，E02不触发
wn $ws 4 5 5
bdr $ws 4 $h.Count
$ws.Cells(2, 3).AddComment("DataSet关键：文本型行号，VarType=vbString，CStr直接读取=00001，满足5位全数字，E01不触发") | Out-Null
$ws.Cells(3, 3).AddComment("连续行号00002，与00001/00003构成从00001起的连续序列，E02不触发") | Out-Null
$ws.Cells(4, 3).AddComment("连续行号00003，三行00001/00002/00003均合法，正向基线验证完毕") | Out-Null
$ws.Columns.AutoFit() | Out-Null

# 质检库存表（数量=15，对应退单合计5+5+5）
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws2.Name = "输入_质检库存表"
$h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
hdr $ws2 1 $h2.Count 0xD9E1F2
wt $ws2 2 1 "SF3190000000048"; wt $ws2 2 2 "H000000048"; wt $ws2 2 3 "ZP"
wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 15
bdr $ws2 2 $h2.Count
$ws2.Columns.AutoFit() | Out-Null

config $wb "SF3190000000048" "TC-48" 200 "关闭" "不敏感" "2099/01/01"
summary $wb "SF3190000000048" "TK10000480" "批量导入"

$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws5.Name = "预期_成功分配明细"
$h5 = @("物流单号","WMS退单号","SKU","行号","退单数量","QC情况","批号","效期","分配数量","行状态","退单号状态")
for ($c = 1; $c -le $h5.Count; $c++) { wt $ws5 1 $c ($h5[$c-1]) }
hdr $ws5 1 $h5.Count 0xE2EFDA
wt $ws5 2 1 "SF3190000000048"; wt $ws5 2 2 "TK10000480"; wt $ws5 2 3 "H000000048"
wt $ws5 2 4 "00001"; wn $ws5 2 5 5; wt $ws5 2 6 "ZP"; wt $ws5 2 7 "LA01"
wt $ws5 2 8 "2029/06/15"; wn $ws5 2 9 5; wt $ws5 2 10 "批量导入"; wt $ws5 2 11 "批量导入"
bdr $ws5 2 $h5.Count
wt $ws5 3 1 "SF3190000000048"; wt $ws5 3 2 "TK10000480"; wt $ws5 3 3 "H000000048"
wt $ws5 3 4 "00002"; wn $ws5 3 5 5; wt $ws5 3 6 "ZP"; wt $ws5 3 7 "LA01"
wt $ws5 3 8 "2029/06/15"; wn $ws5 3 9 5; wt $ws5 3 10 "批量导入"; wt $ws5 3 11 "批量导入"
bdr $ws5 3 $h5.Count
wt $ws5 4 1 "SF3190000000048"; wt $ws5 4 2 "TK10000480"; wt $ws5 4 3 "H000000048"
wt $ws5 4 4 "00003"; wn $ws5 4 5 5; wt $ws5 4 6 "ZP"; wt $ws5 4 7 "LA01"
wt $ws5 4 8 "2029/06/15"; wn $ws5 4 9 5; wt $ws5 4 10 "批量导入"; wt $ws5 4 11 "批量导入"
bdr $ws5 4 $h5.Count
$ws5.Columns.AutoFit() | Out-Null

while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
$p48 = "$outputDir\SF0048_测试数据.xlsx"
$wb.SaveAs($p48, 51); $wb.Close($false)
Write-Host "  已保存: $p48"

# ============================================================
# SF0051 — TC-51 / R042 行号跳号/不连续 → E02 负向测试
# 覆盖：① 有跳号（TK10000510: 00001→00003, 缺00002）；② 不从00001起（TK10000511: 00002→00003）
# ============================================================
Write-Host "`n[5/5] 生成 SF0051_测试数据.xlsx ..."
$wb = $excel.Workbooks.Add()

# 退单表（两个WMS退单号，各触发E02）
$ws = $wb.Sheets(1); $ws.Name = "输入_退单表"
$h = @("物流单号","WMS退单号","SKU","行号","数量")
for ($c = 1; $c -le $h.Count; $c++) { wt $ws 1 $c ($h[$c-1]) }
hdr $ws 1 $h.Count 0xD9E1F2
# TK10000510 — 跳号：00001, 00003（缺00002）
wt $ws 2 1 "SF3190000000051"; wt $ws 2 2 "TK10000510"; wt $ws 2 3 "H000000051"; wt $ws 2 4 "00001"; wn $ws 2 5 5
bdr $ws 2 $h.Count
wt $ws 3 1 "SF3190000000051"; wt $ws 3 2 "TK10000510"; wt $ws 3 3 "H000000051"; wt $ws 3 4 "00003"; wn $ws 3 5 5
bdr $ws 3 $h.Count
# TK10000511 — 不从00001起：00002, 00003
wt $ws 4 1 "SF3190000000051"; wt $ws 4 2 "TK10000511"; wt $ws 4 3 "H000000051"; wt $ws 4 4 "00002"; wn $ws 4 5 3
bdr $ws 4 $h.Count
wt $ws 5 1 "SF3190000000051"; wt $ws 5 2 "TK10000511"; wt $ws 5 3 "H000000051"; wt $ws 5 4 "00003"; wn $ws 5 5 3
bdr $ws 5 $h.Count
$ws.Cells(3, 3).AddComment("跳号：TK10000510行号00001→00003，缺00002，触发E02") | Out-Null
$ws.Cells(4, 3).AddComment("不从00001起：TK10000511行号从00002开始，缺00001，触发E02") | Out-Null
$ws.Columns.AutoFit() | Out-Null

# 质检库存表（提供匹配库存，避免E06/E07干扰E02验证）
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws2.Name = "输入_质检库存表"
$h2 = @("物流单号","SKU","QC情况","批号","效期","数量")
for ($c = 1; $c -le $h2.Count; $c++) { wt $ws2 1 $c ($h2[$c-1]) }
hdr $ws2 1 $h2.Count 0xD9E1F2
wt $ws2 2 1 "SF3190000000051"; wt $ws2 2 2 "H000000051"; wt $ws2 2 3 "ZP"
wt $ws2 2 4 "LA01"; wt $ws2 2 5 "2029/06/15"; wn $ws2 2 6 16
bdr $ws2 2 $h2.Count
$ws2.Columns.AutoFit() | Out-Null

# 配置页
config $wb "SF3190000000051" "TC-51" 200 "关闭" "不敏感" "2099/01/01"

# 预期_汇总表（两个退单号均E02失败，无成功分配明细）
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws3.Name = "预期_汇总表"
$h3 = @("物流单号","WMS退单号","退单号状态","原因")
for ($c = 1; $c -le $h3.Count; $c++) { wt $ws3 1 $c ($h3[$c-1]) }
hdr $ws3 1 $h3.Count 0xFFF2CC
wt $ws3 2 1 "SF3190000000051"; wt $ws3 2 2 "TK10000510"; wt $ws3 2 3 "无法分配"; wt $ws3 2 4 "E02 - 退单表行号重复或不连续（00001→00003，缺00002）"
bdr $ws3 2 $h3.Count
wt $ws3 3 1 "SF3190000000051"; wt $ws3 3 2 "TK10000511"; wt $ws3 3 3 "无法分配"; wt $ws3 3 4 "E02 - 退单表行号重复或不连续（行号不从00001起，首行为00002）"
bdr $ws3 3 $h3.Count
$ws3.Columns.AutoFit() | Out-Null

# 预期_数据异常明细（E02不连续子情形记录相关退单行）
$ws4 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count))
$ws4.Name = "预期_数据异常明细"
$h4 = @("来源表","原始行号","物流单号","WMS退单号","SKU","异常字段名","原始值","错误码","原因说明")
for ($c = 1; $c -le $h4.Count; $c++) { wt $ws4 1 $c ($h4[$c-1]) }
hdr $ws4 1 $h4.Count 0xFCE4D6
wt $ws4 2 1 "退单表"; wn $ws4 2 2 2; wt $ws4 2 3 "SF3190000000051"; wt $ws4 2 4 "TK10000510"; wt $ws4 2 5 "H000000051"; wt $ws4 2 6 "行号"; wt $ws4 2 7 "00001"; wt $ws4 2 8 "E02"; wt $ws4 2 9 "行号不连续：当前序列为 00001、00003，缺少 00002"
bdr $ws4 2 $h4.Count
wt $ws4 3 1 "退单表"; wn $ws4 3 2 3; wt $ws4 3 3 "SF3190000000051"; wt $ws4 3 4 "TK10000510"; wt $ws4 3 5 "H000000051"; wt $ws4 3 6 "行号"; wt $ws4 3 7 "00003"; wt $ws4 3 8 "E02"; wt $ws4 3 9 "行号不连续：当前序列为 00001、00003，缺少 00002"
bdr $ws4 3 $h4.Count
wt $ws4 4 1 "退单表"; wn $ws4 4 2 4; wt $ws4 4 3 "SF3190000000051"; wt $ws4 4 4 "TK10000511"; wt $ws4 4 5 "H000000051"; wt $ws4 4 6 "行号"; wt $ws4 4 7 "00002"; wt $ws4 4 8 "E02"; wt $ws4 4 9 "行号不从 00001 起：当前序列首行为 00002"
bdr $ws4 4 $h4.Count
wt $ws4 5 1 "退单表"; wn $ws4 5 2 5; wt $ws4 5 3 "SF3190000000051"; wt $ws4 5 4 "TK10000511"; wt $ws4 5 5 "H000000051"; wt $ws4 5 6 "行号"; wt $ws4 5 7 "00003"; wt $ws4 5 8 "E02"; wt $ws4 5 9 "行号不从 00001 起：当前序列首行为 00002"
bdr $ws4 5 $h4.Count
$ws4.Columns.AutoFit() | Out-Null

while ($wb.Sheets.Count -gt 5) { $wb.Sheets($wb.Sheets.Count).Delete() }
$p51 = "$outputDir\SF0051_测试数据.xlsx"
$wb.SaveAs($p51, 51); $wb.Close($false)
Write-Host "  已保存: $p51"

# ============================================================
# SF0055 — TC-55 / E12-② WMS退单号跨物流单号重复
# ============================================================
$p55 = Generate-SF0055Workbook $excel $outputDir
$p36 = Generate-SF0036Workbook $excel $outputDir

# ============================================================
# M05 缺口 DataSet（TC-13~20、TC-37）
# ============================================================
$p13  = Generate-SF0013Workbook $excel $outputDir
$p14  = Generate-SF0014Workbook $excel $outputDir
$p15  = Generate-SF0015Workbook $excel $outputDir
$p56  = Generate-SF0056Workbook $excel $outputDir
$p57  = Generate-SF0057Workbook $excel $outputDir
$p58  = Generate-SF0058Workbook $excel $outputDir
$p59  = Generate-SF0059Workbook $excel $outputDir
$p60  = Generate-SF0060Workbook $excel $outputDir
$p37  = Generate-SF0037Workbook $excel $outputDir

# ============================================================
# 待自动化 TC DataSet（TC-28/40/41）
# ============================================================
$p62  = Generate-SF0062Workbook $excel $outputDir
$p63  = Generate-SF0063Workbook $excel $outputDir
$p64  = Generate-SF0064Workbook $excel $outputDir

# ============================================================
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "`n全部完成！共生成 19 个 Excel 测试工作簿"
Write-Host "  $p32"; Write-Host "  $p46"; Write-Host "  $p47"; Write-Host "  $p48"
Write-Host "  $p51"; Write-Host "  $p55"; Write-Host "  $p36"
Write-Host "  $p13";  Write-Host "  $p14";  Write-Host "  $p15"
Write-Host "  $p56";  Write-Host "  $p57";  Write-Host "  $p58"
Write-Host "  $p59";  Write-Host "  $p60";  Write-Host "  $p37"
Write-Host "  $p62";  Write-Host "  $p63";  Write-Host "  $p64"
