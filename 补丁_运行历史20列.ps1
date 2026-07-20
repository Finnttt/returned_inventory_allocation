# 对 modRunner.bas 做精确文本替换（M15 运行历史 20 列接线：时间戳/分段耗时/错误码分布）。
# 每次替换要求恰好命中 1 处，否则中止，防止误改。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\modRunner.bas"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($path, $utf8)

function Replace-Once([string]$old, [string]$new, [string]$label) {
    # 文件是 CRLF/LF 混合：把模式按行拆分转义后用 \r?\n 连接，兼容两种行尾。
    $pattern = ($old -split "`n" | ForEach-Object { [regex]::Escape($_) }) -join '\r?\n'
    $rx = New-Object System.Text.RegularExpressions.Regex($pattern)
    $count = $rx.Matches($script:text).Count
    if ($count -ne 1) { throw "[$label] 命中 $count 处（应为 1），已中止" }
    $script:text = $rx.Replace($script:text, $new, 1)
    Write-Output "OK $label"
}

# R1：干跑入口加时间戳/计时起点
Replace-Once @'
    ' 读取配置（M02）
    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
'@ @'
    ' 运行开始时间戳与计时起点（需求 §5.6：运行时间/校验耗时/总耗时）
    Dim runStartText As String
    runStartText = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Dim t0 As Single
    t0 = Timer

    ' 读取配置（M02）
    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
'@ "R1 干跑计时起点"

# R2：干跑历史行带计时与错误码分布
Replace-Once @'
    Dim emptyEvents() As AllocationEvent
    Dim runHistoryRow As OutputRow
    runHistoryRow = BuildRunHistoryRow(stats, cfg, True)
'@ @'
    Dim emptyEvents() As AllocationEvent
    Dim runHistoryRow As OutputRow
    runHistoryRow = BuildRunHistoryRow(stats, cfg, True, runStartText, _
        RN_ElapsedSecs(t0, Timer), 0, RN_ElapsedSecs(t0, Timer), _
        RN_BuildErrorCodeDistribution(validationIssues, emptyResults))
'@ "R2 干跑历史行"

# R3：完整运行入口加时间戳/计时起点
Replace-Once @'
    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))

    On Error GoTo ClearFail
'@ @'
    ' 运行开始时间戳与计时起点（需求 §5.6：运行时间/校验耗时/分配耗时/总耗时）
    Dim runStartText As String
    runStartText = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Dim t0 As Single
    t0 = Timer

    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))

    On Error GoTo ClearFail
'@ "R3 完整运行计时起点"

# R4a：前校验结束计时点
Replace-Once @'
    validationResult = ValidatePre(orders, inventory, normalizedIssues, cfg, validationIssues)

    ' 建账本（M06）
'@ @'
    validationResult = ValidatePre(orders, inventory, normalizedIssues, cfg, validationIssues)
    Dim tValidate As Single
    tValidate = Timer

    ' 建账本（M06）
'@ "R4a 校验结束计时"

# R4b：分配循环结束计时点
Replace-Once @'
    shipmentResults = RN_RunAllAllocations(orders, ledger, cfg, validationIssues)

    On Error GoTo RunFail
'@ @'
    shipmentResults = RN_RunAllAllocations(orders, ledger, cfg, validationIssues)
    Dim tAlloc As Single
    tAlloc = Timer

    On Error GoTo RunFail
'@ "R4b 分配结束计时"

# R5：完整运行历史行带分段计时与错误码分布
Replace-Once @'
    runHistoryRow = BuildRunHistoryRow(stats, cfg, False)
'@ @'
    runHistoryRow = BuildRunHistoryRow(stats, cfg, False, runStartText, _
        RN_ElapsedSecs(t0, tValidate), RN_ElapsedSecs(tValidate, tAlloc), RN_ElapsedSecs(t0, Timer), _
        RN_BuildErrorCodeDistribution(validationIssues, shipmentResults))
'@ "R5 完整运行历史行"

[System.IO.File]::WriteAllText($path, $text, $utf8)
Write-Output "PATCHED modRunner.bas"
