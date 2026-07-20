# 对 modRunner.bas 做精确文本替换：两个运行入口的 LoadConfig 包裹友好错误处理。
# 交互模式（按钮）→ 友好 MsgBox 提示修正配置；静默模式（批量/COM）→ 原样抛错供断言匹配。
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = "D:\cursor_practice\returned_inventory_allocation\modRunner.bas"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($path, $utf8)

function Replace-Once([string]$old, [string]$new, [string]$label) {
    $pattern = ($old -split "`n" | ForEach-Object { [regex]::Escape($_) }) -join '\r?\n'
    $rx = New-Object System.Text.RegularExpressions.Regex($pattern)
    $count = $rx.Matches($script:text).Count
    if ($count -ne 1) { throw "[$label] 命中 $count 处（应为 1），已中止" }
    $script:text = $rx.Replace($script:text, $new, 1)
    Write-Output "OK $label"
}

# C1：干跑入口 LoadConfig 包裹
Replace-Once @'
    ' 读取配置（M02）
    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))

    ' 清空输出表；受保护时捕获错误并中止，避免覆盖已保护内容
    On Error GoTo ClearFail
'@ @'
    ' 读取配置（M02）；配置非法时：交互模式给友好提示，静默模式原样抛错供批量断言
    Dim cfg As ConfigStruct
    On Error GoTo ConfigFail
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
    On Error GoTo 0

    ' 清空输出表；受保护时捕获错误并中止，避免覆盖已保护内容
    On Error GoTo ClearFail
'@ "C1 干跑配置包裹"

# C2：干跑入口追加 ConfigFail 处理块（放在 ClearFail 标签之前）
Replace-Once @'
    Exit Sub

ClearFail:
    If showMessages Then
        MsgBox "清空输出表失败（" & Err.Description & "），已中止运行。" & vbNewLine & _
               "请检查输出工作表是否受保护。", vbCritical
    Else
        Err.Raise Err.Number, "RunValidationOnly", Err.Description
    End If
    Exit Sub
'@ @'
    Exit Sub

ConfigFail:
    If showMessages Then
        MsgBox "配置读取失败：" & Err.Description & vbNewLine & _
               "请修正 输入_配置 后重试。", vbCritical
    Else
        Err.Raise Err.Number, "RunValidationOnly", Err.Description
    End If
    Exit Sub

ClearFail:
    If showMessages Then
        MsgBox "清空输出表失败（" & Err.Description & "），已中止运行。" & vbNewLine & _
               "请检查输出工作表是否受保护。", vbCritical
    Else
        Err.Raise Err.Number, "RunValidationOnly", Err.Description
    End If
    Exit Sub
'@ "C2 干跑ConfigFail处理"

# C3：完整运行入口 LoadConfig 包裹
Replace-Once @'
    Dim cfg As ConfigStruct
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))

    On Error GoTo ClearFail
'@ @'
    Dim cfg As ConfigStruct
    On Error GoTo ConfigFail
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
    On Error GoTo 0

    On Error GoTo ClearFail
'@ "C3 完整运行配置包裹"

# C4：完整运行入口追加 ConfigFail 处理块
Replace-Once @'
    Exit Sub

ClearFail:
    If showMessages Then
        MsgBox "清空输出表失败（" & Err.Description & "），已中止运行。" & vbNewLine & _
               "请检查输出工作表是否受保护。", vbCritical
    Else
        Err.Raise Err.Number, "RunFullAllocation", Err.Description
    End If
    Exit Sub
'@ @'
    Exit Sub

ConfigFail:
    If showMessages Then
        MsgBox "配置读取失败：" & Err.Description & vbNewLine & _
               "请修正 输入_配置 后重试。", vbCritical
    Else
        Err.Raise Err.Number, "RunFullAllocation", Err.Description
    End If
    Exit Sub

ClearFail:
    If showMessages Then
        MsgBox "清空输出表失败（" & Err.Description & "），已中止运行。" & vbNewLine & _
               "请检查输出工作表是否受保护。", vbCritical
    Else
        Err.Raise Err.Number, "RunFullAllocation", Err.Description
    End If
    Exit Sub
'@ "C4 完整运行ConfigFail处理"

[System.IO.File]::WriteAllText($path, $text, $utf8)
Write-Output "PATCHED modRunner.bas（配置错误友好化）"
