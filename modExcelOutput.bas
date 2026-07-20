Option Explicit

' =============================================================================
' M14_Excel写入（modExcelOutput）
' =============================================================================
' 职责：接收 M13 生成的 OutputRow，负责输出工作表清空、写入、调试日志分表、运行历史追加。
' 说明：本模块只做“写表动作”，不做任何业务计算，便于把业务逻辑与 Excel IO 解耦。
' =============================================================================

Private Const SHEET_SUMMARY As String = "分配状态汇总表"
Private Const SHEET_DETAIL As String = "成功分配明细表"
Private Const SHEET_ANOMALY As String = "数据异常明细表"
Private Const SHEET_DEBUG As String = "调试日志"
' 注：运行历史表名不在此定义，AppendRunHistory 接收 ws 由调用方(M15)传入，保持低耦合。

' 调试日志表头（19 列，与 M13 BuildDebugLogRows / 调试日志19列规格说明.md 一致）
Private Const HEADER_DEBUG_1 As String = "物流单号"
Private Const HEADER_DEBUG_2 As String = "SKU"
Private Const HEADER_DEBUG_3 As String = "WMS退单号"
Private Const HEADER_DEBUG_4 As String = "行号"
Private Const HEADER_DEBUG_5 As String = "D"
Private Const HEADER_DEBUG_6 As String = "处理序"
Private Const HEADER_DEBUG_7 As String = "动态nextMinQty"
Private Const HEADER_DEBUG_8 As String = "候选QC数"
Private Const HEADER_DEBUG_9 As String = "被排除QC列表"
Private Const HEADER_DEBUG_10 As String = "策略"
Private Const HEADER_DEBUG_11 As String = "分配QC"
Private Const HEADER_DEBUG_12 As String = "分配前QC剩余"
Private Const HEADER_DEBUG_13 As String = "分配后QC剩余"
Private Const HEADER_DEBUG_14 As String = "批号/效期组合数"
Private Const HEADER_DEBUG_15 As String = "是否回溯重试"
Private Const HEADER_DEBUG_16 As String = "实际回溯次数"
Private Const HEADER_DEBUG_17 As String = "行状态"
Private Const HEADER_DEBUG_18 As String = "错误码"
Private Const HEADER_DEBUG_19 As String = "分配失败子类型"

' 清空输出工作表的数据区（保留表头）。
' 规则：只清空输出相关表，不清空输入表、配置表、运行历史表；会删除调试日志分表。
Public Sub ClearOutputSheets(wb As Workbook, cfg As ConfigStruct)
    If wb Is Nothing Then Err.Raise vbObjectError + 1400, "ClearOutputSheets", "工作簿对象为空，无法清空输出表。"

    EO_ClearDataKeepHeader wb.Worksheets(SHEET_SUMMARY), SHEET_SUMMARY
    EO_ClearDataKeepHeader wb.Worksheets(SHEET_DETAIL), SHEET_DETAIL
    EO_ClearDataKeepHeader wb.Worksheets(SHEET_ANOMALY), SHEET_ANOMALY
    EO_ClearDataKeepHeader wb.Worksheets(SHEET_DEBUG), SHEET_DEBUG

    EO_DeleteDebugSplitSheets wb
End Sub

' 向指定工作表写入表头和数据（覆盖数据区，保留第1行作为表头）。
Public Sub WriteSheet(ws As Worksheet, rows() As OutputRow, headers() As String)
    If ws Is Nothing Then Err.Raise vbObjectError + 1401, "WriteSheet", "目标工作表为空，无法写入。"
    EO_EnsureSheetWritable ws, ws.Name

    Dim colCount As Long
    colCount = EO_StringArrayCount(headers)
    If colCount <= 0 Then Err.Raise vbObjectError + 1402, "WriteSheet", "表头为空，无法写入工作表 [" & ws.Name & "]。"

    EO_ClearDataKeepHeader ws, ws.Name

    Dim headerMatrix() As Variant
    headerMatrix = EO_HeaderMatrix(headers)
    ws.Cells(1, 1).Resize(1, colCount).Value = headerMatrix
    EO_ApplyTextFormats ws, headers

    Dim rowCount As Long
    rowCount = EO_OutputRowCount(rows)
    If rowCount <= 0 Then Exit Sub

    Dim dataMatrix() As Variant
    dataMatrix = EO_OutputRowsToMatrix(rows, colCount)
    ws.Cells(2, 1).Resize(rowCount, colCount).Value = dataMatrix
End Sub

' 写入调试日志（按阈值自动分表：调试日志、调试日志_2、调试日志_3...）。
Public Sub WriteDebugLog(wb As Workbook, rows() As OutputRow, cfg As ConfigStruct)
    If wb Is Nothing Then Err.Raise vbObjectError + 1403, "WriteDebugLog", "工作簿对象为空，无法写入调试日志。"

    Dim headers(1 To 19) As String
    headers(1) = HEADER_DEBUG_1
    headers(2) = HEADER_DEBUG_2
    headers(3) = HEADER_DEBUG_3
    headers(4) = HEADER_DEBUG_4
    headers(5) = HEADER_DEBUG_5
    headers(6) = HEADER_DEBUG_6
    headers(7) = HEADER_DEBUG_7
    headers(8) = HEADER_DEBUG_8
    headers(9) = HEADER_DEBUG_9
    headers(10) = HEADER_DEBUG_10
    headers(11) = HEADER_DEBUG_11
    headers(12) = HEADER_DEBUG_12
    headers(13) = HEADER_DEBUG_13
    headers(14) = HEADER_DEBUG_14
    headers(15) = HEADER_DEBUG_15
    headers(16) = HEADER_DEBUG_16
    headers(17) = HEADER_DEBUG_17
    headers(18) = HEADER_DEBUG_18
    headers(19) = HEADER_DEBUG_19

    EO_DeleteDebugSplitSheets wb

    Dim total As Long
    total = EO_OutputRowCount(rows)
    Dim limitPerSheet As Long
    limitPerSheet = cfg.DetailedLogLimit
    If limitPerSheet <= 0 Then limitPerSheet = DEFAULT_DETAILED_LOG_LIMIT

    Dim mainWs As Worksheet
    Set mainWs = wb.Worksheets(SHEET_DEBUG)

    ' 调试日志关闭时：保持主表只有表头，不写入任何日志数据。
    If cfg.DebugLogLevel = DEBUG_LEVEL_OFF Then
        Dim emptyRows() As OutputRow
        WriteSheet mainWs, emptyRows, headers
        Exit Sub
    End If

    If total = 0 Then
        WriteSheet mainWs, rows, headers
        Exit Sub
    End If

    Dim chunkStart As Long
    chunkStart = 1

    Dim sheetNo As Long
    sheetNo = 1

    Do While chunkStart <= total
        Dim chunkEnd As Long
        chunkEnd = chunkStart + limitPerSheet - 1
        If chunkEnd > total Then chunkEnd = total

        Dim chunkRows() As OutputRow
        chunkRows = EO_SliceRows(rows, chunkStart, chunkEnd)

        Dim targetWs As Worksheet
        If sheetNo = 1 Then
            Set targetWs = mainWs
        Else
            Set targetWs = EO_GetOrCreateSheet(wb, SHEET_DEBUG & "_" & CStr(sheetNo))
        End If

        WriteSheet targetWs, chunkRows, headers

        sheetNo = sheetNo + 1
        chunkStart = chunkEnd + 1
    Loop
End Sub

' 向运行历史追加单行，不覆盖已有记录。
Public Sub AppendRunHistory(ws As Worksheet, row As OutputRow)
    If ws Is Nothing Then Err.Raise vbObjectError + 1404, "AppendRunHistory", "运行历史工作表为空，无法追加。"
    EO_EnsureSheetWritable ws, ws.Name

    If Not IsArray(row.Values) Then Exit Sub

    Dim colCount As Long
    colCount = UBound(row.Values) - LBound(row.Values) + 1
    If colCount <= 0 Then Exit Sub

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    Dim matrix() As Variant
    ReDim matrix(1 To 1, 1 To colCount)

    Dim i As Long
    For i = 1 To colCount
        matrix(1, i) = row.Values(i)
    Next i

    ' 第 2 列“运行时间”与第 20 列“无保质期哨兵值”按文本写入，
    ' 防止 Excel 把 2026/07/19 10:33:05、2099/01/01 转成本机日期格式。
    If colCount >= 2 Then ws.Cells(nextRow, 2).NumberFormat = "@"
    If colCount >= 20 Then ws.Cells(nextRow, 20).NumberFormat = "@"
    ws.Cells(nextRow, 1).Resize(1, colCount).Value = matrix

    ' 第 1 列“运行编号”自增：表头占第 1 行，数据第 N 行编号 = N - 1。
    ws.Cells(nextRow, 1).Value = nextRow - 1
End Sub

' -----------------------------------------------------------------------------
' 私有工具函数
' -----------------------------------------------------------------------------

Private Sub EO_ClearDataKeepHeader(ByVal ws As Worksheet, ByVal displayName As String)
    EO_EnsureSheetWritable ws, displayName
    ws.Rows("2:" & ws.Rows.Count).ClearContents
End Sub

Private Sub EO_EnsureSheetWritable(ByVal ws As Worksheet, ByVal displayName As String)
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 1410, "modExcelOutput", "工作表 [" & displayName & "] 受保护，已中止写入。"
    End If
End Sub

Private Sub EO_DeleteDebugSplitSheets(ByVal wb As Workbook)
    Dim oldDisplayAlerts As Boolean
    oldDisplayAlerts = Application.DisplayAlerts

    On Error GoTo RestoreAlerts

    Dim i As Long
    For i = wb.Worksheets.Count To 1 Step -1
        If EO_IsDebugSplitSheet(wb.Worksheets(i).Name) Then
            EO_EnsureSheetWritable wb.Worksheets(i), wb.Worksheets(i).Name
            Application.DisplayAlerts = False
            wb.Worksheets(i).Delete
        End If
    Next i

RestoreAlerts:
    Application.DisplayAlerts = oldDisplayAlerts
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function EO_IsDebugSplitSheet(ByVal sheetName As String) As Boolean
    If Len(sheetName) <= Len(SHEET_DEBUG) + 1 Then Exit Function
    If Left$(sheetName, Len(SHEET_DEBUG) + 1) <> SHEET_DEBUG & "_" Then Exit Function

    Dim suffix As String
    suffix = Mid$(sheetName, Len(SHEET_DEBUG) + 2)
    If Len(suffix) = 0 Then Exit Function

    EO_IsDebugSplitSheet = EO_IsAllDigits(suffix)
End Function

Private Function EO_IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then Exit Function
    Next i
    EO_IsAllDigits = True
End Function

Private Function EO_GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EO_GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EO_GetOrCreateSheet Is Nothing Then
        Set EO_GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EO_GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function EO_StringArrayCount(ByRef values() As String) As Long
    On Error GoTo EmptyArr
    EO_StringArrayCount = UBound(values) - LBound(values) + 1
    Exit Function
EmptyArr:
    EO_StringArrayCount = 0
End Function

Private Function EO_OutputRowCount(ByRef rows() As OutputRow) As Long
    On Error GoTo EmptyArr
    EO_OutputRowCount = UBound(rows) - LBound(rows) + 1
    Exit Function
EmptyArr:
    EO_OutputRowCount = 0
End Function

Private Function EO_HeaderMatrix(ByRef headers() As String) As Variant
    Dim count As Long
    count = EO_StringArrayCount(headers)

    Dim matrix() As Variant
    ReDim matrix(1 To 1, 1 To count)

    Dim i As Long
    For i = 1 To count
        matrix(1, i) = headers(i)
    Next i

    EO_HeaderMatrix = matrix
End Function

Private Sub EO_ApplyTextFormats(ByVal ws As Worksheet, ByRef headers() As String)
    Dim colCount As Long
    colCount = EO_StringArrayCount(headers)
    If colCount <= 0 Then Exit Sub

    Dim c As Long
    For c = 1 To colCount
        Select Case headers(c)
            ' 行号/批号可能有前导零；效期/哨兵值须保留 YYYY/MM/DD 文本；
            ' 原始值列承载各类录入原值（行号、批号、效期等），均须按文本写入，
            ' 防止 Excel 把 "00001" 强转成数值 1、"00123" 强转成 123。
            Case "行号", "效期", "无保质期哨兵值", "原始值", "批号"
                ws.Columns(c).NumberFormat = "@"
        End Select
    Next c
End Sub

Private Function EO_OutputRowsToMatrix(ByRef rows() As OutputRow, ByVal colCount As Long) As Variant
    Dim rowCount As Long
    rowCount = EO_OutputRowCount(rows)

    Dim matrix() As Variant
    ReDim matrix(1 To rowCount, 1 To colCount)

    Dim r As Long
    For r = 1 To rowCount
        Dim c As Long
        For c = 1 To colCount
            If IsArray(rows(r).Values) Then
                If c >= LBound(rows(r).Values) And c <= UBound(rows(r).Values) Then
                    matrix(r, c) = rows(r).Values(c)
                Else
                    matrix(r, c) = vbNullString
                End If
            Else
                matrix(r, c) = vbNullString
            End If
        Next c
    Next r

    EO_OutputRowsToMatrix = matrix
End Function

Private Function EO_SliceRows(ByRef source() As OutputRow, ByVal startIndex As Long, ByVal endIndex As Long) As OutputRow()
    Dim count As Long
    count = endIndex - startIndex + 1
    
    Dim result() As OutputRow
    ReDim result(1 To count)

    Dim i As Long
    For i = 1 To count
        result(i) = source(startIndex + i - 1)
    Next i
    
    EO_SliceRows = result
End Function

