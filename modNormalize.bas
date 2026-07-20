Option Explicit

' =============================================================================
' M04_数据标准化（modNormalize）
' =============================================================================
' 职责：
' 1. 把 M03 读取的 Raw* 原始数据转换为强类型 Normalized* 数据。
' 2. 保留字段是否合法的标记，例如 LineNoValid / QtyValid / ExpiryValid。
' 3. 对每个非法字段生成 FieldNormalizeIssue，供 M05 输出异常明细。
'
' 注意：
' 标准化阶段只记录“字段问题”，不直接生成 E01/E03/E04/E05 等错误码。
' 错误码由 M05 校验层统一生成。
' =============================================================================

' -----------------------------------------------------------------------------
' 公开函数
' -----------------------------------------------------------------------------

Public Function NormalizeReturnRows( _
    ByRef raws() As RawReturnRow, _
    ByRef cfg As ConfigStruct, _
    ByRef outIssues() As FieldNormalizeIssue) As NormalizedReturnLine()

    ' cfg 为与库存标准化保持一致的预留参数；当前退单字段标准化不读取配置值。
    Dim result() As NormalizedReturnLine
    If Not HasRawReturnRows(raws) Then
        NormalizeReturnRows = result
        Exit Function
    End If

    ReDim result(LBound(raws) To UBound(raws))

    Dim i As Long
    For i = LBound(raws) To UBound(raws)
        NormalizeReturnLine raws(i), result(i), outIssues
    Next i

    NormalizeReturnRows = result
End Function

Public Function NormalizeInventoryRows( _
    ByRef raws() As RawInventoryRow, _
    ByRef cfg As ConfigStruct, _
    ByRef outIssues() As FieldNormalizeIssue) As NormalizedInventoryLine()

    Dim result() As NormalizedInventoryLine
    If Not HasRawInventoryRows(raws) Then
        NormalizeInventoryRows = result
        Exit Function
    End If

    ReDim result(LBound(raws) To UBound(raws))

    Dim i As Long
    For i = LBound(raws) To UBound(raws)
        NormalizeInventoryLine raws(i), cfg, result(i), outIssues
    Next i

    NormalizeInventoryRows = result
End Function

' -----------------------------------------------------------------------------
' 单行标准化
' -----------------------------------------------------------------------------

Private Sub NormalizeReturnLine( _
    ByRef raw As RawReturnRow, _
    ByRef target As NormalizedReturnLine, _
    ByRef issues() As FieldNormalizeIssue)

    Dim isValid As Boolean
    Dim issueKind As String

    target.ExcelRowNum = raw.ExcelRowNum

    target.ShipmentNo = NormalizeRequiredText(raw.ShipmentNo, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_RETURN_TABLE, "物流单号", raw.ShipmentNo, issueKind, target.EmptyFields

    target.WMSOrderNo = NormalizeRequiredText(raw.WMSOrderNo, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_RETURN_TABLE, "WMS退单号", raw.WMSOrderNo, issueKind, target.EmptyFields

    target.SKU = NormalizeRequiredText(raw.SKU, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_RETURN_TABLE, "SKU", raw.SKU, issueKind, target.EmptyFields

    target.LineNo = NormalizeLineNo(raw.LineNo, isValid, issueKind)
    target.LineNoValid = isValid
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_RETURN_TABLE, "行号", raw.LineNo, issueKind, target.EmptyFields

    target.Qty = NormalizeQty(raw.Qty, isValid, issueKind)
    target.QtyValid = isValid
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_RETURN_TABLE, "数量", raw.Qty, issueKind, target.EmptyFields
End Sub

Private Sub NormalizeInventoryLine( _
    ByRef raw As RawInventoryRow, _
    ByRef cfg As ConfigStruct, _
    ByRef target As NormalizedInventoryLine, _
    ByRef issues() As FieldNormalizeIssue)

    Dim isValid As Boolean
    Dim issueKind As String

    target.ExcelRowNum = raw.ExcelRowNum

    target.ShipmentNo = NormalizeRequiredText(raw.ShipmentNo, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "物流单号", raw.ShipmentNo, issueKind, target.EmptyFields

    target.SKU = NormalizeRequiredText(raw.SKU, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "SKU", raw.SKU, issueKind, target.EmptyFields

    target.QC = NormalizeQC(raw.QC, isValid, issueKind)
    target.QCValid = isValid
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "QC情况", raw.QC, issueKind, target.EmptyFields

    target.LotNo = NormalizeLotNo(raw.LotNo, cfg, isValid, issueKind)
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "批号", raw.LotNo, issueKind, target.EmptyFields

    target.Expiry = NormalizeExpiry(raw.Expiry, raw.ExpiryCellKind, isValid, issueKind)
    target.ExpiryValid = isValid
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "效期", raw.Expiry, issueKind, target.EmptyFields

    target.Qty = NormalizeQty(raw.Qty, isValid, issueKind)
    target.QtyValid = isValid
    If Not isValid Then AddFieldIssue issues, raw.ExcelRowNum, SOURCE_INVENTORY_TABLE, "数量", raw.Qty, issueKind, target.EmptyFields
End Sub

' -----------------------------------------------------------------------------
' 字段标准化
' -----------------------------------------------------------------------------

Private Function NormalizeRequiredText(ByVal rawValue As Variant, ByRef isValid As Boolean, ByRef issueKind As String) As String
    NormalizeRequiredText = Trim$(VariantToText(rawValue))

    If NormalizeRequiredText = vbNullString Then
        isValid = False
        issueKind = ISSUE_KIND_EMPTY
    Else
        isValid = True
        issueKind = vbNullString
    End If
End Function

Private Function NormalizeLineNo(ByVal rawValue As Variant, ByRef isValid As Boolean, ByRef issueKind As String) As String
    NormalizeLineNo = Trim$(VariantToText(rawValue))

    If NormalizeLineNo = vbNullString Then
        isValid = False
        issueKind = ISSUE_KIND_EMPTY
    ElseIf Len(NormalizeLineNo) = 5 And IsAllDigits(NormalizeLineNo) Then
        isValid = True
        issueKind = vbNullString
    Else
        isValid = False
        issueKind = ISSUE_KIND_FORMAT_ERROR
    End If
End Function

Private Function NormalizeQC(ByVal rawValue As Variant, ByRef isValid As Boolean, ByRef issueKind As String) As String
    NormalizeQC = UCase$(Trim$(VariantToText(rawValue)))

    If NormalizeQC = vbNullString Then
        isValid = False
        issueKind = ISSUE_KIND_EMPTY
    ElseIf NormalizeQC = QC_ZP Or NormalizeQC = QC_QC Or NormalizeQC = QC_NG Then
        isValid = True
        issueKind = vbNullString
    Else
        isValid = False
        issueKind = ISSUE_KIND_FORMAT_ERROR
    End If
End Function

Private Function NormalizeLotNo( _
    ByVal rawValue As Variant, _
    ByRef cfg As ConfigStruct, _
    ByRef isValid As Boolean, _
    ByRef issueKind As String) As String

    NormalizeLotNo = Trim$(VariantToText(rawValue))

    If NormalizeLotNo = vbNullString Then
        isValid = False
        issueKind = ISSUE_KIND_EMPTY
        Exit Function
    End If

    If Not cfg.LotCaseSensitive Then
        NormalizeLotNo = UCase$(NormalizeLotNo)
    End If

    isValid = True
    issueKind = vbNullString
End Function

Private Function NormalizeQty(ByVal rawValue As Variant, ByRef isValid As Boolean, ByRef issueKind As String) As Long
    Dim textValue As String
    textValue = Trim$(VariantToText(rawValue))

    If textValue = vbNullString Then
        isValid = False
        issueKind = ISSUE_KIND_EMPTY
        NormalizeQty = 0
        Exit Function
    End If

    If Not IsNumeric(textValue) Then
        isValid = False
        issueKind = ISSUE_KIND_FORMAT_ERROR
        NormalizeQty = 0
        Exit Function
    End If

    Dim numericValue As Double
    numericValue = CDbl(textValue)

    If numericValue <= 0 Or numericValue <> Fix(numericValue) Then
        isValid = False
        issueKind = ISSUE_KIND_RANGE_ERROR
        NormalizeQty = 0
        Exit Function
    End If

    isValid = True
    issueKind = vbNullString
    NormalizeQty = CLng(numericValue)
End Function

Private Function NormalizeExpiry( _
    ByVal rawValue As Variant, _
    ByVal cellKind As String, _
    ByRef isValid As Boolean, _
    ByRef issueKind As String) As String

    Select Case cellKind
        Case CELL_KIND_EXCEL_DATE
            If IsDate(rawValue) Then
                NormalizeExpiry = Format$(CDate(rawValue), "yyyy/mm/dd")
                isValid = True
                issueKind = vbNullString
            Else
                NormalizeExpiry = vbNullString
                isValid = False
                issueKind = ISSUE_KIND_FORMAT_ERROR
            End If

        Case CELL_KIND_TEXT
            NormalizeExpiry = NormalizeTextDate(Trim$(VariantToText(rawValue)), isValid)
            If isValid Then
                issueKind = vbNullString
            Else
                issueKind = ISSUE_KIND_FORMAT_ERROR
            End If

        Case CELL_KIND_BLANK
            NormalizeExpiry = vbNullString
            isValid = False
            issueKind = ISSUE_KIND_EMPTY

        Case Else
            NormalizeExpiry = vbNullString
            isValid = False
            issueKind = ISSUE_KIND_FORMAT_ERROR
    End Select
End Function

' -----------------------------------------------------------------------------
' 日期文本校验
' -----------------------------------------------------------------------------

Private Function NormalizeTextDate(ByVal textValue As String, ByRef isValid As Boolean) As String
    If textValue = vbNullString Then
        isValid = False
        NormalizeTextDate = vbNullString
        Exit Function
    End If

    Dim separator As String
    If InStr(1, textValue, "/", vbBinaryCompare) > 0 Then
        separator = "/"
    ElseIf InStr(1, textValue, "-", vbBinaryCompare) > 0 Then
        separator = "-"
    Else
        isValid = False
        NormalizeTextDate = vbNullString
        Exit Function
    End If

    Dim parts() As String
    parts = Split(textValue, separator)
    If UBound(parts) - LBound(parts) + 1 <> 3 Then
        isValid = False
        NormalizeTextDate = vbNullString
        Exit Function
    End If

    If Len(parts(0)) <> 4 Or Len(parts(1)) <> 2 Or Len(parts(2)) <> 2 Then
        isValid = False
        NormalizeTextDate = vbNullString
        Exit Function
    End If

    If Not IsAllDigits(parts(0)) Or Not IsAllDigits(parts(1)) Or Not IsAllDigits(parts(2)) Then
        isValid = False
        NormalizeTextDate = vbNullString
        Exit Function
    End If

    Dim yearValue As Long
    Dim monthValue As Long
    Dim dayValue As Long

    yearValue = CLng(parts(0))
    monthValue = CLng(parts(1))
    dayValue = CLng(parts(2))

    If yearValue < 1900 Or yearValue > 2999 Then
        isValid = False
    ElseIf monthValue < 1 Or monthValue > 12 Then
        isValid = False
    ElseIf dayValue < 1 Or dayValue > DaysInMonth(yearValue, monthValue) Then
        isValid = False
    Else
        isValid = True
    End If

    If isValid Then
        NormalizeTextDate = Format$(DateSerial(yearValue, monthValue, dayValue), "yyyy/mm/dd")
    Else
        NormalizeTextDate = vbNullString
    End If
End Function

Private Function DaysInMonth(ByVal yearValue As Long, ByVal monthValue As Long) As Long
    Select Case monthValue
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = 31
        Case 4, 6, 9, 11
            DaysInMonth = 30
        Case 2
            If IsLeapYear(yearValue) Then
                DaysInMonth = 29
            Else
                DaysInMonth = 28
            End If
    End Select
End Function

Private Function IsLeapYear(ByVal yearValue As Long) As Boolean
    IsLeapYear = (yearValue Mod 4 = 0 And yearValue Mod 100 <> 0) Or (yearValue Mod 400 = 0)
End Function

' -----------------------------------------------------------------------------
' Issue 记录工具
' -----------------------------------------------------------------------------

Private Sub AddFieldIssue( _
    ByRef issues() As FieldNormalizeIssue, _
    ByVal excelRowNum As Long, _
    ByVal sourceTable As String, _
    ByVal fieldName As String, _
    ByVal rawValue As Variant, _
    ByVal issueKind As String, _
    ByRef emptyFields As String)

    If issueKind = ISSUE_KIND_EMPTY Then
        emptyFields = AppendFieldName(emptyFields, fieldName)
    End If

    Dim newIndex As Long
    If HasFieldIssues(issues) Then
        newIndex = UBound(issues) + 1
        ReDim Preserve issues(LBound(issues) To newIndex)
    Else
        newIndex = 1
        ReDim issues(1 To 1)
    End If

    With issues(newIndex)
        .ExcelRowNum = excelRowNum
        .SourceTable = sourceTable
        .FieldName = fieldName
        .RawValue = VariantToText(rawValue)
        .IssueKind = issueKind
    End With
End Sub

Private Function AppendFieldName(ByVal currentValue As String, ByVal fieldName As String) As String
    If currentValue = vbNullString Then
        AppendFieldName = fieldName
    Else
        AppendFieldName = currentValue & "," & fieldName
    End If
End Function

' -----------------------------------------------------------------------------
' 通用工具
' -----------------------------------------------------------------------------

Private Function VariantToText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        VariantToText = vbNullString
    Else
        VariantToText = CStr(rawValue)
    End If
End Function

Private Function IsAllDigits(ByVal textValue As String) As Boolean
    If textValue = vbNullString Then
        IsAllDigits = False
        Exit Function
    End If

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i

    IsAllDigits = True
End Function

Private Function HasRawReturnRows(ByRef raws() As RawReturnRow) As Boolean
    On Error GoTo NotAllocated
    HasRawReturnRows = (UBound(raws) >= LBound(raws))
    Exit Function

NotAllocated:
    HasRawReturnRows = False
End Function

Private Function HasRawInventoryRows(ByRef raws() As RawInventoryRow) As Boolean
    On Error GoTo NotAllocated
    HasRawInventoryRows = (UBound(raws) >= LBound(raws))
    Exit Function

NotAllocated:
    HasRawInventoryRows = False
End Function

Private Function HasFieldIssues(ByRef issues() As FieldNormalizeIssue) As Boolean
    On Error GoTo NotAllocated
    HasFieldIssues = (UBound(issues) >= LBound(issues))
    Exit Function

NotAllocated:
    HasFieldIssues = False
End Function
