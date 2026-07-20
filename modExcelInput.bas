Option Explicit

' =============================================================================
' M03_数据加载（modExcelInput）
' =============================================================================
' 职责：
' 1. 从 Excel 工作表读取原始数据到 RawReturnRow / RawInventoryRow。
' 2. 保留单元格原始值，不做 Trim、不补零、不转业务类型。
' 3. 对效期列额外记录 ExpiryCellKind，供 M04 选择正确的标准化路径。
'
' 当前已锁定输入表结构：
' 输入_退单表：物流单号 | WMS退单号 | SKU | 行号 | 数量
' 输入_质检库存表：物流单号 | SKU | QC情况 | 批号 | 效期 | 数量 | 备注
' =============================================================================

Private Const ERR_BASE As Long = vbObjectError + 3000

' 退单表列号
Private Const COL_RETURN_SHIPMENT_NO As Long = 1
Private Const COL_RETURN_WMS_ORDER_NO As Long = 2
Private Const COL_RETURN_SKU As Long = 3
Private Const COL_RETURN_LINE_NO As Long = 4
Private Const COL_RETURN_QTY As Long = 5

' 质检库存表列号
Private Const COL_INV_SHIPMENT_NO As Long = 1
Private Const COL_INV_SKU As Long = 2
Private Const COL_INV_QC As Long = 3
Private Const COL_INV_LOT_NO As Long = 4
Private Const COL_INV_EXPIRY As Long = 5
Private Const COL_INV_QTY As Long = 6

' -----------------------------------------------------------------------------
' 公开函数
' -----------------------------------------------------------------------------

Public Function ReadReturnOrders(ByVal ws As Worksheet) As RawReturnRow()
    ValidateReturnHeader ws

    Dim lastRow As Long
    lastRow = GetLastUsedRow(ws)

    Dim result() As RawReturnRow
    If lastRow < 2 Then
        ReadReturnOrders = result
        Exit Function
    End If

    ReDim result(1 To lastRow - 1)

    ' E12-②：维护 WMS退单号 → 物流单号 映射，发现跨物流单号重复立即中止（§4.1 第0层 / §6.3.1.1 步骤 1b）
    Dim wmsToShipment As Object
    Set wmsToShipment = CreateObject("Scripting.Dictionary")

    Dim rowIndex As Long
    Dim outIndex As Long
    For rowIndex = 2 To lastRow
        outIndex = rowIndex - 1

        Dim wmsKey As String
        Dim shipKey As String
        wmsKey = Trim$(CStr(ws.Cells(rowIndex, COL_RETURN_WMS_ORDER_NO).Value))
        shipKey = Trim$(CStr(ws.Cells(rowIndex, COL_RETURN_SHIPMENT_NO).Value))

        If wmsKey <> vbNullString Then
            If wmsToShipment.Exists(wmsKey) Then
                If wmsToShipment(wmsKey) <> shipKey Then
                    RaiseInputError "E12：WMS退单号 [" & wmsKey & "] 已出现在物流单号 [" & wmsToShipment(wmsKey) & "]，" & _
                                    "当前行 " & CStr(rowIndex) & " 又出现在物流单号 [" & shipKey & "]。"
                End If
            Else
                wmsToShipment.Add wmsKey, shipKey
            End If
        End If

        With result(outIndex)
            .ExcelRowNum = rowIndex
            .ShipmentNo = ws.Cells(rowIndex, COL_RETURN_SHIPMENT_NO).Value
            .WMSOrderNo = ws.Cells(rowIndex, COL_RETURN_WMS_ORDER_NO).Value
            .SKU = ws.Cells(rowIndex, COL_RETURN_SKU).Value
            .LineNo = ws.Cells(rowIndex, COL_RETURN_LINE_NO).Value
            .Qty = ws.Cells(rowIndex, COL_RETURN_QTY).Value
        End With
    Next rowIndex

    ReadReturnOrders = result
End Function

Public Function ReadQCInventory(ByVal ws As Worksheet) As RawInventoryRow()
    ValidateInventoryHeader ws

    Dim lastRow As Long
    lastRow = GetLastUsedRow(ws)

    Dim result() As RawInventoryRow
    If lastRow < 2 Then
        ReadQCInventory = result
        Exit Function
    End If

    ReDim result(1 To lastRow - 1)

    Dim rowIndex As Long
    Dim outIndex As Long
    For rowIndex = 2 To lastRow
        outIndex = rowIndex - 1

        With result(outIndex)
            .ExcelRowNum = rowIndex
            .ShipmentNo = ws.Cells(rowIndex, COL_INV_SHIPMENT_NO).Value
            .SKU = ws.Cells(rowIndex, COL_INV_SKU).Value
            .QC = ws.Cells(rowIndex, COL_INV_QC).Value
            .LotNo = ws.Cells(rowIndex, COL_INV_LOT_NO).Value
            .Expiry = ws.Cells(rowIndex, COL_INV_EXPIRY).Value
            .ExpiryCellKind = GetExpiryCellKind(ws.Cells(rowIndex, COL_INV_EXPIRY))
            .Qty = ws.Cells(rowIndex, COL_INV_QTY).Value
        End With
    Next rowIndex

    ReadQCInventory = result
End Function

' -----------------------------------------------------------------------------
' 表头校验
' -----------------------------------------------------------------------------

Private Sub ValidateReturnHeader(ByVal ws As Worksheet)
    ValidateWorksheet ws

    RequireHeader ws, COL_RETURN_SHIPMENT_NO, "物流单号"
    RequireHeader ws, COL_RETURN_WMS_ORDER_NO, "WMS退单号"
    RequireHeader ws, COL_RETURN_SKU, "SKU"
    RequireHeader ws, COL_RETURN_LINE_NO, "行号"
    RequireHeader ws, COL_RETURN_QTY, "数量"
End Sub

Private Sub ValidateInventoryHeader(ByVal ws As Worksheet)
    ValidateWorksheet ws

    RequireHeader ws, COL_INV_SHIPMENT_NO, "物流单号"
    RequireHeader ws, COL_INV_SKU, "SKU"
    RequireHeader ws, COL_INV_QC, "QC情况"
    RequireHeader ws, COL_INV_LOT_NO, "批号"
    RequireHeader ws, COL_INV_EXPIRY, "效期"
    RequireHeader ws, COL_INV_QTY, "数量"
End Sub

Private Sub RequireHeader(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal expectedName As String)
    Dim actualName As String
    actualName = Trim$(CStr(ws.Cells(1, colIndex).Value))

    If actualName <> expectedName Then
        RaiseInputError "表头校验失败：工作表 [" & ws.Name & "] 第 " & CStr(colIndex) & _
                        " 列应为 [" & expectedName & "]，实际为 [" & actualName & "]。"
    End If
End Sub

' -----------------------------------------------------------------------------
' 效期单元格类型
' -----------------------------------------------------------------------------

Private Function GetExpiryCellKind(ByVal cell As Range) As String
    ' 必须使用 cell.Value，而不是 Value2。
    ' Value2 会把 Excel 日期返回为数字，无法准确得到 vbDate。
    Select Case VarType(cell.Value)
        Case vbDate
            GetExpiryCellKind = CELL_KIND_EXCEL_DATE
        Case vbString
            GetExpiryCellKind = CELL_KIND_TEXT
        Case vbEmpty
            GetExpiryCellKind = CELL_KIND_BLANK
        Case Else
            GetExpiryCellKind = CELL_KIND_OTHER
    End Select
End Function

' -----------------------------------------------------------------------------
' 通用工具
' -----------------------------------------------------------------------------

Private Sub ValidateWorksheet(ByVal ws As Worksheet)
    If ws Is Nothing Then
        RaiseInputError "数据加载失败：工作表对象为空。"
    End If
End Sub

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False)

    If lastCell Is Nothing Then
        GetLastUsedRow = 1
    Else
        GetLastUsedRow = lastCell.Row
    End If
End Function

Private Sub RaiseInputError(ByVal message As String)
    Err.Raise ERR_BASE + 1, "modExcelInput", message
End Sub
