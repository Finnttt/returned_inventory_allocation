Option Explicit

' =============================================================================
' M05_分配前校验（modValidate）
' =============================================================================
' 职责：
' 1. 接收 M04 标准化结果，执行 §4.1 四层校验（E01~E08、E11）。
' 2. 收集所有命中错误码，不因前层错误跳过后层（除 E08/E11 的跳过规则）。
' 3. 产出 ValidationResult + ValidationIssue[]，并可通过 BuildAnomalyRows 生成异常明细。
'
' 给新手的解释：
' M04 只判断“某个字段本身是否合法”；M05 把这些字段问题升级成业务错误码，
' 并额外检查跨行、跨表、跨 SKU 的规则，例如行号连续性和数量是否一致。
' =============================================================================

' -----------------------------------------------------------------------------
' 公开函数
' -----------------------------------------------------------------------------

Public Function ValidatePre( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef fieldIssues() As FieldNormalizeIssue, _
    ByRef cfg As ConfigStruct, _
    ByRef outIssues() As ValidationIssue) As ValidationResult

    ' cfg 为统一校验接口的预留参数；当前 E01～E11 规则不读取配置值。
    Dim result As ValidationResult
    Dim issues() As ValidationIssue
    Dim failedShipments As Object
    Dim shipmentHasE04 As Object
    Dim shipmentHasE08 As Object

    Set failedShipments = CreateObject("Scripting.Dictionary")
    Set shipmentHasE04 = CreateObject("Scripting.Dictionary")
    Set shipmentHasE08 = CreateObject("Scripting.Dictionary")

    ' 第1层：E01~E05（字段级）+ E02（行号重复/不连续）
    ApplyLayer1FieldIssues orders, inventory, fieldIssues, issues, failedShipments, shipmentHasE04
    ApplyLayer1LineNoContinuity orders, issues, failedShipments

    ' 第2层：E06、E07（物流单号集合比对）
    ApplyLayer2ShipmentConsistency orders, inventory, issues, failedShipments

    ' 第3层：E08（数量一致性）；已命中 E04 或任一输入表缺少物流单号时跳过
    ApplyLayer3QtyConsistency orders, inventory, issues, failedShipments, shipmentHasE04, shipmentHasE08

    ' 第4层：E11（碎片库存）；已命中 E04 或 E08 的物流单号跳过
    ApplyLayer4FragmentInventory orders, inventory, issues, failedShipments, shipmentHasE04, shipmentHasE08

    outIssues = issues
    result.HasFailures = (failedShipments.Count > 0)
    result.FailedShipmentCount = failedShipments.Count
    ValidatePre = result
End Function

Public Function BuildAnomalyRows( _
    ByRef validationIssues() As ValidationIssue) As AnomalyRow()

    Dim result() As AnomalyRow
    Dim i As Long
    Dim outIndex As Long

    If Not HasValidationIssues(validationIssues) Then
        BuildAnomalyRows = result
        Exit Function
    End If

    ReDim result(1 To CountValidationIssues(validationIssues))

    For i = LBound(validationIssues) To UBound(validationIssues)
        If IsAnomalyDetailError(validationIssues(i).ErrorCode) Then
            outIndex = outIndex + 1
            With result(outIndex)
                .SourceTable = validationIssues(i).SourceTable
                .ExcelRowNum = validationIssues(i).ExcelRowNum
                .ShipmentNo = validationIssues(i).ShipmentNo
                .WMSOrderNo = validationIssues(i).WMSOrderNo
                .SKU = validationIssues(i).SKU
                .FieldName = validationIssues(i).FieldName
                .RawValue = validationIssues(i).RawValue
                .ErrorCode = validationIssues(i).ErrorCode
                .Reason = validationIssues(i).Reason
            End With
        End If
    Next i

    If outIndex = 0 Then
        Erase result
    ElseIf outIndex < UBound(result) Then
        ReDim Preserve result(1 To outIndex)
    End If

    BuildAnomalyRows = result
End Function

' -----------------------------------------------------------------------------
' 第1层：字段合法性 + E02
' -----------------------------------------------------------------------------

Private Sub ApplyLayer1FieldIssues( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef fieldIssues() As FieldNormalizeIssue, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object, _
    ByRef shipmentHasE04 As Object)

    Dim i As Long

    If HasFieldIssues(fieldIssues) Then
        For i = LBound(fieldIssues) To UBound(fieldIssues)
            AppendFieldValidationIssue fieldIssues(i), orders, inventory, issues, failedShipments, shipmentHasE04
        Next i
    End If
End Sub

Private Sub AppendFieldValidationIssue( _
    ByRef fieldIssue As FieldNormalizeIssue, _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object, _
    ByRef shipmentHasE04 As Object)

    Dim issue As ValidationIssue
    Dim errorCode As String
    Dim reason As String

    errorCode = MapFieldIssueToErrorCode(fieldIssue)
    reason = BuildFieldIssueReason(fieldIssue, errorCode)

    issue.SourceTable = fieldIssue.SourceTable
    issue.ExcelRowNum = fieldIssue.ExcelRowNum
    issue.FieldName = fieldIssue.FieldName
    issue.RawValue = fieldIssue.RawValue
    issue.ErrorCode = errorCode
    issue.Reason = reason

    FillIssueContext issue, orders, inventory
    AppendValidationIssue issues, issue, failedShipments

    If errorCode = ERR_E04 And issue.ShipmentNo <> vbNullString Then
        shipmentHasE04(issue.ShipmentNo) = True
    End If
End Sub

Private Function MapFieldIssueToErrorCode(ByRef fieldIssue As FieldNormalizeIssue) As String
    If fieldIssue.IssueKind = ISSUE_KIND_EMPTY Then
        MapFieldIssueToErrorCode = ERR_E01
        Exit Function
    End If

    Select Case fieldIssue.FieldName
        Case "QC情况"
            MapFieldIssueToErrorCode = ERR_E03
        Case "效期"
            MapFieldIssueToErrorCode = ERR_E05
        Case "数量"
            MapFieldIssueToErrorCode = ERR_E04
        Case Else
            MapFieldIssueToErrorCode = ERR_E01
    End Select
End Function

Private Function BuildFieldIssueReason(ByRef fieldIssue As FieldNormalizeIssue, ByVal errorCode As String) As String
    Select Case errorCode
        Case ERR_E01
            If fieldIssue.IssueKind = ISSUE_KIND_EMPTY Then
                BuildFieldIssueReason = "字段为空"
            ElseIf fieldIssue.FieldName = "行号" Then
                BuildFieldIssueReason = "行号格式不符（须为五位前导零文本）"
            Else
                BuildFieldIssueReason = "关键字段为空或格式异常"
            End If
        Case ERR_E03
            BuildFieldIssueReason = "QC情况非法（仅允许ZP/QC/NG）"
        Case ERR_E04
            BuildFieldIssueReason = "数量非法（非正整数）"
        Case ERR_E05
            BuildFieldIssueReason = "效期无法解析为合法日期"
        Case Else
            BuildFieldIssueReason = "关键字段为空或格式异常"
    End Select
End Function

Private Sub FillIssueContext( _
    ByRef issue As ValidationIssue, _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine)

    Dim i As Long

    If issue.SourceTable = SOURCE_RETURN_TABLE And HasNormalizedReturnRows(orders) Then
        For i = LBound(orders) To UBound(orders)
            If orders(i).ExcelRowNum = issue.ExcelRowNum Then
                issue.ShipmentNo = orders(i).ShipmentNo
                issue.WMSOrderNo = orders(i).WMSOrderNo
                issue.SKU = orders(i).SKU
                ' 这里只结束查找，不能退出整个过程；后面还要把空字段统一为 [N/A]。
                Exit For
            End If
        Next i
    ElseIf issue.SourceTable = SOURCE_INVENTORY_TABLE And HasNormalizedInventoryRows(inventory) Then
        For i = LBound(inventory) To UBound(inventory)
            If inventory(i).ExcelRowNum = issue.ExcelRowNum Then
                issue.ShipmentNo = inventory(i).ShipmentNo
                issue.WMSOrderNo = NA_PLACEHOLDER
                issue.SKU = inventory(i).SKU
                Exit For
            End If
        Next i
    End If

    If issue.ShipmentNo = vbNullString Then issue.ShipmentNo = NA_PLACEHOLDER
    If issue.WMSOrderNo = vbNullString Then issue.WMSOrderNo = NA_PLACEHOLDER
    If issue.SKU = vbNullString Then issue.SKU = NA_PLACEHOLDER
End Sub

Private Sub ApplyLayer1LineNoContinuity( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object)

    Dim wmsGroups As Object
    Dim wmsKey As Variant
    Dim rowIndexes() As Long
    Dim lineNums() As Long
    Dim reason As String

    If Not HasNormalizedReturnRows(orders) Then Exit Sub

    Set wmsGroups = GroupReturnRowsByWms(orders)

    For Each wmsKey In wmsGroups.Keys
        ExtractWmsLineNumbers orders, wmsGroups(wmsKey), rowIndexes, lineNums

        If UBound(lineNums) >= LBound(lineNums) Then
            If DetectLineNoDuplicate(lineNums) Then
                reason = "退单表行号重复"
                AppendE02IssuesForWms orders, rowIndexes, reason, issues, failedShipments
            ElseIf DetectLineNoDiscontinuity(lineNums, reason) Then
                AppendE02IssuesForWms orders, rowIndexes, reason, issues, failedShipments
            End If
        End If
    Next wmsKey
End Sub

Private Function GroupReturnRowsByWms(ByRef orders() As NormalizedReturnLine) As Object
    Dim groups As Object
    Dim i As Long
    Dim wmsKey As String
    Dim bucket As String

    Set groups = CreateObject("Scripting.Dictionary")

    For i = LBound(orders) To UBound(orders)
        If orders(i).LineNoValid And orders(i).WMSOrderNo <> vbNullString Then
            wmsKey = orders(i).WMSOrderNo

            If groups.Exists(wmsKey) Then
                bucket = groups(wmsKey) & "," & CStr(i)
            Else
                bucket = CStr(i)
            End If

            groups(wmsKey) = bucket
        End If
    Next i

    Set GroupReturnRowsByWms = groups
End Function

Private Sub ExtractWmsLineNumbers( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal indexList As String, _
    ByRef rowIndexes() As Long, _
    ByRef lineNums() As Long)

    Dim parts() As String
    Dim i As Long
    Dim idx As Long
    Dim count As Long
    Dim n As Long

    parts = Split(indexList, ",")
    n = UBound(parts) - LBound(parts) + 1
    ReDim rowIndexes(1 To n)
    ReDim lineNums(1 To n)

    For i = LBound(parts) To UBound(parts)
        idx = CLng(parts(i))
        count = count + 1
        rowIndexes(count) = idx
        lineNums(count) = CLng(orders(idx).LineNo)
    Next i

    SortLongArray lineNums, rowIndexes, count
End Sub

Private Function DetectLineNoDuplicate(ByRef lineNums() As Long) As Boolean
    Dim i As Long

    For i = LBound(lineNums) To UBound(lineNums) - 1
        If lineNums(i) = lineNums(i + 1) Then
            DetectLineNoDuplicate = True
            Exit Function
        End If
    Next i
End Function

Private Function DetectLineNoDiscontinuity(ByRef lineNums() As Long, ByRef reason As String) As Boolean
    Dim i As Long
    Dim minVal As Long
    Dim maxVal As Long
    Dim expectedCount As Long

    minVal = lineNums(LBound(lineNums))
    maxVal = lineNums(UBound(lineNums))
    expectedCount = maxVal - minVal + 1

    If minVal <> 1 Then
        reason = "行号不从 00001 起：当前序列首行为 " & FormatLineNo(minVal)
        DetectLineNoDiscontinuity = True
        Exit Function
    End If

    If expectedCount <> (UBound(lineNums) - LBound(lineNums) + 1) Then
        reason = BuildGapReason(lineNums)
        DetectLineNoDiscontinuity = True
        Exit Function
    End If

    For i = LBound(lineNums) To UBound(lineNums) - 1
        If lineNums(i + 1) <> lineNums(i) + 1 Then
            reason = BuildGapReason(lineNums)
            DetectLineNoDiscontinuity = True
            Exit Function
        End If
    Next i
End Function

Private Function BuildGapReason(ByRef lineNums() As Long) As String
    Dim i As Long
    Dim textList As String

    For i = LBound(lineNums) To UBound(lineNums)
        If textList <> vbNullString Then textList = textList & "、"
        textList = textList & FormatLineNo(lineNums(i))
    Next i

    BuildGapReason = "行号不连续：当前序列为 " & textList
End Function

Private Function FormatLineNo(ByVal lineNo As Long) As String
    ' 仅用于 E02 错误原因的序列展示，不会回写输入数据，也不代表系统自动补零。
    FormatLineNo = Right$("00000" & CStr(lineNo), 5)
End Function

Private Sub AppendE02IssuesForWms( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef rowIndexes() As Long, _
    ByVal reason As String, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object)

    Dim i As Long
    Dim idx As Long
    Dim issue As ValidationIssue

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        idx = rowIndexes(i)

        issue.ShipmentNo = orders(idx).ShipmentNo
        issue.WMSOrderNo = orders(idx).WMSOrderNo
        issue.SKU = orders(idx).SKU
        issue.ErrorCode = ERR_E02
        issue.SourceTable = SOURCE_RETURN_TABLE
        issue.ExcelRowNum = orders(idx).ExcelRowNum
        issue.FieldName = "行号"
        issue.RawValue = orders(idx).LineNo
        issue.Reason = reason

        AppendValidationIssue issues, issue, failedShipments
    Next i
End Sub

' -----------------------------------------------------------------------------
' 第2层：E06、E07
' -----------------------------------------------------------------------------

Private Sub ApplyLayer2ShipmentConsistency( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object)

    Dim orderShipments As Object
    Dim inventoryShipments As Object
    Dim shipmentKey As Variant
    Dim issue As ValidationIssue
    Dim i As Long

    Set orderShipments = CollectShipmentNosFromOrders(orders)
    Set inventoryShipments = CollectShipmentNosFromInventory(inventory)

    For Each shipmentKey In orderShipments.Keys
        If Not inventoryShipments.Exists(shipmentKey) Then
            issue.ShipmentNo = CStr(shipmentKey)
            issue.WMSOrderNo = NA_PLACEHOLDER
            issue.SKU = NA_PLACEHOLDER
            issue.ErrorCode = ERR_E06
            issue.SourceTable = SOURCE_RETURN_TABLE
            issue.ExcelRowNum = 0
            issue.FieldName = "物流单号"
            issue.RawValue = CStr(shipmentKey)
            issue.Reason = "物流单号仅存在于退单表"

            AppendValidationIssue issues, issue, failedShipments
        End If
    Next shipmentKey

    For Each shipmentKey In inventoryShipments.Keys
        If Not orderShipments.Exists(shipmentKey) Then
            If HasNormalizedInventoryRows(inventory) Then
                For i = LBound(inventory) To UBound(inventory)
                    If inventory(i).ShipmentNo = CStr(shipmentKey) Then
                        issue.ShipmentNo = inventory(i).ShipmentNo
                        issue.WMSOrderNo = NA_PLACEHOLDER
                        issue.SKU = inventory(i).SKU
                        issue.ErrorCode = ERR_E07
                        issue.SourceTable = SOURCE_INVENTORY_TABLE
                        issue.ExcelRowNum = inventory(i).ExcelRowNum
                        issue.FieldName = "物流单号"
                        issue.RawValue = inventory(i).ShipmentNo
                        issue.Reason = "物流单号仅存在于质检库存表"

                        AppendValidationIssue issues, issue, failedShipments
                    End If
                Next i
            End If
        End If
    Next shipmentKey
End Sub

Private Function CollectShipmentNosFromOrders(ByRef orders() As NormalizedReturnLine) As Object
    Dim dict As Object
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If HasNormalizedReturnRows(orders) Then
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo <> vbNullString Then
                dict(orders(i).ShipmentNo) = True
            End If
        Next i
    End If

    Set CollectShipmentNosFromOrders = dict
End Function

Private Function CollectShipmentNosFromInventory(ByRef inventory() As NormalizedInventoryLine) As Object
    Dim dict As Object
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If HasNormalizedInventoryRows(inventory) Then
        For i = LBound(inventory) To UBound(inventory)
            If inventory(i).ShipmentNo <> vbNullString Then
                dict(inventory(i).ShipmentNo) = True
            End If
        Next i
    End If

    Set CollectShipmentNosFromInventory = dict
End Function

' -----------------------------------------------------------------------------
' 第3层：E08
' -----------------------------------------------------------------------------

Private Sub ApplyLayer3QtyConsistency( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object, _
    ByRef shipmentHasE04 As Object, _
    ByRef shipmentHasE08 As Object)

    Dim groupKeys As Object
    Dim groupKey As Variant
    Dim parts() As String
    Dim shipmentNo As String
    Dim sku As String
    Dim orderQty As Long
    Dim inventoryQty As Long
    Dim issue As ValidationIssue

    Set groupKeys = CollectShipmentSkuGroups(orders, inventory)

    ' E08 只比较“两张表都存在”的物流单号。
    ' 仅存在单侧时应由 E06/E07 精确说明，不再叠加没有业务价值的 E08。
    Dim orderShipments As Object
    Dim inventoryShipments As Object
    Set orderShipments = CollectShipmentNosFromOrders(orders)
    Set inventoryShipments = CollectShipmentNosFromInventory(inventory)

    For Each groupKey In groupKeys.Keys
        parts = Split(CStr(groupKey), vbTab)
        shipmentNo = parts(0)
        sku = parts(1)

        If shipmentHasE04.Exists(shipmentNo) Then GoTo NextGroup
        If Not orderShipments.Exists(shipmentNo) Then GoTo NextGroup
        If Not inventoryShipments.Exists(shipmentNo) Then GoTo NextGroup

        orderQty = SumOrderQty(orders, shipmentNo, sku)
        inventoryQty = SumInventoryQty(inventory, shipmentNo, sku)

        If orderQty <> inventoryQty Then
            issue.ShipmentNo = shipmentNo
            issue.WMSOrderNo = NA_PLACEHOLDER
            issue.SKU = sku
            issue.ErrorCode = ERR_E08
            issue.SourceTable = NA_PLACEHOLDER
            issue.ExcelRowNum = 0
            issue.FieldName = "数量"
            issue.RawValue = CStr(orderQty) & " vs " & CStr(inventoryQty)
            issue.Reason = "同物流单号+SKU数量不一致"

            AppendValidationIssue issues, issue, failedShipments
            shipmentHasE08(shipmentNo) = True
        End If
NextGroup:
    Next groupKey
End Sub

Private Function CollectShipmentSkuGroups( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine) As Object

    Dim dict As Object
    Dim i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    If HasNormalizedReturnRows(orders) Then
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo <> vbNullString And orders(i).SKU <> vbNullString Then
                key = orders(i).ShipmentNo & vbTab & orders(i).SKU
                dict(key) = True
            End If
        Next i
    End If

    If HasNormalizedInventoryRows(inventory) Then
        For i = LBound(inventory) To UBound(inventory)
            If inventory(i).ShipmentNo <> vbNullString And inventory(i).SKU <> vbNullString Then
                key = inventory(i).ShipmentNo & vbTab & inventory(i).SKU
                dict(key) = True
            End If
        Next i
    End If

    Set CollectShipmentSkuGroups = dict
End Function

Private Function SumOrderQty( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipmentNo As String, _
    ByVal sku As String) As Long

    Dim i As Long
    Dim total As Long

    If Not HasNormalizedReturnRows(orders) Then Exit Function

    For i = LBound(orders) To UBound(orders)
        If orders(i).ShipmentNo = shipmentNo And orders(i).SKU = sku And orders(i).QtyValid Then
            total = total + orders(i).Qty
        End If
    Next i

    SumOrderQty = total
End Function

Private Function SumInventoryQty( _
    ByRef inventory() As NormalizedInventoryLine, _
    ByVal shipmentNo As String, _
    ByVal sku As String) As Long

    Dim i As Long
    Dim total As Long

    If Not HasNormalizedInventoryRows(inventory) Then Exit Function

    For i = LBound(inventory) To UBound(inventory)
        If inventory(i).ShipmentNo = shipmentNo And inventory(i).SKU = sku And inventory(i).QtyValid Then
            total = total + inventory(i).Qty
        End If
    Next i

    SumInventoryQty = total
End Function

' -----------------------------------------------------------------------------
' 第4层：E11
' -----------------------------------------------------------------------------

Private Sub ApplyLayer4FragmentInventory( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine, _
    ByRef issues() As ValidationIssue, _
    ByRef failedShipments As Object, _
    ByRef shipmentHasE04 As Object, _
    ByRef shipmentHasE08 As Object)

    Dim groupKeys As Object
    Dim groupKey As Variant
    Dim parts() As String
    Dim shipmentNo As String
    Dim sku As String
    Dim groupMinQty As Long
    Dim qcTotals As Object
    Dim qcKey As Variant
    Dim totalQty As Long
    Dim issue As ValidationIssue

    Set groupKeys = CollectOrderShipmentSkuGroups(orders)

    For Each groupKey In groupKeys.Keys
        parts = Split(CStr(groupKey), vbTab)
        shipmentNo = parts(0)
        sku = parts(1)

        If shipmentHasE04.Exists(shipmentNo) Then GoTo NextGroup
        If shipmentHasE08.Exists(shipmentNo) Then GoTo NextGroup

        groupMinQty = CalcGroupMinQty(orders, shipmentNo, sku)
        If groupMinQty <= 0 Then GoTo NextGroup

        Set qcTotals = SumInventoryByQc(inventory, shipmentNo, sku)

        For Each qcKey In qcTotals.Keys
            totalQty = CLng(qcTotals(qcKey))

            If totalQty > 0 And totalQty < groupMinQty Then
                issue.ShipmentNo = shipmentNo
                issue.WMSOrderNo = NA_PLACEHOLDER
                issue.SKU = sku
                issue.ErrorCode = ERR_E11
                issue.SourceTable = NA_PLACEHOLDER
                issue.ExcelRowNum = 0
                issue.FieldName = "QC情况"
                issue.RawValue = CStr(qcKey) & ":" & CStr(totalQty)
                issue.Reason = "QC库存碎片无法分配（0 < T < groupMinQty）"

                AppendValidationIssue issues, issue, failedShipments
            End If
        Next qcKey
NextGroup:
    Next groupKey
End Sub

Private Function CollectOrderShipmentSkuGroups(ByRef orders() As NormalizedReturnLine) As Object
    Dim dict As Object
    Dim i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    If HasNormalizedReturnRows(orders) Then
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo <> vbNullString And orders(i).SKU <> vbNullString And orders(i).QtyValid Then
                key = orders(i).ShipmentNo & vbTab & orders(i).SKU
                dict(key) = True
            End If
        Next i
    End If

    Set CollectOrderShipmentSkuGroups = dict
End Function

Private Function CalcGroupMinQty( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipmentNo As String, _
    ByVal sku As String) As Long

    Dim i As Long
    Dim minQty As Long
    Dim initialized As Boolean

    If Not HasNormalizedReturnRows(orders) Then Exit Function

    For i = LBound(orders) To UBound(orders)
        If orders(i).ShipmentNo = shipmentNo And orders(i).SKU = sku And orders(i).QtyValid Then
            If Not initialized Or orders(i).Qty < minQty Then
                minQty = orders(i).Qty
                initialized = True
            End If
        End If
    Next i

    If initialized Then CalcGroupMinQty = minQty
End Function

Private Function SumInventoryByQc( _
    ByRef inventory() As NormalizedInventoryLine, _
    ByVal shipmentNo As String, _
    ByVal sku As String) As Object

    Dim dict As Object
    Dim i As Long
    Dim qc As String
    Dim current As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If Not HasNormalizedInventoryRows(inventory) Then
        Set SumInventoryByQc = dict
        Exit Function
    End If

    For i = LBound(inventory) To UBound(inventory)
        If inventory(i).ShipmentNo = shipmentNo And inventory(i).SKU = sku And inventory(i).QtyValid And inventory(i).QCValid Then
            qc = inventory(i).QC

            If dict.Exists(qc) Then
                current = CLng(dict(qc))
            Else
                current = 0
            End If

            dict(qc) = current + inventory(i).Qty
        End If
    Next i

    Set SumInventoryByQc = dict
End Function

' -----------------------------------------------------------------------------
' 通用工具
' -----------------------------------------------------------------------------

Private Sub AppendValidationIssue( _
    ByRef issues() As ValidationIssue, _
    ByRef issue As ValidationIssue, _
    ByRef failedShipments As Object)

    Dim nextIndex As Long

    On Error GoTo FirstIssue
    nextIndex = UBound(issues) + 1
    GoTo AppendIssue

FirstIssue:
    nextIndex = 1

AppendIssue:
    ReDim Preserve issues(1 To nextIndex)
    issues(nextIndex) = issue

    If issue.ShipmentNo <> vbNullString And issue.ShipmentNo <> NA_PLACEHOLDER Then
        failedShipments(issue.ShipmentNo) = True
    End If
End Sub

Private Function IsAnomalyDetailError(ByVal errorCode As String) As Boolean
    Select Case errorCode
        Case ERR_E01, ERR_E02, ERR_E03, ERR_E04, ERR_E05, ERR_E07
            IsAnomalyDetailError = True
        Case Else
            IsAnomalyDetailError = False
    End Select
End Function

Private Sub SortLongArray(ByRef values() As Long, ByRef indexes() As Long, ByVal count As Long)
    Dim i As Long
    Dim j As Long
    Dim tmpVal As Long
    Dim tmpIdx As Long

    For i = 1 To count - 1
        For j = i + 1 To count
            If values(j) < values(i) Then
                tmpVal = values(i)
                values(i) = values(j)
                values(j) = tmpVal

                tmpIdx = indexes(i)
                indexes(i) = indexes(j)
                indexes(j) = tmpIdx
            End If
        Next j
    Next i
End Sub

Private Function HasFieldIssues(ByRef issues() As FieldNormalizeIssue) As Boolean
    On Error GoTo NotAllocated
    HasFieldIssues = (UBound(issues) >= LBound(issues))
    Exit Function
NotAllocated:
    HasFieldIssues = False
End Function

Private Function HasNormalizedReturnRows(ByRef rows() As NormalizedReturnLine) As Boolean
    On Error GoTo NotAllocated
    HasNormalizedReturnRows = (UBound(rows) >= LBound(rows))
    Exit Function
NotAllocated:
    HasNormalizedReturnRows = False
End Function

Private Function HasNormalizedInventoryRows(ByRef rows() As NormalizedInventoryLine) As Boolean
    On Error GoTo NotAllocated
    HasNormalizedInventoryRows = (UBound(rows) >= LBound(rows))
    Exit Function
NotAllocated:
    HasNormalizedInventoryRows = False
End Function

Private Function HasValidationIssues(ByRef issues() As ValidationIssue) As Boolean
    On Error GoTo NotAllocated
    HasValidationIssues = (UBound(issues) >= LBound(issues))
    Exit Function
NotAllocated:
    HasValidationIssues = False
End Function

Public Function CountValidationIssues(ByRef issues() As ValidationIssue) As Long
    On Error GoTo NotAllocated
    CountValidationIssues = UBound(issues) - LBound(issues) + 1
    Exit Function
NotAllocated:
    CountValidationIssues = 0
End Function

Public Function CountAnomalyRows(ByRef rows() As AnomalyRow) As Long
    On Error GoTo NotAllocated
    CountAnomalyRows = UBound(rows) - LBound(rows) + 1
    Exit Function
NotAllocated:
    CountAnomalyRows = 0
End Function

Public Function ShipmentHasError( _
    ByRef issues() As ValidationIssue, _
    ByVal shipmentNo As String, _
    ByVal errorCode As String) As Boolean

    Dim i As Long

    If Not HasValidationIssues(issues) Then Exit Function

    For i = LBound(issues) To UBound(issues)
        If issues(i).ShipmentNo = shipmentNo And issues(i).ErrorCode = errorCode Then
            ShipmentHasError = True
            Exit Function
        End If
    Next i
End Function

Public Function CountIssuesByError( _
    ByRef issues() As ValidationIssue, _
    ByVal errorCode As String) As Long

    Dim i As Long

    If Not HasValidationIssues(issues) Then Exit Function

    For i = LBound(issues) To UBound(issues)
        If issues(i).ErrorCode = errorCode Then
            CountIssuesByError = CountIssuesByError + 1
        End If
    Next i
End Function
