Option Explicit

' =============================================================================
' M12_分配后校验（modPostValidate）
' =============================================================================
' 职责：对已经成功分配的物流单号做最后一道一致性检查：
'   1. 每个退单行的分配数量合计必须等于退单数量。
'   2. 同一个退单行不能同时使用两种 QC。
'   3. 成功明细中的关键字段必须能回到原始退单行。
'   4. 退单号状态必须等于其下所有行状态的聚合结果。
'   5. 已经整单回滚的物流单号不参与成功后校验。
'
' 给新手的解释：
'   M11 会把最终结果整理成一个 Dictionary（可以理解为“带标签的结果表”）。
'   本模块只读取这些标签，不再改动分配结果；如果发现问题，就把问题记录到
'   PostValidationResult(Dictionary) 里，供后续编排层决定是否终止或提示。
'
' 公开函数：
'   ValidatePost(orders(), finalResult) → Object(PostValidationResult)
'
' PostValidationResult(Dictionary) 键约定：
'   HasFailures As Boolean
'   IssueCount As Long
'   Issue_i_Code / ShipmentNo / WMSOrderNo / SKU / LineNo / Message
' =============================================================================

Private Const POST_ERR_QTY_MISMATCH As String = "POST_QTY_MISMATCH"
Private Const POST_ERR_QC_MISMATCH  As String = "POST_QC_MISMATCH"
Private Const POST_ERR_DATA_MISMATCH As String = "POST_DATA_MISMATCH"
Private Const POST_ERR_STATUS_MISMATCH As String = "POST_STATUS_MISMATCH"


' =============================================================================
' 一、公开函数
' =============================================================================

Public Function ValidatePost(ByRef orders() As NormalizedReturnLine, ByVal finalResult As Object) As Object
    Dim result As Object
    Set result = PV_CreateResult()

    Dim rollbackShipments As Object
    Set rollbackShipments = PV_CollectRollbackShipments(finalResult)

    Dim qtyByLine As Object
    Dim qcByLine As Object
    Dim qcConflictByLine As Object
    Dim orderQtyByLine As Object
    Dim summaryStatusByWms As Object
    Set qtyByLine = CreateObject("Scripting.Dictionary")
    Set qcByLine = CreateObject("Scripting.Dictionary")
    Set qcConflictByLine = CreateObject("Scripting.Dictionary")
    Set orderQtyByLine = CreateObject("Scripting.Dictionary")
    Set summaryStatusByWms = PV_CollectSummaryStatus(finalResult)

    PV_CollectOrderFacts orders, rollbackShipments, orderQtyByLine
    PV_CollectDetailFacts finalResult, qtyByLine, qcByLine, qcConflictByLine
    PV_CheckOrders orders, rollbackShipments, qtyByLine, qcConflictByLine, result
    PV_CheckDetailIntegrity finalResult, rollbackShipments, orderQtyByLine, summaryStatusByWms, result
    PV_CheckWMSStatus finalResult, rollbackShipments, summaryStatusByWms, result

    Set ValidatePost = result
End Function


' =============================================================================
' 二、核心校验流程
' =============================================================================

Private Sub PV_CheckOrders( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal rollbackShipments As Object, _
    ByVal qtyByLine As Object, _
    ByVal qcConflictByLine As Object, _
    ByVal result As Object)

    If PV_ReturnLineCount(orders) = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(orders) To UBound(orders)
        If PV_ShouldSkipShipment(orders(i).ShipmentNo, rollbackShipments) Then GoTo NextOrder

        Dim lineKey As String
        lineKey = PV_BuildLineKey( _
            orders(i).ShipmentNo, orders(i).WMSOrderNo, orders(i).SKU, orders(i).LineNo)

        Dim actualQty As Long
        actualQty = PV_GetDictLong(qtyByLine, lineKey)

        ' 数量守恒是最基础的后校验：输入要几件，最终成功明细就必须合计几件。
        If actualQty <> orders(i).Qty Then
            PV_AppendIssue result, POST_ERR_QTY_MISMATCH, orders(i), _
                "分配量合计 " & CStr(actualQty) & "，退单量 " & CStr(orders(i).Qty)
        End If

        ' 同一退单行允许拆批号/效期，但不能跨 QC，否则后续人工处理含义会变得不一致。
        If qcConflictByLine.Exists(lineKey) Then
            PV_AppendIssue result, POST_ERR_QC_MISMATCH, orders(i), _
                "同一退单行使用了多种 QC：" & CStr(qcConflictByLine(lineKey))
        End If

NextOrder:
    Next i
End Sub

Private Sub PV_CollectOrderFacts( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal rollbackShipments As Object, _
    ByVal orderQtyByLine As Object)

    If PV_ReturnLineCount(orders) = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(orders) To UBound(orders)
        If Not PV_ShouldSkipShipment(orders(i).ShipmentNo, rollbackShipments) Then
            orderQtyByLine(PV_BuildLineKey( _
                orders(i).ShipmentNo, orders(i).WMSOrderNo, orders(i).SKU, orders(i).LineNo)) = orders(i).Qty
        End If
    Next i
End Sub

Private Sub PV_CollectDetailFacts( _
    ByVal finalResult As Object, _
    ByVal qtyByLine As Object, _
    ByVal qcByLine As Object, _
    ByVal qcConflictByLine As Object)

    If finalResult Is Nothing Then Exit Sub

    Dim detailCount As Long
    detailCount = PV_GetResultLong(finalResult, "DetailCount")
    If detailCount <= 0 Then Exit Sub

    Dim i As Long
    For i = 1 To detailCount
        Dim lineKey As String
        lineKey = PV_BuildLineKey( _
            PV_GetResultText(finalResult, "Detail_" & i & "_ShipmentNo"), _
            PV_GetResultText(finalResult, "Detail_" & i & "_WMSOrderNo"), _
            PV_GetResultText(finalResult, "Detail_" & i & "_SKU"), _
            PV_GetResultText(finalResult, "Detail_" & i & "_LineNo"))

        Dim allocQty As Long
        allocQty = PV_GetResultLong(finalResult, "Detail_" & i & "_AllocQty")
        qtyByLine(lineKey) = PV_GetDictLong(qtyByLine, lineKey) + allocQty

        If allocQty > 0 Then
            PV_RecordLineQc lineKey, PV_GetResultText(finalResult, "Detail_" & i & "_QC"), _
                qcByLine, qcConflictByLine
        End If
    Next i
End Sub

Private Sub PV_CheckDetailIntegrity( _
    ByVal finalResult As Object, _
    ByVal rollbackShipments As Object, _
    ByVal orderQtyByLine As Object, _
    ByVal summaryStatusByWms As Object, _
    ByVal result As Object)

    If finalResult Is Nothing Then Exit Sub

    Dim detailCount As Long
    detailCount = PV_GetResultLong(finalResult, "DetailCount")
    If detailCount <= 0 Then Exit Sub

    Dim i As Long
    For i = 1 To detailCount
        Dim shipNo As String
        Dim wmsOrderNo As String
        Dim sku As String
        Dim lineNo As String
        shipNo = PV_GetResultText(finalResult, "Detail_" & i & "_ShipmentNo")
        wmsOrderNo = PV_GetResultText(finalResult, "Detail_" & i & "_WMSOrderNo")
        sku = PV_GetResultText(finalResult, "Detail_" & i & "_SKU")
        lineNo = PV_GetResultText(finalResult, "Detail_" & i & "_LineNo")

        If PV_ShouldSkipShipment(shipNo, rollbackShipments) Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "整单回滚物流单号不应出现在成功分配明细中"
            GoTo NextDetail
        End If

        If PV_IsBlank(shipNo) Or PV_IsBlank(wmsOrderNo) Or PV_IsBlank(sku) Or PV_IsBlank(lineNo) Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "成功分配明细关键字段为空"
            GoTo NextDetail
        End If

        Dim lineKey As String
        lineKey = PV_BuildLineKey(shipNo, wmsOrderNo, sku, lineNo)

        If Not orderQtyByLine.Exists(lineKey) Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "成功分配明细找不到对应退单行"
            GoTo NextDetail
        End If

        Dim expectedOrderQty As Long
        Dim actualOrderQty As Long
        expectedOrderQty = CLng(orderQtyByLine(lineKey))
        actualOrderQty = PV_GetResultLong(finalResult, "Detail_" & i & "_OrderQty")
        If actualOrderQty <> expectedOrderQty Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "成功明细退单数量 " & CStr(actualOrderQty) & "，输入退单数量 " & CStr(expectedOrderQty)
        End If

        If PV_GetResultLong(finalResult, "Detail_" & i & "_AllocQty") <= 0 Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "成功分配明细分配数量必须大于 0"
        End If

        If PV_IsBlank(PV_GetResultText(finalResult, "Detail_" & i & "_QC")) _
           Or PV_IsBlank(PV_GetResultText(finalResult, "Detail_" & i & "_LotNo")) _
           Or PV_IsBlank(PV_GetResultText(finalResult, "Detail_" & i & "_Expiry")) Then
            PV_AppendDetailIssue result, POST_ERR_DATA_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                "成功分配明细 QC/批号/效期不能为空"
        End If

        Dim detailWmsStatus As String
        detailWmsStatus = PV_GetResultText(finalResult, "Detail_" & i & "_WMSOrderStatus")
        If summaryStatusByWms.Exists(wmsOrderNo) Then
            If detailWmsStatus <> CStr(summaryStatusByWms(wmsOrderNo)) Then
                PV_AppendDetailIssue result, POST_ERR_STATUS_MISMATCH, shipNo, wmsOrderNo, sku, lineNo, _
                    "明细退单号状态 " & detailWmsStatus & "，汇总退单号状态 " & CStr(summaryStatusByWms(wmsOrderNo))
            End If
        End If

NextDetail:
    Next i
End Sub

Private Sub PV_CheckWMSStatus( _
    ByVal finalResult As Object, _
    ByVal rollbackShipments As Object, _
    ByVal summaryStatusByWms As Object, _
    ByVal result As Object)

    If finalResult Is Nothing Then Exit Sub

    Dim expectedByWms As Object
    Dim sampleShipByWms As Object
    Dim sampleSkuByWms As Object
    Dim sampleLineByWms As Object
    Set expectedByWms = CreateObject("Scripting.Dictionary")
    Set sampleShipByWms = CreateObject("Scripting.Dictionary")
    Set sampleSkuByWms = CreateObject("Scripting.Dictionary")
    Set sampleLineByWms = CreateObject("Scripting.Dictionary")

    Dim detailCount As Long
    detailCount = PV_GetResultLong(finalResult, "DetailCount")

    Dim i As Long
    For i = 1 To detailCount
        Dim shipNo As String
        Dim wmsOrderNo As String
        Dim lineStatus As String
        shipNo = PV_GetResultText(finalResult, "Detail_" & i & "_ShipmentNo")
        If PV_ShouldSkipShipment(shipNo, rollbackShipments) Then GoTo NextDetail

        wmsOrderNo = PV_GetResultText(finalResult, "Detail_" & i & "_WMSOrderNo")
        If PV_IsBlank(wmsOrderNo) Then GoTo NextDetail

        If Not sampleShipByWms.Exists(wmsOrderNo) Then
            sampleShipByWms(wmsOrderNo) = shipNo
            sampleSkuByWms(wmsOrderNo) = PV_GetResultText(finalResult, "Detail_" & i & "_SKU")
            sampleLineByWms(wmsOrderNo) = PV_GetResultText(finalResult, "Detail_" & i & "_LineNo")
        End If

        lineStatus = PV_GetResultText(finalResult, "Detail_" & i & "_LineStatus")
        Select Case lineStatus
            Case STATUS_MANUAL
                expectedByWms(wmsOrderNo) = STATUS_MANUAL
            Case STATUS_BATCH_IMPORT
                If Not expectedByWms.Exists(wmsOrderNo) Then expectedByWms(wmsOrderNo) = STATUS_BATCH_IMPORT
            Case Else
                PV_AppendDetailIssue result, POST_ERR_STATUS_MISMATCH, shipNo, wmsOrderNo, _
                    PV_GetResultText(finalResult, "Detail_" & i & "_SKU"), _
                    PV_GetResultText(finalResult, "Detail_" & i & "_LineNo"), _
                    "行状态非法：" & lineStatus
        End Select

NextDetail:
    Next i

    Dim wmsKey As Variant
    For Each wmsKey In expectedByWms.Keys
        Dim expectedStatus As String
        expectedStatus = CStr(expectedByWms(wmsKey))

        If Not summaryStatusByWms.Exists(CStr(wmsKey)) Then
            PV_AppendDetailIssue result, POST_ERR_STATUS_MISMATCH, _
                CStr(sampleShipByWms(wmsKey)), CStr(wmsKey), _
                CStr(sampleSkuByWms(wmsKey)), CStr(sampleLineByWms(wmsKey)), _
                "汇总表缺少该 WMS 退单号状态"
        ElseIf CStr(summaryStatusByWms(CStr(wmsKey))) <> expectedStatus Then
            PV_AppendDetailIssue result, POST_ERR_STATUS_MISMATCH, _
                CStr(sampleShipByWms(wmsKey)), CStr(wmsKey), _
                CStr(sampleSkuByWms(wmsKey)), CStr(sampleLineByWms(wmsKey)), _
                "行状态聚合应为 " & expectedStatus & "，汇总表实际为 " & CStr(summaryStatusByWms(CStr(wmsKey)))
        End If
    Next wmsKey
End Sub

Private Sub PV_RecordLineQc( _
    ByVal lineKey As String, _
    ByVal qcValue As String, _
    ByVal qcByLine As Object, _
    ByVal qcConflictByLine As Object)

    If Len(qcValue) = 0 Then Exit Sub

    If Not qcByLine.Exists(lineKey) Then
        qcByLine(lineKey) = qcValue
        Exit Sub
    End If

    If CStr(qcByLine(lineKey)) <> qcValue Then
        qcConflictByLine(lineKey) = CStr(qcByLine(lineKey)) & "," & qcValue
    End If
End Sub


' =============================================================================
' 三、整单回滚跳过规则
' =============================================================================

Private Function PV_CollectRollbackShipments(ByVal finalResult As Object) As Object
    Dim shipments As Object
    Set shipments = CreateObject("Scripting.Dictionary")

    If finalResult Is Nothing Then
        Set PV_CollectRollbackShipments = shipments
        Exit Function
    End If

    Dim summaryCount As Long
    summaryCount = PV_GetResultLong(finalResult, "SummaryCount")

    Dim i As Long
    For i = 1 To summaryCount
        If PV_GetResultText(finalResult, "Summary_" & i & "_Status") = STATUS_UNALLOCATED Then
            shipments(PV_GetResultText(finalResult, "Summary_" & i & "_ShipmentNo")) = True
        End If
    Next i

    Set PV_CollectRollbackShipments = shipments
End Function

Private Function PV_CollectSummaryStatus(ByVal finalResult As Object) As Object
    Dim statusByWms As Object
    Set statusByWms = CreateObject("Scripting.Dictionary")

    If finalResult Is Nothing Then
        Set PV_CollectSummaryStatus = statusByWms
        Exit Function
    End If

    Dim summaryCount As Long
    summaryCount = PV_GetResultLong(finalResult, "SummaryCount")

    Dim i As Long
    For i = 1 To summaryCount
        Dim wmsOrderNo As String
        wmsOrderNo = PV_GetResultText(finalResult, "Summary_" & i & "_WMSOrderNo")
        If Len(wmsOrderNo) > 0 Then
            statusByWms(wmsOrderNo) = PV_GetResultText(finalResult, "Summary_" & i & "_Status")
        End If
    Next i

    Set PV_CollectSummaryStatus = statusByWms
End Function

Private Function PV_ShouldSkipShipment(ByVal shipNo As String, ByVal rollbackShipments As Object) As Boolean
    If rollbackShipments Is Nothing Then Exit Function
    PV_ShouldSkipShipment = rollbackShipments.Exists(shipNo)
End Function


' =============================================================================
' 四、结果构造与通用取值
' =============================================================================

Private Function PV_CreateResult() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "HasFailures", False
    result.Add "IssueCount", CLng(0)
    Set PV_CreateResult = result
End Function

Private Sub PV_AppendIssue( _
    ByVal result As Object, _
    ByVal issueCode As String, _
    ByRef orderLine As NormalizedReturnLine, _
    ByVal message As String)

    Dim issueIndex As Long
    issueIndex = CLng(result("IssueCount")) + 1
    result("IssueCount") = issueIndex
    result("HasFailures") = True

    result.Add "Issue_" & issueIndex & "_Code", issueCode
    result.Add "Issue_" & issueIndex & "_ShipmentNo", orderLine.ShipmentNo
    result.Add "Issue_" & issueIndex & "_WMSOrderNo", orderLine.WMSOrderNo
    result.Add "Issue_" & issueIndex & "_SKU", orderLine.SKU
    result.Add "Issue_" & issueIndex & "_LineNo", orderLine.LineNo
    result.Add "Issue_" & issueIndex & "_Message", message
End Sub

Private Sub PV_AppendDetailIssue( _
    ByVal result As Object, _
    ByVal issueCode As String, _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String, _
    ByVal message As String)

    Dim issueIndex As Long
    issueIndex = CLng(result("IssueCount")) + 1
    result("IssueCount") = issueIndex
    result("HasFailures") = True

    result.Add "Issue_" & issueIndex & "_Code", issueCode
    result.Add "Issue_" & issueIndex & "_ShipmentNo", shipNo
    result.Add "Issue_" & issueIndex & "_WMSOrderNo", wmsOrderNo
    result.Add "Issue_" & issueIndex & "_SKU", sku
    result.Add "Issue_" & issueIndex & "_LineNo", lineNo
    result.Add "Issue_" & issueIndex & "_Message", message
End Sub

Private Function PV_BuildLineKey( _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String) As String

    PV_BuildLineKey = shipNo & vbTab & wmsOrderNo & vbTab & sku & vbTab & lineNo
End Function

Private Function PV_IsBlank(ByVal value As String) As Boolean
    PV_IsBlank = (Len(Trim$(value)) = 0)
End Function

Private Function PV_GetDictLong(ByVal dict As Object, ByVal keyName As String) As Long
    If dict Is Nothing Then Exit Function
    If dict.Exists(keyName) Then PV_GetDictLong = CLng(dict(keyName))
End Function

Private Function PV_GetResultLong(ByVal dict As Object, ByVal keyName As String) As Long
    If dict Is Nothing Then Exit Function
    If dict.Exists(keyName) Then PV_GetResultLong = CLng(dict(keyName))
End Function

Private Function PV_GetResultText(ByVal dict As Object, ByVal keyName As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(keyName) Then PV_GetResultText = CStr(dict(keyName))
End Function

Private Function PV_ReturnLineCount(ByRef orders() As NormalizedReturnLine) As Long
    On Error GoTo EmptyOrders
    PV_ReturnLineCount = UBound(orders) - LBound(orders) + 1
    Exit Function

EmptyOrders:
    PV_ReturnLineCount = 0
End Function
