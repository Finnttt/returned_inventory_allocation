    Option Explicit

    ' =============================================================================
    ' M11_状态判定（modStatus）
    ' =============================================================================
    ' 职责：接收 M09 分配结果与 M05 校验结果，执行：
    '   1. 行级状态判定（批号+效期组合数 → 批量导入 / 手工操作）
    '   2. 物流单号级整单回滚（任一 SKU 组失败则丢弃该单全部成功明细）
    '   3. 退单号状态聚合（手工操作优先；失败单 → 无法分配 + 原因字段）
    '
    ' 给新手的解释：
    '   分配引擎（M09）只负责"怎么分货"；本模块负责"这次分货最终算成功还是失败、
    '   文员应该批量导入还是手工录入"。
    '
    ' 公开函数（与规格 modStatus 一致；VBA 复杂结构用 Object/Dictionary 承载）：
    '   DetermineLineStatus(details(), lineNo) → String
    '   ApplyRollback(shipmentResults(), validationResult, validationIssues()) → Object(FinalResult)
    '   AggregateWMSStatus(finalResult) → WMSStatusEntry() 即 WMSStatusMap
    '   BuildRollbackReason(directCodes(), triggerCode) → String
    '
    ' FinalResult（Object/Dictionary）键约定：
    '   SummaryCount
    '   Summary_i_ShipmentNo / WMSOrderNo / Status / Reason
    '   DetailCount（仅整单成功的物流单号有明细）
    '   Detail_i_*（AllocationDetail 字段 + WMSOrderStatus）
    '
    ' 说明：规格 ApplyRollback 仅写 ValidationResult，但 M05 的问题明细在
    ' ValidationIssue[] 中（与 ValidatePre 模式相同），因此增加 validationIssues 参数。
    ' shipmentResults() 每个元素为 M09 AllocateShipment 返回的 Dictionary。
    ' =============================================================================


    ' =============================================================================
    ' 一、公开函数
    ' =============================================================================

' 根据某退货行在分配明细中实际使用的"批号+效期"组合种类数判定行状态。
' 组合数 = 1 → 批量导入；≥ 2 → 手工操作；无有效明细 → 分配失败。
' 注意：此公开函数只按 lineNo 统计，适合单 WMS/单 SKU 的简单测试。
' 生产聚合路径必须使用完整退货行身份（物流单号+WMS退单号+SKU+行号），避免不同退单号的 00001 互相污染。
    Public Function DetermineLineStatus(ByRef details() As AllocationDetail, ByVal lineNo As String) As String
        Dim comboCount As Long
        comboCount = ST_CountLotExpiryCombos(details, lineNo)

        If comboCount = 0 Then
            DetermineLineStatus = LINE_STATUS_FAILED
        ElseIf comboCount = 1 Then
            DetermineLineStatus = STATUS_BATCH_IMPORT
        Else
            DetermineLineStatus = STATUS_MANUAL
        End If
    End Function

    ' 构造"无法分配"原因字符串。
    ' directCodes 为空 → 连带回滚格式；非空 → 按 E01→E99 升序去重后拼接直接原因。
    Public Function BuildRollbackReason(ByRef directCodes() As String, ByVal triggerCode As String) As String
        If ST_StringArrayCount(directCodes) = 0 Then
            BuildRollbackReason = "整单回滚（触发原因：" & triggerCode & "）"
        Else
            BuildRollbackReason = ST_FormatDirectCodes(directCodes)
        End If
    End Function

    ' 汇总校验失败与分配失败，执行整单回滚，产出 FinalResult（Object/Dictionary）。
    ' orders() 可传空数组；用于 E08 等错误把 [N/A] 展开到具体 WMS 退单号。
    ' 注意：① VBA 不允许 Optional + ByRef；② ByRef 数组参数必须传变量，不能传函数返回值表达式。
    Public Function ApplyRollback( _
        ByRef shipmentResults() As Object, _
        validationResult As ValidationResult, _
        ByRef validationIssues() As ValidationIssue, _
        ByRef orders() As NormalizedReturnLine) As Object

        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")

        Dim shipmentNos As Object
        Set shipmentNos = ST_CollectAllShipmentNos(shipmentResults, validationIssues)

        Dim summaryCount As Long
        Dim detailCount As Long
        Dim shipKey As Variant

        For Each shipKey In shipmentNos.Keys
            Dim shipNo As String
            shipNo = CStr(shipKey)

            Dim allocMap As Object
            Set allocMap = ST_FindShipmentResult(shipmentResults, shipNo)

            Dim validationFailed As Boolean
            validationFailed = ST_ShipmentHasValidationIssue(validationIssues, shipNo)

            Dim allocationFailed As Boolean
            allocationFailed = ST_ShipmentAllocationFailed(allocMap)

            If validationFailed Or allocationFailed Then
                ' 整单回滚：不写成功明细，只写汇总表"无法分配"行
                ST_AppendFailureSummary result, summaryCount, shipNo, validationIssues, allocMap, orders
            Else
                ' 全部 SKU 组成功：写入明细并聚合退单号状态
                ST_AppendSuccessOutput result, summaryCount, detailCount, shipNo, allocMap
            End If
        Next shipKey

        result.Add "SummaryCount", summaryCount
        result.Add "DetailCount", detailCount

        Set ApplyRollback = result
    End Function

    ' 从 FinalResult 中提取 WMSStatusMap（WMSStatusEntry 数组）。
    Public Function AggregateWMSStatus(ByVal finalResult As Object) As WMSStatusEntry()
        Dim entries() As WMSStatusEntry
        Dim count As Long
        Dim i As Long

        If finalResult Is Nothing Then
            AggregateWMSStatus = entries
            Exit Function
        End If

        If Not finalResult.Exists("SummaryCount") Then
            AggregateWMSStatus = entries
            Exit Function
        End If

        count = CLng(finalResult("SummaryCount"))
        If count <= 0 Then
            AggregateWMSStatus = entries
            Exit Function
        End If

        ReDim entries(1 To count)
        For i = 1 To count
            entries(i).ShipmentNo = CStr(finalResult("Summary_" & i & "_ShipmentNo"))
            entries(i).WMSOrderNo = CStr(finalResult("Summary_" & i & "_WMSOrderNo"))
            entries(i).Status = CStr(finalResult("Summary_" & i & "_Status"))
            entries(i).Reason = CStr(finalResult("Summary_" & i & "_Reason"))
        Next i

        AggregateWMSStatus = entries
    End Function


    ' =============================================================================
    ' 二、整单回滚 / 成功输出
    ' =============================================================================

    Private Sub ST_AppendFailureSummary( _
        ByRef result As Object, _
        ByRef summaryCount As Long, _
        ByVal shipNo As String, _
        ByRef validationIssues() As ValidationIssue, _
        ByVal allocMap As Object, _
        ByRef orders() As NormalizedReturnLine)

        Dim wmsOrders As Object
        Set wmsOrders = ST_CollectWmsOrders(shipNo, validationIssues, allocMap, orders)

        Dim triggerCode As String
        triggerCode = ST_FindTriggerCode(validationIssues, allocMap, shipNo)

        Dim wmsKey As Variant
        For Each wmsKey In wmsOrders.Keys
            Dim wmsOrderNo As String
            wmsOrderNo = CStr(wmsKey)

            Dim directCodes() As String
            directCodes = ST_CollectDirectCodesForWms( _
                shipNo, wmsOrderNo, validationIssues, allocMap, orders)

            Dim reason As String
            If ST_StringArrayCount(directCodes) > 0 Then
                reason = ST_FormatDirectCodes(directCodes)
            Else
                reason = BuildRollbackReason(directCodes, triggerCode)
            End If

            summaryCount = summaryCount + 1
            result.Add "Summary_" & summaryCount & "_ShipmentNo", shipNo
            result.Add "Summary_" & summaryCount & "_WMSOrderNo", wmsOrderNo
            result.Add "Summary_" & summaryCount & "_Status", STATUS_UNALLOCATED
            result.Add "Summary_" & summaryCount & "_Reason", reason
        Next wmsKey
    End Sub

    Private Sub ST_AppendSuccessOutput( _
        ByRef result As Object, _
        ByRef summaryCount As Long, _
        ByRef detailCount As Long, _
        ByVal shipNo As String, _
        ByVal allocMap As Object)

        If allocMap Is Nothing Then Exit Sub

        Dim allDetails() As AllocationDetail
        Dim allDetailCount As Long
        allDetailCount = ST_ExtractDetailsFromAllocMap(allocMap, shipNo, allDetails)

        If allDetailCount = 0 Then Exit Sub

        ' 行状态以"同一退货行实际使用的批号+效期组合数"为准。
        ' 注意：不同 WMS 退单号都会有 00001，不能只按行号统计，否则会把不同退货行误合并。
        Dim i As Long
        For i = 1 To allDetailCount
            allDetails(i).LineStatus = ST_DetermineDetailLineStatus(allDetails, allDetails(i))
        Next i

        ' 先聚合每个 WMS 退单号的退单号状态
        Dim wmsStatus As Object
        Set wmsStatus = CreateObject("Scripting.Dictionary")
        ST_BuildWmsStatusMap allDetails, allDetailCount, wmsStatus

        ' 写成功明细
        For i = 1 To allDetailCount
            detailCount = detailCount + 1
            With allDetails(i)
                result.Add "Detail_" & detailCount & "_ShipmentNo", .ShipmentNo
                result.Add "Detail_" & detailCount & "_WMSOrderNo", .WMSOrderNo
                result.Add "Detail_" & detailCount & "_SKU", .SKU
                result.Add "Detail_" & detailCount & "_LineNo", .LineNo
                result.Add "Detail_" & detailCount & "_OrderQty", .OrderQty
                result.Add "Detail_" & detailCount & "_QC", .QC
                result.Add "Detail_" & detailCount & "_LotNo", .LotNo
                result.Add "Detail_" & detailCount & "_Expiry", .Expiry
                result.Add "Detail_" & detailCount & "_AllocQty", .AllocQty
                result.Add "Detail_" & detailCount & "_LineStatus", .LineStatus
                result.Add "Detail_" & detailCount & "_WMSOrderStatus", CStr(wmsStatus(.WMSOrderNo))
            End With
        Next i

        ' 写成功汇总（每个 WMS 退单号一行，原因为空）
        Dim wmsKey As Variant
        For Each wmsKey In wmsStatus.Keys
            summaryCount = summaryCount + 1
            result.Add "Summary_" & summaryCount & "_ShipmentNo", shipNo
            result.Add "Summary_" & summaryCount & "_WMSOrderNo", CStr(wmsKey)
            result.Add "Summary_" & summaryCount & "_Status", CStr(wmsStatus(wmsKey))
            result.Add "Summary_" & summaryCount & "_Reason", vbNullString
        Next wmsKey
    End Sub


    ' =============================================================================
    ' 三、行状态与原因格式
    ' =============================================================================

    Private Function ST_CountLotExpiryCombos( _
        ByRef details() As AllocationDetail, _
        ByVal lineNo As String) As Long

        Dim combos As Object
        Set combos = CreateObject("Scripting.Dictionary")

        On Error GoTo EmptyDetails
        Dim i As Long
        For i = LBound(details) To UBound(details)
            If details(i).LineNo = lineNo And details(i).AllocQty > 0 Then
                combos(details(i).LotNo & vbTab & details(i).Expiry) = True
            End If
        Next i

    EmptyDetails:
        ST_CountLotExpiryCombos = combos.Count
    End Function

    Private Function ST_DetermineDetailLineStatus( _
        ByRef details() As AllocationDetail, _
        ByRef target As AllocationDetail) As String

        Dim comboCount As Long
        comboCount = ST_CountLotExpiryCombosForDetail(details, target)

        If comboCount = 0 Then
            ST_DetermineDetailLineStatus = LINE_STATUS_FAILED
        ElseIf comboCount = 1 Then
            ST_DetermineDetailLineStatus = STATUS_BATCH_IMPORT
        Else
            ST_DetermineDetailLineStatus = STATUS_MANUAL
        End If
    End Function

    Private Function ST_CountLotExpiryCombosForDetail( _
        ByRef details() As AllocationDetail, _
        ByRef target As AllocationDetail) As Long

        Dim combos As Object
        Set combos = CreateObject("Scripting.Dictionary")

        On Error GoTo EmptyDetails
        Dim i As Long
        For i = LBound(details) To UBound(details)
            If ST_IsSameReturnLine(details(i), target) And details(i).AllocQty > 0 Then
                combos(details(i).LotNo & vbTab & details(i).Expiry) = True
            End If
        Next i

    EmptyDetails:
        ST_CountLotExpiryCombosForDetail = combos.Count
    End Function

    Private Function ST_IsSameReturnLine( _
        ByRef leftDetail As AllocationDetail, _
        ByRef rightDetail As AllocationDetail) As Boolean

        ST_IsSameReturnLine = _
            (leftDetail.ShipmentNo = rightDetail.ShipmentNo) And _
            (leftDetail.WMSOrderNo = rightDetail.WMSOrderNo) And _
            (leftDetail.SKU = rightDetail.SKU) And _
            (leftDetail.LineNo = rightDetail.LineNo)
    End Function

    Private Function ST_FormatDirectCodes(ByRef codes() As String) As String
        If ST_StringArrayCount(codes) = 0 Then
            ST_FormatDirectCodes = vbNullString
            Exit Function
        End If

        Dim sorted() As String
        sorted = ST_SortUniqueCodes(codes)

        If ST_StringArrayCount(sorted) = 0 Then
            ST_FormatDirectCodes = vbNullString
            Exit Function
        End If

        Dim parts() As String
        ReDim parts(LBound(sorted) To UBound(sorted))

        Dim i As Long
        For i = LBound(sorted) To UBound(sorted)
            parts(i) = sorted(i) & " - " & ST_GetStandardReasonText(sorted(i))
        Next i

        ST_FormatDirectCodes = ST_JoinStringArray(parts)
    End Function

    ' Join 对数组下标敏感；改用手动拼接，避免 0 基/1 基数组混用导致运行时错误。
    Private Function ST_JoinStringArray(ByRef parts() As String) As String
        Dim result As String
        Dim i As Long

        On Error GoTo EmptyParts
        For i = LBound(parts) To UBound(parts)
            If Len(result) = 0 Then
                result = parts(i)
            Else
                result = result & "; " & parts(i)
            End If
        Next i
    EmptyParts:
        ST_JoinStringArray = result
    End Function

    Private Function ST_GetStandardReasonText(ByVal errorCode As String) As String
        Select Case errorCode
            Case ERR_E01: ST_GetStandardReasonText = "关键字段为空或格式异常"
            Case ERR_E02: ST_GetStandardReasonText = "退单表行号重复或不连续"
            Case ERR_E03: ST_GetStandardReasonText = "QC情况非法"
            Case ERR_E04: ST_GetStandardReasonText = "数量非法"
            Case ERR_E05: ST_GetStandardReasonText = "效期格式非法"
            Case ERR_E06: ST_GetStandardReasonText = "物流单号仅存在于退单表"
            Case ERR_E07: ST_GetStandardReasonText = "物流单号仅存在于质检库存表"
            Case ERR_E08: ST_GetStandardReasonText = "同物流单号+SKU数量不一致"
            Case ERR_E09: ST_GetStandardReasonText = "分配路径穷尽"
            Case ERR_E10: ST_GetStandardReasonText = "回溯超限"
            Case ERR_E11: ST_GetStandardReasonText = "QC库存碎片无法分配"
            Case ERR_E99: ST_GetStandardReasonText = "未知异常"
            Case Else:    ST_GetStandardReasonText = "未知错误"
        End Select
    End Function

    Private Function ST_IsSplitReasonErrorCode(ByVal errorCode As String) As Boolean
        Select Case errorCode
            Case ERR_E08, ERR_E09, ERR_E10, ERR_E11, ERR_E99
                ST_IsSplitReasonErrorCode = True
            Case Else
                ST_IsSplitReasonErrorCode = False
        End Select
    End Function

    Private Function ST_IsDirectAllocErrorCode(ByVal errorCode As String) As Boolean
        Select Case errorCode
            Case ERR_E09, ERR_E10, ERR_E99
                ST_IsDirectAllocErrorCode = True
            Case Else
                ST_IsDirectAllocErrorCode = False
        End Select
    End Function


    ' =============================================================================
    ' 四、物流单号 / WMS 退单号收集
    ' =============================================================================

    Private Function ST_CollectAllShipmentNos( _
        ByRef shipmentResults() As Object, _
        ByRef validationIssues() As ValidationIssue) As Object

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        Dim i As Long
        If ST_SafeObjectArrayCount(shipmentResults) > 0 Then
            For i = LBound(shipmentResults) To UBound(shipmentResults)
                If Not shipmentResults(i) Is Nothing Then
                    If shipmentResults(i).Exists("ShipmentNo") Then
                        dict(CStr(shipmentResults(i)("ShipmentNo"))) = True
                    End If
                End If
            Next i
        End If

        If HasValidationIssuesArray(validationIssues) Then
            For i = LBound(validationIssues) To UBound(validationIssues)
                If validationIssues(i).ShipmentNo <> vbNullString Then
                    dict(validationIssues(i).ShipmentNo) = True
                End If
            Next i
        End If

        Set ST_CollectAllShipmentNos = dict
    End Function

    Private Function ST_FindShipmentResult( _
        ByRef shipmentResults() As Object, _
        ByVal shipNo As String) As Object

        If ST_SafeObjectArrayCount(shipmentResults) = 0 Then
            Set ST_FindShipmentResult = Nothing
            Exit Function
        End If

        Dim i As Long
        For i = LBound(shipmentResults) To UBound(shipmentResults)
            If Not shipmentResults(i) Is Nothing Then
                If shipmentResults(i).Exists("ShipmentNo") Then
                    If CStr(shipmentResults(i)("ShipmentNo")) = shipNo Then
                        Set ST_FindShipmentResult = shipmentResults(i)
                        Exit Function
                    End If
                End If
            End If
        Next i

        Set ST_FindShipmentResult = Nothing
    End Function

    Private Function ST_ShipmentHasValidationIssue( _
        ByRef validationIssues() As ValidationIssue, _
        ByVal shipNo As String) As Boolean

        Dim i As Long
        If Not HasValidationIssuesArray(validationIssues) Then Exit Function

        For i = LBound(validationIssues) To UBound(validationIssues)
            If validationIssues(i).ShipmentNo = shipNo Then
                ST_ShipmentHasValidationIssue = True
                Exit Function
            End If
        Next i
    End Function

    Private Function ST_ShipmentAllocationFailed(ByVal allocMap As Object) As Boolean
        If allocMap Is Nothing Then Exit Function
        If Not allocMap.Exists("GroupCount") Then Exit Function

        Dim groupCount As Long
        Dim g As Long
        groupCount = CLng(allocMap("GroupCount"))

        For g = 1 To groupCount
            If allocMap.Exists("Group_" & g & "_Success") Then
                If Not CBool(allocMap("Group_" & g & "_Success")) Then
                    ST_ShipmentAllocationFailed = True
                    Exit Function
                End If
            End If
        Next g
    End Function

    Private Function ST_CollectWmsOrders( _
        ByVal shipNo As String, _
        ByRef validationIssues() As ValidationIssue, _
        ByVal allocMap As Object, _
        ByRef orders() As NormalizedReturnLine) As Object

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        Dim i As Long
        If HasValidationIssuesArray(validationIssues) Then
            For i = LBound(validationIssues) To UBound(validationIssues)
                If validationIssues(i).ShipmentNo = shipNo Then
                    If validationIssues(i).WMSOrderNo <> vbNullString _
                    And validationIssues(i).WMSOrderNo <> NA_PLACEHOLDER Then
                        dict(validationIssues(i).WMSOrderNo) = True
                    ElseIf validationIssues(i).WMSOrderNo = NA_PLACEHOLDER _
                        And validationIssues(i).SourceTable = SOURCE_RETURN_TABLE _
                        And validationIssues(i).ExcelRowNum > 0 Then
                        ' 退单表真实数据行的 WMS 为空时，除正常 WMS 汇总外，
                        ' 还必须保留一条 [N/A]，让文员能看到并定位该异常行。
                        ' E06 等整单级问题 ExcelRowNum=0，不应额外生成 [N/A]。
                        dict(NA_PLACEHOLDER) = True
                    End If
                End If
            Next i
        End If

        ST_AddWmsFromAllocMap dict, allocMap
        ST_AddWmsFromOrders dict, shipNo, orders

        ' E07 / E06 / E08 等可能只有 [N/A]：至少保留占位行
        If dict.Count = 0 Then
            dict(NA_PLACEHOLDER) = True
        End If

        Set ST_CollectWmsOrders = dict
    End Function

    Private Sub ST_AddWmsFromAllocMap(ByVal dict As Object, ByVal allocMap As Object)
        If allocMap Is Nothing Then Exit Sub
        If Not allocMap.Exists("GroupCount") Then Exit Sub

        Dim groupCount As Long
        Dim g As Long
        Dim d As Long
        groupCount = CLng(allocMap("GroupCount"))

        For g = 1 To groupCount
            If Not allocMap.Exists("Group_" & g & "_DetailCount") Then GoTo NextGroup
            Dim detailCount As Long
            detailCount = CLng(allocMap("Group_" & g & "_DetailCount"))
            For d = 1 To detailCount
                Dim keyName As String
                keyName = "Group_" & g & "_WMSOrderNo_" & d
                If allocMap.Exists(keyName) Then
                    dict(CStr(allocMap(keyName))) = True
                End If
            Next d
    NextGroup:
        Next g
    End Sub

    Private Sub ST_AddWmsFromOrders( _
        ByVal dict As Object, _
        ByVal shipNo As String, _
        ByRef orders() As NormalizedReturnLine)

        On Error GoTo Done
        Dim i As Long
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo = shipNo And orders(i).WMSOrderNo <> vbNullString Then
                dict(orders(i).WMSOrderNo) = True
            End If
        Next i
    Done:
    End Sub

    Private Function ST_FindTriggerCode( _
        ByRef validationIssues() As ValidationIssue, _
        ByVal allocMap As Object, _
        ByVal shipNo As String) As String

        Dim bestCode As String
        bestCode = vbNullString

        Dim i As Long
        If HasValidationIssuesArray(validationIssues) Then
            For i = LBound(validationIssues) To UBound(validationIssues)
                If validationIssues(i).ShipmentNo = shipNo Then
                    bestCode = ST_PickEarlierErrorCode(bestCode, validationIssues(i).ErrorCode)
                End If
            Next i
        End If

        If Not allocMap Is Nothing Then
            If allocMap.Exists("GroupCount") Then
                Dim g As Long
                For g = 1 To CLng(allocMap("GroupCount"))
                    If allocMap.Exists("Group_" & g & "_ErrorCode") Then
                        Dim ec As String
                        ec = CStr(allocMap("Group_" & g & "_ErrorCode"))
                        If ST_IsDirectAllocErrorCode(ec) Then
                            bestCode = ST_PickEarlierErrorCode(bestCode, ec)
                        End If
                    End If
                Next g
            End If
        End If

        If bestCode = vbNullString Then bestCode = ERR_E09
        ST_FindTriggerCode = bestCode
    End Function

    Private Function ST_CollectDirectCodesForWms( _
        ByVal shipNo As String, _
        ByVal wmsOrderNo As String, _
        ByRef validationIssues() As ValidationIssue, _
        ByVal allocMap As Object, _
        ByRef orders() As NormalizedReturnLine) As String()

        Dim codeDict As Object
        Set codeDict = CreateObject("Scripting.Dictionary")

        Dim i As Long
        If HasValidationIssuesArray(validationIssues) Then
            For i = LBound(validationIssues) To UBound(validationIssues)
                If validationIssues(i).ShipmentNo <> shipNo Then GoTo NextValidationIssue

                If ST_ValidationIssueAppliesToWms(validationIssues(i), wmsOrderNo, orders) Then
                    If Not ST_IsSplitReasonErrorCode(validationIssues(i).ErrorCode) _
                    Or ST_ValidationIssueAppliesToWmsDirectly(validationIssues(i), wmsOrderNo, orders) Then
                        codeDict(validationIssues(i).ErrorCode) = True
                    End If
                End If
    NextValidationIssue:
            Next i
        End If

        ST_AddAllocDirectCodes codeDict, shipNo, wmsOrderNo, allocMap, orders

        ST_CollectDirectCodesForWms = ST_DictionaryKeysToStringArray(codeDict)
    End Function

    Private Function ST_ValidationIssueAppliesToWms( _
        ByRef issue As ValidationIssue, _
        ByVal wmsOrderNo As String, _
        ByRef orders() As NormalizedReturnLine) As Boolean

        ' E07 等孤立物流单号：汇总行 WMS=[N/A]
        If wmsOrderNo = NA_PLACEHOLDER Then
            If issue.WMSOrderNo = NA_PLACEHOLDER Then
                ST_ValidationIssueAppliesToWms = True
            End If
            Exit Function
        End If

        If issue.WMSOrderNo = wmsOrderNo Then
            ST_ValidationIssueAppliesToWms = True
            Exit Function
        End If

        ' 物流单号+SKU 级错误（WMS 为 [N/A]）：扩展到该 SKU 下所有 WMS 退单号
        If issue.WMSOrderNo = NA_PLACEHOLDER And issue.SKU <> NA_PLACEHOLDER Then
            ST_ValidationIssueAppliesToWms = ST_WmsHasSku(wmsOrderNo, issue.ShipmentNo, issue.SKU, orders, Nothing)
            Exit Function
        End If

        ' E06 等整单级错误：扩展到该物流单号下所有 WMS 退单号
        If issue.WMSOrderNo = NA_PLACEHOLDER Then
            ST_ValidationIssueAppliesToWms = ST_WmsBelongsToShipment(wmsOrderNo, issue.ShipmentNo, orders)
        End If
    End Function

    Private Function ST_ValidationIssueAppliesToWmsDirectly( _
        ByRef issue As ValidationIssue, _
        ByVal wmsOrderNo As String, _
        ByRef orders() As NormalizedReturnLine) As Boolean

        ' E01~E07 行级/退单号级：Issue 已带具体 WMS
        If Not ST_IsSplitReasonErrorCode(issue.ErrorCode) Then
            ST_ValidationIssueAppliesToWmsDirectly = ST_ValidationIssueAppliesToWms(issue, wmsOrderNo, orders)
            Exit Function
        End If

        ' E08/E11：Issue 挂在物流单号+SKU，退单号下含该 SKU 即视为直接原因
        If issue.WMSOrderNo = NA_PLACEHOLDER And issue.SKU <> NA_PLACEHOLDER Then
            ST_ValidationIssueAppliesToWmsDirectly = ST_WmsHasSku(wmsOrderNo, issue.ShipmentNo, issue.SKU, orders, Nothing)
        End If
    End Function

    Private Sub ST_AddAllocDirectCodes( _
        ByVal codeDict As Object, _
        ByVal shipNo As String, _
        ByVal wmsOrderNo As String, _
        ByVal allocMap As Object, _
        ByRef orders() As NormalizedReturnLine)

        If allocMap Is Nothing Then Exit Sub
        If Not allocMap.Exists("GroupCount") Then Exit Sub

        Dim g As Long
        For g = 1 To CLng(allocMap("GroupCount"))
            If Not allocMap.Exists("Group_" & g & "_ErrorCode") Then GoTo NextGroup
            Dim ec As String
            ec = CStr(allocMap("Group_" & g & "_ErrorCode"))
            If Not ST_IsDirectAllocErrorCode(ec) Then GoTo NextGroup

            Dim sku As String
            sku = vbNullString
            If allocMap.Exists("Group_" & g & "_SKU") Then sku = CStr(allocMap("Group_" & g & "_SKU"))

            If ST_WmsHasSku(wmsOrderNo, shipNo, sku, orders, allocMap) Then
                codeDict(ec) = True
            End If
    NextGroup:
        Next g
    End Sub

    Private Function ST_WmsHasSku( _
        ByVal wmsOrderNo As String, _
        ByVal shipNo As String, _
        ByVal sku As String, _
        ByRef orders() As NormalizedReturnLine, _
        Optional ByVal allocMap As Object) As Boolean

        On Error GoTo NoOrders
        Dim i As Long
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo = shipNo _
            And orders(i).WMSOrderNo = wmsOrderNo _
            And orders(i).SKU = sku Then
                ST_WmsHasSku = True
                Exit Function
            End If
        Next i
    NoOrders:

        ' 未传 orders 时，从分配结果字典中反查 WMS+SKU 关系
        If sku <> vbNullString Then
            If Not allocMap Is Nothing Then
                ST_WmsHasSku = ST_AllocMapLinksWmsSku(allocMap, wmsOrderNo, sku)
            End If
        End If
    End Function

    Private Function ST_AllocMapLinksWmsSku( _
        ByVal allocMap As Object, _
        ByVal wmsOrderNo As String, _
        ByVal sku As String) As Boolean

        If Not allocMap.Exists("GroupCount") Then Exit Function

        Dim g As Long
        For g = 1 To CLng(allocMap("GroupCount"))
            If Not allocMap.Exists("Group_" & g & "_SKU") Then GoTo NextGroupLink
            If CStr(allocMap("Group_" & g & "_SKU")) <> sku Then GoTo NextGroupLink

            If allocMap.Exists("Group_" & g & "_DetailCount") Then
                Dim d As Long
                For d = 1 To CLng(allocMap("Group_" & g & "_DetailCount"))
                    If CStr(allocMap("Group_" & g & "_WMSOrderNo_" & d)) = wmsOrderNo Then
                        ST_AllocMapLinksWmsSku = True
                        Exit Function
                    End If
                Next d
            End If

            ' 失败组常无明细：该 SKU 组存在且 WMS 在 orders 未知时，由调用方用 orders 判定
    NextGroupLink:
        Next g
    End Function

    Private Function ST_WmsBelongsToShipment( _
        ByVal wmsOrderNo As String, _
        ByVal shipNo As String, _
        ByRef orders() As NormalizedReturnLine) As Boolean

        On Error GoTo NoOrders
        Dim i As Long
        For i = LBound(orders) To UBound(orders)
            If orders(i).ShipmentNo = shipNo And orders(i).WMSOrderNo = wmsOrderNo Then
                ST_WmsBelongsToShipment = True
                Exit Function
            End If
        Next i
    NoOrders:
        ST_WmsBelongsToShipment = False
    End Function


    ' =============================================================================
    ' 五、成功路径：明细提取与退单号聚合
    ' =============================================================================

    Private Function ST_ExtractDetailsFromAllocMap( _
        ByVal allocMap As Object, _
        ByVal shipNo As String, _
        ByRef outDetails() As AllocationDetail) As Long

        If allocMap Is Nothing Then Exit Function
        If Not allocMap.Exists("GroupCount") Then Exit Function

        Dim total As Long
        Dim g As Long
        Dim groupCount As Long
        groupCount = CLng(allocMap("GroupCount"))

        For g = 1 To groupCount
            If allocMap.Exists("Group_" & g & "_DetailCount") Then
                total = total + CLng(allocMap("Group_" & g & "_DetailCount"))
            End If
        Next g

        If total = 0 Then Exit Function

        ReDim outDetails(1 To total)
        Dim idx As Long
        idx = 0

        For g = 1 To groupCount
            If Not allocMap.Exists("Group_" & g & "_DetailCount") Then GoTo NextGroupExtract
            Dim sku As String
            If allocMap.Exists("Group_" & g & "_SKU") Then sku = CStr(allocMap("Group_" & g & "_SKU"))

            Dim d As Long
            Dim detailCount As Long
            detailCount = CLng(allocMap("Group_" & g & "_DetailCount"))
            For d = 1 To detailCount
                idx = idx + 1
                With outDetails(idx)
                    .ShipmentNo = shipNo
                    .SKU = sku
                    .WMSOrderNo = CStr(allocMap("Group_" & g & "_WMSOrderNo_" & d))
                    .LineNo = CStr(allocMap("Group_" & g & "_LineNo_" & d))
                    .OrderQty = CLng(allocMap("Group_" & g & "_OrderQty_" & d))
                    .QC = CStr(allocMap("Group_" & g & "_QC_" & d))
                    .LotNo = CStr(allocMap("Group_" & g & "_LotNo_" & d))
                    .Expiry = CStr(allocMap("Group_" & g & "_Expiry_" & d))
                    .AllocQty = CLng(allocMap("Group_" & g & "_AllocQty_" & d))
                    .LineStatus = CStr(allocMap("Group_" & g & "_LineStatus_" & d))
                End With
            Next d
    NextGroupExtract:
        Next g

        ST_ExtractDetailsFromAllocMap = total
    End Function

    Private Sub ST_BuildWmsStatusMap( _
        ByRef details() As AllocationDetail, _
        ByVal detailCount As Long, _
        ByVal wmsStatus As Object)

        Dim lineStatusByWmsLine As Object
        Set lineStatusByWmsLine = CreateObject("Scripting.Dictionary")

        Dim i As Long
        For i = 1 To detailCount
            Dim lineKey As String
            lineKey = details(i).WMSOrderNo & vbTab & details(i).SKU & vbTab & details(i).LineNo
            lineStatusByWmsLine(lineKey) = details(i).LineStatus
        Next i

        Dim wmsKeys As Object
        Set wmsKeys = CreateObject("Scripting.Dictionary")
        For i = 1 To detailCount
            wmsKeys(details(i).WMSOrderNo) = True
        Next i

        Dim wmsKey As Variant
        For Each wmsKey In wmsKeys.Keys
            If ST_WmsHasManualLine(CStr(wmsKey), lineStatusByWmsLine) Then
                wmsStatus(CStr(wmsKey)) = STATUS_MANUAL
            Else
                wmsStatus(CStr(wmsKey)) = STATUS_BATCH_IMPORT
            End If
        Next wmsKey
    End Sub

    Private Function ST_WmsHasManualLine(ByVal wmsOrderNo As String, ByVal lineStatusByWmsLine As Object) As Boolean
        Dim lk As Variant
        For Each lk In lineStatusByWmsLine.Keys
            Dim parts() As String
            parts = Split(CStr(lk), vbTab)
            If parts(0) = wmsOrderNo Then
                If lineStatusByWmsLine(lk) = STATUS_MANUAL Then
                    ST_WmsHasManualLine = True
                    Exit Function
                End If
            End If
        Next lk
    End Function


    ' =============================================================================
    ' 六、通用小工具
    ' =============================================================================

    Private Function HasValidationIssuesArray(ByRef issues() As ValidationIssue) As Boolean
        On Error GoTo NotAllocated
        HasValidationIssuesArray = (UBound(issues) >= LBound(issues))
        Exit Function
    NotAllocated:
        HasValidationIssuesArray = False
    End Function

    Private Function ST_StringArrayCount(ByRef arr() As String) As Long
        On Error GoTo EmptyArr
        ST_StringArrayCount = UBound(arr) - LBound(arr) + 1
        Exit Function
    EmptyArr:
        ST_StringArrayCount = 0
    End Function

    Private Function ST_SafeObjectArrayCount(ByRef arr() As Object) As Long
        On Error GoTo EmptyArr
        ST_SafeObjectArrayCount = UBound(arr) - LBound(arr) + 1
        Exit Function
    EmptyArr:
        ST_SafeObjectArrayCount = 0
    End Function

    Private Function ST_DictionaryKeysToStringArray(ByVal dict As Object) As String()
        Dim arr() As String
        If dict Is Nothing Then
            ST_DictionaryKeysToStringArray = arr
            Exit Function
        End If

        If dict.Count = 0 Then
            ST_DictionaryKeysToStringArray = arr
            Exit Function
        End If

        ReDim arr(0 To dict.Count - 1)
        Dim i As Long
        Dim k As Variant
        i = 0
        For Each k In dict.Keys
            arr(i) = CStr(k)
            i = i + 1
        Next k
        ST_DictionaryKeysToStringArray = arr
    End Function

    Private Function ST_PickEarlierErrorCode(ByVal currentCode As String, ByVal newCode As String) As String
        If currentCode = vbNullString Then
            ST_PickEarlierErrorCode = newCode
            Exit Function
        End If

        If ST_ErrorCodeSortKey(newCode) < ST_ErrorCodeSortKey(currentCode) Then
            ST_PickEarlierErrorCode = newCode
        Else
            ST_PickEarlierErrorCode = currentCode
        End If
    End Function

    Private Function ST_ErrorCodeSortKey(ByVal errorCode As String) As Long
        If Len(errorCode) >= 2 And Left(errorCode, 1) = "E" Then
            On Error Resume Next
            ST_ErrorCodeSortKey = CLng(Mid(errorCode, 2))
            On Error GoTo 0
        Else
            ST_ErrorCodeSortKey = 9999
        End If
    End Function

    Private Function ST_SortUniqueCodes(ByRef codes() As String) As String()
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        Dim i As Long
        On Error GoTo EmptyInput
        For i = LBound(codes) To UBound(codes)
            If Len(codes(i)) > 0 Then dict(codes(i)) = ST_ErrorCodeSortKey(codes(i))
        Next i

        Dim count As Long
        count = dict.Count
        If count = 0 Then
            Dim emptyArr() As String
            ST_SortUniqueCodes = emptyArr
            Exit Function
        End If

        Dim keys() As String
        ReDim keys(0 To count - 1)
        Dim sortKeys() As Long
        ReDim sortKeys(0 To count - 1)

        i = 0
        Dim k As Variant
        For Each k In dict.Keys
            keys(i) = CStr(k)
            sortKeys(i) = CLng(dict(k))
            i = i + 1
        Next k

        Dim a As Long
        Dim b As Long
        For a = 0 To count - 2
            For b = a + 1 To count - 1
                If sortKeys(b) < sortKeys(a) Then
                    Dim tmpKey As Long
                    tmpKey = sortKeys(a)
                    sortKeys(a) = sortKeys(b)
                    sortKeys(b) = tmpKey
                    Dim tmpCode As String
                    tmpCode = keys(a)
                    keys(a) = keys(b)
                    keys(b) = tmpCode
                End If
            Next b
        Next a

        ST_SortUniqueCodes = keys
        Exit Function
    EmptyInput:
        Dim blank() As String
        ST_SortUniqueCodes = blank
    End Function
