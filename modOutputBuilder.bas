Option Explicit

' =============================================================================
' M13_输出构建（modOutputBuilder）
' =============================================================================
' 职责：把领域层结果整理成“可直接写表”的输出行数组，但本模块不执行任何 Excel 写入。
'
' 给新手的解释：
' 可以把本模块理解成“翻译层”：
' - 上游模块给的是结构体、字典、事件等程序对象；
' - 本模块把它们翻译成每一行可写入表格的值数组。
' =============================================================================


' =============================================================================
' 一、公开函数（与规格 M13 对齐）
' =============================================================================

' 构建“分配状态汇总表”行数组。
' dryRunMode=True 时，只输出失败项（通过校验的物流单号不进汇总），用于干跑场景。
Public Function BuildSummaryRows( _
    ByRef wmsStatusMap() As WMSStatusEntry, _
    ByVal dryRunMode As Boolean) As OutputRow()

    Dim total As Long
    total = OB_WmsStatusCount(wmsStatusMap)
    If total = 0 Then Exit Function

    Dim rows() As OutputRow
    ReDim rows(1 To total)

    Dim i As Long
    Dim outIndex As Long
    For i = LBound(wmsStatusMap) To UBound(wmsStatusMap)
        ' 干跑只保留失败项，避免把“仅通过校验但未分配”的单据误导成成功输出。
        If dryRunMode Then
            If wmsStatusMap(i).Status <> STATUS_UNALLOCATED Then
                GoTo NextEntry
            End If
        End If

        outIndex = outIndex + 1
        rows(outIndex) = OB_CreateRow( _
            wmsStatusMap(i).ShipmentNo, _
            OB_NormalizeSummaryWmsNo(wmsStatusMap(i).WMSOrderNo), _
            wmsStatusMap(i).Status, _
            wmsStatusMap(i).Reason)
NextEntry:
    Next i

    If outIndex = 0 Then
        Exit Function
    End If

    ReDim Preserve rows(1 To outIndex)
    BuildSummaryRows = rows
End Function

' 构建“成功分配明细表”行数组。
' finalResult 采用当前项目约定的 Dictionary 键（DetailCount、Detail_i_*）。
Public Function BuildDetailRows(ByVal finalResult As Object) As OutputRow()
    If finalResult Is Nothing Then Exit Function
    If Not finalResult.Exists("DetailCount") Then Exit Function

    Dim detailCount As Long
    detailCount = CLng(finalResult("DetailCount"))
    If detailCount <= 0 Then Exit Function

    Dim rows() As OutputRow
    ReDim rows(1 To detailCount)

    Dim i As Long
    For i = 1 To detailCount
        rows(i) = OB_CreateRow( _
            OB_GetText(finalResult, "Detail_" & i & "_ShipmentNo"), _
            OB_GetText(finalResult, "Detail_" & i & "_WMSOrderNo"), _
            OB_GetText(finalResult, "Detail_" & i & "_SKU"), _
            OB_GetText(finalResult, "Detail_" & i & "_LineNo"), _
            OB_GetLong(finalResult, "Detail_" & i & "_OrderQty"), _
            OB_GetText(finalResult, "Detail_" & i & "_QC"), _
            OB_GetText(finalResult, "Detail_" & i & "_LotNo"), _
            OB_GetText(finalResult, "Detail_" & i & "_Expiry"), _
            OB_GetLong(finalResult, "Detail_" & i & "_AllocQty"), _
            OB_GetText(finalResult, "Detail_" & i & "_LineStatus"), _
            OB_GetText(finalResult, "Detail_" & i & "_WMSOrderStatus"))
    Next i

    BuildDetailRows = rows
End Function

' 构建“数据异常明细表”行数组。
' 规则：E06/E08/E11 只进汇总，不进入异常明细。
Public Function BuildAnomalyOutputRows(ByRef anomalyRows() As AnomalyRow) As OutputRow()
    Dim total As Long
    total = OB_AnomalyCount(anomalyRows)
    If total = 0 Then Exit Function

    Dim rows() As OutputRow
    ReDim rows(1 To total)

    Dim i As Long
    Dim outIndex As Long
    For i = LBound(anomalyRows) To UBound(anomalyRows)
        If OB_IsSummaryOnlyError(anomalyRows(i).ErrorCode) Then
            GoTo NextAnomaly
        End If

        outIndex = outIndex + 1
        rows(outIndex) = OB_CreateRow( _
            anomalyRows(i).SourceTable, _
            anomalyRows(i).ExcelRowNum, _
            anomalyRows(i).ShipmentNo, _
            anomalyRows(i).WMSOrderNo, _
            anomalyRows(i).SKU, _
            anomalyRows(i).FieldName, _
            anomalyRows(i).RawValue, _
            anomalyRows(i).ErrorCode, _
            anomalyRows(i).Reason)
NextAnomaly:
    Next i

    If outIndex = 0 Then Exit Function

    ReDim Preserve rows(1 To outIndex)
    BuildAnomalyOutputRows = rows
End Function

' 构建“调试日志表”行数组（19 列，见 调试日志19列规格说明.md）。
' 关闭=不写数据行；简版=仅 IsFinalResult=True；详细=全部事件。
Public Function BuildDebugLogRows( _
    ByRef events() As AllocationEvent, _
    cfg As ConfigStruct) As OutputRow()

    If cfg.DebugLogLevel = DEBUG_LEVEL_OFF Then Exit Function

    Dim total As Long
    total = OB_EventCount(events)
    If total = 0 Then Exit Function

    Dim rows() As OutputRow
    ReDim rows(1 To total)

    Dim i As Long
    Dim outIndex As Long
    For i = LBound(events) To UBound(events)
        If OB_ShouldIncludeDebugEvent(events(i), cfg) Then
            outIndex = outIndex + 1
            rows(outIndex) = OB_MapDebugEventToRow(events(i))
        End If
    Next i

    If outIndex = 0 Then Exit Function

    ReDim Preserve rows(1 To outIndex)
    BuildDebugLogRows = rows
End Function

' 构建“运行历史记录表”单行（20 列：需求 §5.6 的 17 字段 + 3 个配置快照字段）。
' 运行历史每次只追加一行，因此这里返回单个 OutputRow。
' 第 1 列“运行编号”由 M14 AppendRunHistory 按表内已有行数自动生成，此处占位空串。
' 新增参数（计时/错误码分布）带默认值，便于既有单元测试按旧三参调用。
Public Function BuildRunHistoryRow( _
    stats As RunStats, _
    cfg As ConfigStruct, _
    ByVal dryRunMode As Boolean, _
    Optional ByVal runTimeText As String = "", _
    Optional ByVal validateSecs As Single = 0, _
    Optional ByVal allocSecs As Single = 0, _
    Optional ByVal totalSecs As Single = 0, _
    Optional ByVal errorCodeDist As String = "") As OutputRow

    BuildRunHistoryRow = OB_CreateRow( _
        "", _
        runTimeText, _
        IIf(dryRunMode, "Dry Run", "Full Run"), _
        stats.InputReturnRows, _
        stats.InputInventoryRows, _
        stats.InputShipmentCount, _
        validateSecs, _
        allocSecs, _
        totalSecs, _
        stats.ValidationFailCount, _
        stats.AllocSuccessCount, _
        stats.AllocFailCount, _
        errorCodeDist, _
        stats.TotalBacktrackCount, _
        stats.MaxGroupBacktrack, _
        cfg.DebugLogLevel, _
        "", _
        cfg.MaxBacktrackCount, _
        IIf(cfg.LotCaseSensitive, LOT_MODE_SENSITIVE, LOT_MODE_INSENSITIVE), _
        cfg.NoExpirySentinel)
End Function


' =============================================================================
' 二、私有工具函数
' =============================================================================

Private Function OB_CreateRow(ParamArray values() As Variant) As OutputRow
    Dim row As OutputRow
    Dim cols() As Variant
    Dim i As Long

    ReDim cols(1 To UBound(values) - LBound(values) + 1)
    For i = LBound(values) To UBound(values)
        cols(i - LBound(values) + 1) = values(i)
    Next i

    row.Values = cols
    OB_CreateRow = row
End Function

Private Function OB_NormalizeSummaryWmsNo(ByVal wmsOrderNo As String) As String
    If Len(wmsOrderNo) > 0 Then
        OB_NormalizeSummaryWmsNo = wmsOrderNo
        Exit Function
    End If

    ' E07 明确要求填 [N/A]；其余空值也统一回填占位符，避免输出空白。
    OB_NormalizeSummaryWmsNo = NA_PLACEHOLDER
End Function

Private Function OB_IsSummaryOnlyError(ByVal errorCode As String) As Boolean
    Select Case errorCode
        Case ERR_E06, ERR_E08, ERR_E11
            OB_IsSummaryOnlyError = True
        Case Else
            OB_IsSummaryOnlyError = False
    End Select
End Function

Private Function OB_BoolText(ByVal value As Boolean) As String
    If value Then
        OB_BoolText = "True"
    Else
        OB_BoolText = "False"
    End If
End Function

Private Function OB_GetText(ByVal dict As Object, ByVal keyName As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(keyName) Then OB_GetText = CStr(dict(keyName))
End Function

Private Function OB_GetLong(ByVal dict As Object, ByVal keyName As String) As Long
    If dict Is Nothing Then Exit Function
    If dict.Exists(keyName) Then OB_GetLong = CLng(dict(keyName))
End Function

Private Function OB_WmsStatusCount(ByRef entries() As WMSStatusEntry) As Long
    On Error GoTo EmptyArr
    OB_WmsStatusCount = UBound(entries) - LBound(entries) + 1
    Exit Function
EmptyArr:
    OB_WmsStatusCount = 0
End Function

Private Function OB_AnomalyCount(ByRef rows() As AnomalyRow) As Long
    On Error GoTo EmptyArr
    OB_AnomalyCount = UBound(rows) - LBound(rows) + 1
    Exit Function
EmptyArr:
    OB_AnomalyCount = 0
End Function

Private Function OB_EventCount(ByRef rows() As AllocationEvent) As Long
    On Error GoTo EmptyArr
    OB_EventCount = UBound(rows) - LBound(rows) + 1
    Exit Function
EmptyArr:
    OB_EventCount = 0
End Function

' 按日志级别决定是否输出该事件（简版只保留最终结果行）。
Private Function OB_ShouldIncludeDebugEvent( _
    ByRef evt As AllocationEvent, _
    ByRef cfg As ConfigStruct) As Boolean

    Select Case cfg.DebugLogLevel
        Case DEBUG_LEVEL_SIMPLE
            OB_ShouldIncludeDebugEvent = evt.IsFinalResult
        Case DEBUG_LEVEL_DETAIL
            OB_ShouldIncludeDebugEvent = True
        Case Else
            OB_ShouldIncludeDebugEvent = False
    End Select
End Function

' 将 AllocationEvent 映射为 19 列 OutputRow。
Private Function OB_MapDebugEventToRow(ByRef evt As AllocationEvent) As OutputRow
    OB_MapDebugEventToRow = OB_CreateRow( _
        evt.ShipmentNo, _
        evt.SKU, _
        evt.WMSOrderNo, _
        evt.LineNo, _
        evt.DemandD, _
        evt.ProcessOrder, _
        evt.DynamicNextMinQty, _
        evt.CandidateQCCount, _
        evt.ExcludedQCList, _
        evt.StrategyUsed, _
        evt.UsedQC, _
        evt.QCBefore, _
        evt.QCAfter, _
        evt.LotExpiryComboCount, _
        evt.IsBacktrackRetry, _
        evt.BacktrackNo, _
        evt.LineStatus, _
        evt.ErrorCode, _
        evt.FailSubType)
End Function
