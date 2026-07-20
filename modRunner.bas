Option Explicit

' =============================================================================
' M15_运行编排（modRunner）
' =============================================================================
' 职责：系统的两个按钮入口（干跑 / 完整运行），串联所有模块的调用顺序；
'       以及 BuildRunStats，统一构造本次运行的汇总统计。
'
' 给新手的解释：
'   本模块相当于整个系统的"总调度员"：
'     - 干跑（RunValidationOnly）：只检查数据，不分配货物，快速发现输入问题。
'     - 完整运行（RunFullAllocation）：在数据通过检查后，执行实际的回溯分配算法。
'   两个按钮的内部流程高度相似，区别在于干跑跳过 M06-M09 的分配核心。
'
' 公开函数（与规格 modRunner 一致）：
'   RunValidationOnly(wb)      — 干跑按钮
'   RunFullAllocation(wb)      — 完整运行按钮
'   BuildRunStats(...)         — 统一构造 RunStats，干跑/完整均可调用
'
' 注意（偏离规格的合理调整）：
'   BuildRunStats 的 shipmentResults 参数类型改为 Object() 而非 ShipmentAllocResult()。
'   原因：M09 AllocateShipment 和 M11 ApplyRollback 已统一使用 Scripting.Dictionary
'   传递复杂结构（VBA Type 无法直接存入 Dictionary），这里保持一致。
'   干跑时传入未初始化的空 Object 数组即可，分配相关字段均为 0。
' =============================================================================

' 输入表和配置表名称（只有 M15 需要直接读取，输出表由 M14 modExcelOutput 内部管理）
Private Const SHEET_RETURN_INPUT    As String = "输入_退单表"
Private Const SHEET_INVENTORY_INPUT As String = "输入_质检库存表"
Private Const SHEET_CONFIG          As String = "输入_配置"
Private Const SHEET_RUN_HISTORY     As String = "运行历史记录表"


' =============================================================================
' 一、公开函数：两个按钮入口
' =============================================================================

' 以下三个无参数过程专门供生产工作簿按钮绑定。
' Excel 按钮不能直接调用带 Workbook 参数的过程，因此由这里统一传入 ThisWorkbook。
Public Sub StartValidationOnly()
    RunValidationOnly ThisWorkbook
End Sub


Public Sub StartFullAllocation()
    RunFullAllocation ThisWorkbook
End Sub


Public Sub ClearAllocationResults()
    On Error GoTo ClearFail

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    ClearOutputSheets ThisWorkbook, cfg

    MsgBox "输出结果已清空，输入数据、配置和运行历史均已保留。", vbInformation
    Exit Sub

ClearFail:
    MsgBox "清空输出结果失败（" & Err.Description & "）。" & vbNewLine & _
           "请检查输出工作表是否受保护。", vbCritical
End Sub


' 干跑模式入口：只做校验，不执行分配。
' 流程：清空输出 → 读取原始数据 → 标准化 → 分配前校验 →
'       BuildRunStats → 构建输出行 → 写入 Excel → 提示完成
' showMessages=False 时供批量测试运行器调用，避免每个子批次弹窗打断流程。
Public Sub RunValidationOnly(wb As Workbook, Optional ByVal showMessages As Boolean = True)
    If wb Is Nothing Then
        If showMessages Then MsgBox "工作簿对象为空，无法运行。", vbCritical
        Exit Sub
    End If

    ' 运行开始时间戳与计时起点（需求 §5.6：运行时间/校验耗时/总耗时）
    Dim runStartText As String
    runStartText = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Dim t0 As Single
    t0 = Timer

    ' 读取配置（M02）；配置非法时：交互模式给友好提示，静默模式原样抛错供批量断言
    Dim cfg As ConfigStruct
    On Error GoTo ConfigFail
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
    On Error GoTo 0

    ' 清空输出表；受保护时捕获错误并中止，避免覆盖已保护内容
    On Error GoTo ClearFail
    ClearOutputSheets wb, cfg
    On Error GoTo RunFail

    ' 读取原始数据（M03）
    Dim rawOrders() As RawReturnRow
    Dim rawInventory() As RawInventoryRow
    rawOrders = ReadReturnOrders(wb.Worksheets(SHEET_RETURN_INPUT))
    rawInventory = ReadQCInventory(wb.Worksheets(SHEET_INVENTORY_INPUT))

    ' 标准化（M04）；两次调用的 FieldNormalizeIssue 分别输出再合并，
    ' 避免第二次调用覆盖第一次的结果
    Dim returnIssues() As FieldNormalizeIssue
    Dim inventoryIssues() As FieldNormalizeIssue
    Dim orders() As NormalizedReturnLine
    Dim inventory() As NormalizedInventoryLine
    orders = NormalizeReturnRows(rawOrders, cfg, returnIssues)
    inventory = NormalizeInventoryRows(rawInventory, cfg, inventoryIssues)
    Dim normalizedIssues() As FieldNormalizeIssue
    normalizedIssues = RN_MergeFieldIssues(returnIssues, inventoryIssues)

    ' 分配前校验（M05）
    Dim validationIssues() As ValidationIssue
    Dim validationResult As ValidationResult
    validationResult = ValidatePre(orders, inventory, normalizedIssues, cfg, validationIssues)

    ' BuildRunStats（干跑：传空数组，分配相关字段全为 0）
    Dim emptyResults() As Object
    Dim stats As RunStats
    stats = BuildRunStats(validationResult, emptyResults, orders, inventory)

    ' 构建 FinalResult：干跑只有校验失败项，无成功分配明细
    Dim emptyAllocResults() As Object
    Dim finalResult As Object
    Set finalResult = ApplyRollback(emptyAllocResults, validationResult, validationIssues, orders)

    ' 构建各输出行（M13）
    Dim anomalyRows() As AnomalyRow
    anomalyRows = BuildAnomalyRows(validationIssues)

    Dim statusMap() As WMSStatusEntry
    statusMap = AggregateWMSStatus(finalResult)

    Dim emptyEvents() As AllocationEvent
    Dim runHistoryRow As OutputRow
    runHistoryRow = BuildRunHistoryRow(stats, cfg, True, runStartText, _
        RN_ElapsedSecs(t0, Timer), 0, RN_ElapsedSecs(t0, Timer), _
        RN_BuildErrorCodeDistribution(validationIssues, emptyResults))

    ' 写入 Excel（M14）
    RN_WriteAllOutput wb, cfg, statusMap, finalResult, anomalyRows, emptyEvents, runHistoryRow, True

    On Error GoTo 0
    If showMessages Then
        MsgBox "干跑完成。" & vbNewLine & _
               "校验失败物流单号：" & stats.ValidationFailCount & " 个", vbInformation
    End If
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

RunFail:
    If showMessages Then
        MsgBox "干跑过程中出现错误（" & Err.Number & "）：" & Err.Description, vbCritical
    Else
        Err.Raise Err.Number, "RunValidationOnly", Err.Description
    End If
End Sub

Public Sub RunValidationOnlySilent(wb As Workbook)
    RunValidationOnly wb, False
End Sub


' 完整运行模式入口：在通过校验的物流单号上执行回溯分配算法。
' 流程：清空 → 读取 → 标准化 → 前校验 → 建账本 → 排序预检 →
'       回溯分配 → 整单状态 → 后校验 → BuildRunStats → 输出 → 写入 → 提示
' E99 守卫错误在最外层统一捕获，写入日志后终止，不继续执行。
Public Sub RunFullAllocation(wb As Workbook, Optional ByVal showMessages As Boolean = True)
    If wb Is Nothing Then
        If showMessages Then MsgBox "工作簿对象为空，无法运行。", vbCritical
        Exit Sub
    End If

    ' 运行开始时间戳与计时起点（需求 §5.6：运行时间/校验耗时/分配耗时/总耗时）
    Dim runStartText As String
    runStartText = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Dim t0 As Single
    t0 = Timer

    Dim cfg As ConfigStruct
    On Error GoTo ConfigFail
    cfg = LoadConfig(wb.Worksheets(SHEET_CONFIG))
    On Error GoTo 0

    On Error GoTo ClearFail
    ClearOutputSheets wb, cfg
    On Error GoTo RunFail

    ' 读取（M03）
    Dim rawOrders() As RawReturnRow
    Dim rawInventory() As RawInventoryRow
    rawOrders = ReadReturnOrders(wb.Worksheets(SHEET_RETURN_INPUT))
    rawInventory = ReadQCInventory(wb.Worksheets(SHEET_INVENTORY_INPUT))

    ' 标准化（M04）
    Dim returnIssues() As FieldNormalizeIssue
    Dim inventoryIssues() As FieldNormalizeIssue
    Dim orders() As NormalizedReturnLine
    Dim inventory() As NormalizedInventoryLine
    orders = NormalizeReturnRows(rawOrders, cfg, returnIssues)
    inventory = NormalizeInventoryRows(rawInventory, cfg, inventoryIssues)
    Dim normalizedIssues() As FieldNormalizeIssue
    normalizedIssues = RN_MergeFieldIssues(returnIssues, inventoryIssues)

    ' 前校验（M05）
    Dim validationIssues() As ValidationIssue
    Dim validationResult As ValidationResult
    validationResult = ValidatePre(orders, inventory, normalizedIssues, cfg, validationIssues)
    Dim tValidate As Single
    tValidate = Timer

    ' 建账本（M06）
    Dim ledger As Object
    Set ledger = BuildLedger(inventory)

    ' E99 守卫断言会以 vbObjectError + 99 号抛出，需要在分配循环外捕获
    On Error GoTo E99Fail

    ' 回溯分配（M07 + M09）：仅对通过校验的物流单号执行
    Dim shipmentResults() As Object
    shipmentResults = RN_RunAllAllocations(orders, ledger, cfg, validationIssues)
    Dim tAlloc As Single
    tAlloc = Timer

    On Error GoTo RunFail

    ' 整单状态判定（M11）
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipmentResults, validationResult, validationIssues, orders)

    ' 分配后校验（M12）：这是写入成功明细前的最后一道防线。
    ' 只要 M12 发现内部一致性问题，就升级为 E99，避免把错误结果交给业务侧。
    Dim postResult As Object
    Set postResult = ValidatePost(orders, finalResult)
    RN_RaisePostValidationE99IfNeeded postResult

    ' 构造运行统计（M15 BuildRunStats）
    Dim stats As RunStats
    stats = BuildRunStats(validationResult, shipmentResults, orders, inventory)

    ' 构建各输出行（M13）
    Dim anomalyRows() As AnomalyRow
    anomalyRows = BuildAnomalyRows(validationIssues)

    Dim statusMap() As WMSStatusEntry
    statusMap = AggregateWMSStatus(finalResult)

    ' 从 M09 分配结果 Dictionary 中提取调试日志事件（19 列源数据）
    Dim events() As AllocationEvent
    events = RN_CollectDebugEvents(shipmentResults)

    Dim runHistoryRow As OutputRow
    runHistoryRow = BuildRunHistoryRow(stats, cfg, False, runStartText, _
        RN_ElapsedSecs(t0, tValidate), RN_ElapsedSecs(tValidate, tAlloc), RN_ElapsedSecs(t0, Timer), _
        RN_BuildErrorCodeDistribution(validationIssues, shipmentResults))

    ' 写入（M14）
    RN_WriteAllOutput wb, cfg, statusMap, finalResult, anomalyRows, events, runHistoryRow, False

    On Error GoTo 0
    If showMessages Then
        MsgBox "完整分配完成。" & vbNewLine & _
               "成功：" & stats.AllocSuccessCount & " 个物流单号" & vbNewLine & _
               "失败：" & stats.AllocFailCount & " 个物流单号", vbInformation
    End If
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

E99Fail:
    ' E99 是库存守恒等式被破坏，属于严重的数据一致性错误，立即中止。
    If showMessages Then
        If Err.Number = E99_ERROR_NUMBER Then
            MsgBox "E99 工程守卫触发：" & Err.Description & vbNewLine & _
                   "已中止运行，请检查库存数据一致性后重试。", vbCritical
        Else
            MsgBox "运行时错误（" & Err.Number & "）：" & Err.Description, vbCritical
        End If
    Else
        Err.Raise Err.Number, "RunFullAllocation", Err.Description
    End If
    Exit Sub

RunFail:
    If showMessages Then
        MsgBox "运行时错误（" & Err.Number & "）：" & Err.Description, vbCritical
    Else
        Err.Raise Err.Number, "RunFullAllocation", Err.Description
    End If
End Sub

Public Sub RunFullAllocationSilent(wb As Workbook)
    RunFullAllocation wb, False
End Sub


' =============================================================================
' 二、公开函数：BuildRunStats 统一构造运行统计
' =============================================================================

' 统一构造本次运行的汇总统计，供 M13 BuildRunHistoryRow 写入运行历史记录表。
'
' 干跑时：shipmentResults 传入未初始化的空数组，
'   TotalBacktrackCount / MaxGroupBacktrack / AllocSuccessCount / AllocFailCount 均为 0。
' 完整运行时：shipmentResults 为 M09 AllocateShipment 返回的 Object 数组，
'   函数遍历每个物流单号的 SKU 组统计回溯次数和整单成功/失败。
'
' 参数：
'   validationResult  - M05 ValidatePre 的输出（含 FailedShipmentCount）
'   shipmentResults() - M09 AllocateShipment 返回的 Object 数组（干跑时传空数组）
'   orders()          - 标准化后的退单行数组（用于计算 InputReturnRows）
'   inventory()       - 标准化后的库存行数组（用于计算 InputInventoryRows）
Public Function BuildRunStats( _
    ByRef validationResult As ValidationResult, _
    ByRef shipmentResults() As Object, _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine) As RunStats

    Dim stats As RunStats

    ' 输入行数直接取数组长度（空数组时安全返回 0）
    stats.InputReturnRows = RN_ReturnLineCount(orders)
    stats.InputInventoryRows = RN_InventoryLineCount(inventory)
    stats.InputShipmentCount = RN_CountDistinctShipments(orders, inventory)

    ' 校验失败数来自 M05 ValidatePre 的输出
    stats.ValidationFailCount = validationResult.FailedShipmentCount

    ' 遍历每个物流单号的分配结果，累计回溯次数和成功/失败计数
    ' 干跑时 shipmentResults 为空数组，循环不执行，相关字段保持 0
    Dim resultCount As Long
    resultCount = RN_ObjectArrayCount(shipmentResults)

    Dim i As Long
    For i = 1 To resultCount
        Dim result As Object
        Set result = shipmentResults(LBound(shipmentResults) + i - 1)

        ' 跳过无效条目（防御性处理）
        If result Is Nothing Then GoTo NextResult
        If Not result.Exists("GroupCount") Then GoTo NextResult

        Dim groupCount As Long
        groupCount = CLng(result("GroupCount"))
        If groupCount = 0 Then GoTo NextResult

        ' 遍历该物流单号下每个 SKU 组，累计回溯次数
        Dim g As Long
        Dim shipAllSuccess As Boolean
        shipAllSuccess = True   ' 假设整单成功，遇到失败组则改 False

        For g = 1 To groupCount
            ' 回溯次数累计
            Dim bt As Long
            bt = 0
            If result.Exists("Group_" & g & "_BacktrackCount") Then
                bt = CLng(result("Group_" & g & "_BacktrackCount"))
            End If
            stats.TotalBacktrackCount = stats.TotalBacktrackCount + bt
            If bt > stats.MaxGroupBacktrack Then stats.MaxGroupBacktrack = bt

            ' 判断该组是否成功
            If result.Exists("Group_" & g & "_Success") Then
                If Not CBool(result("Group_" & g & "_Success")) Then
                    shipAllSuccess = False
                End If
            End If
        Next g

        ' 整单成功/失败以物流单号为粒度：所有 SKU 组均成功才算整单成功
        If shipAllSuccess Then
            stats.AllocSuccessCount = stats.AllocSuccessCount + 1
        Else
            stats.AllocFailCount = stats.AllocFailCount + 1
        End If

NextResult:
    Next i

    BuildRunStats = stats
End Function


' =============================================================================
' 三、私有工具函数
' =============================================================================

' 对所有通过校验的物流单号执行回溯分配，返回 ShipmentResult Object 数组。
' 内部按"物流单号 → SKU"两层分组，依次调用 M07 BuildStaticPlan / RunPrecheck
' 和 M09 AllocateShipment，符合规格 §4.2.7 描述的分配编排逻辑。
Private Function RN_RunAllAllocations( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal ledger As Object, _
    ByRef cfg As ConfigStruct, _
    ByRef validationIssues() As ValidationIssue) As Object()

    ' 收集所有物流单号（按出现顺序去重）
    Dim shipNos As Object
    Set shipNos = CreateObject("Scripting.Dictionary")
    shipNos.CompareMode = vbTextCompare

    Dim orderCount As Long
    orderCount = RN_ReturnLineCount(orders)
    Dim i As Long
    ' orderCount=0 时 orders 可能是未初始化数组，直接求 LBound 会触发错误 9；
    ' 因此只在有行时才进入循环（空输入属合法场景，由后续空结果链路处理）。
    If orderCount > 0 Then
        For i = LBound(orders) To LBound(orders) + orderCount - 1
            Dim sno As String
            sno = orders(i).ShipmentNo
            If Not shipNos.Exists(sno) Then shipNos.Add sno, True
        Next i
    End If

    ' 移除在校验阶段已失败的物流单号，只对通过校验的进行分配
    RN_RemoveFailedShipments shipNos, validationIssues

    Dim passCount As Long
    passCount = shipNos.Count
    If passCount = 0 Then
        Dim emptyArr() As Object
        RN_RunAllAllocations = emptyArr
        Exit Function
    End If

    Dim results() As Object
    ReDim results(1 To passCount)

    Dim resultIdx As Long
    Dim shipKey As Variant
    For Each shipKey In shipNos.Keys
        Dim shipNo As String
        shipNo = CStr(shipKey)

        ' 为该物流单号构建 SKU 分组计划和预检测结论
        Dim skuList() As String
        Dim planMap As Object
        Dim precheckMap As Object
        Set planMap = CreateObject("Scripting.Dictionary")
        Set precheckMap = CreateObject("Scripting.Dictionary")
        RN_BuildSkuGroupsForShipment orders, shipNo, ledger, skuList, planMap, precheckMap

        ' 调用 M09 执行带回溯的分配（AllocateShipment 内部处理短路和连带回滚）
        resultIdx = resultIdx + 1
        Set results(resultIdx) = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)
    Next shipKey

    RN_RunAllAllocations = results
End Function


' 为指定物流单号构建 skuList（SKU 列表）、planMap（静态计划）和 precheckMap（预检测结论）。
' 每个 SKU 对应一个 BuildStaticPlan 返回的 Object，并附带 RunPrecheck 的预检测结论。
Private Sub RN_BuildSkuGroupsForShipment( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipNo As String, _
    ByVal ledger As Object, _
    ByRef skuList() As String, _
    ByRef planMap As Object, _
    ByRef precheckMap As Object)

    ' 收集该物流单号下的 SKU 列表（去重，保持原顺序）
    Dim skuOrder As Object
    Set skuOrder = CreateObject("Scripting.Dictionary")
    skuOrder.CompareMode = vbTextCompare

    Dim orderCount As Long
    orderCount = RN_ReturnLineCount(orders)
    Dim i As Long
    For i = LBound(orders) To LBound(orders) + orderCount - 1
        If orders(i).ShipmentNo = shipNo Then
            Dim sku As String
            sku = orders(i).SKU
            If Not skuOrder.Exists(sku) Then skuOrder.Add sku, True
        End If
    Next i

    Dim skuCount As Long
    skuCount = skuOrder.Count
    If skuCount = 0 Then Exit Sub

    ' 填充 skuList 数组
    ReDim skuList(1 To skuCount)
    Dim idx As Long
    Dim skuKey As Variant
    For Each skuKey In skuOrder.Keys
        idx = idx + 1
        skuList(idx) = CStr(skuKey)
    Next skuKey

    ' 为每个 SKU 调用 BuildStaticPlan（M07）和 RunPrecheck（M07）
    For idx = 1 To skuCount
        Dim curSku As String
        curSku = skuList(idx)

        ' 筛选该物流单号 + SKU 对应的退单行，作为 BuildStaticPlan 的输入
        Dim groupRows() As NormalizedReturnLine
        groupRows = RN_FilterOrdersByShipmentSKU(orders, shipNo, curSku)

        Dim plan As Object
        Set plan = BuildStaticPlan(groupRows, ledger)
        planMap.Add curSku, plan

        ' RunPrecheck 返回 PrecheckResult，存入 precheckMap 时用 Variant Array，
        ' 与 modBacktracking 约定一致（VBA Type 无法直接存入 Dictionary）
        Dim precheck As PrecheckResult
        precheck = RunPrecheck(plan, ledger)
        precheckMap.Add curSku, Array(precheck.PrecheckAHit, precheck.PrecheckBHit)
    Next idx
End Sub


' 从 orders 中筛选匹配指定物流单号和 SKU 的行，返回子数组。
' 采用两遍扫描（先计数后填充），避免使用动态增长数组带来的性能开销。
Private Function RN_FilterOrdersByShipmentSKU( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipNo As String, _
    ByVal sku As String) As NormalizedReturnLine()

    Dim orderCount As Long
    orderCount = RN_ReturnLineCount(orders)
    If orderCount = 0 Then Exit Function

    ' 第一遍：统计匹配行数
    Dim matchCount As Long
    Dim i As Long
    For i = LBound(orders) To LBound(orders) + orderCount - 1
        If orders(i).ShipmentNo = shipNo And orders(i).SKU = sku Then
            matchCount = matchCount + 1
        End If
    Next i
    If matchCount = 0 Then Exit Function

    ' 第二遍：填充结果数组
    Dim result() As NormalizedReturnLine
    ReDim result(1 To matchCount)
    Dim resultIdx As Long
    For i = LBound(orders) To LBound(orders) + orderCount - 1
        If orders(i).ShipmentNo = shipNo And orders(i).SKU = sku Then
            resultIdx = resultIdx + 1
            result(resultIdx) = orders(i)
        End If
    Next i

    RN_FilterOrdersByShipmentSKU = result
End Function


' 从 shipNos 字典中移除在校验阶段已标记失败的物流单号。
' 逻辑：validationIssues 中出现过的物流单号均视为校验失败（与 M11 ApplyRollback 逻辑一致）。
Private Sub RN_RemoveFailedShipments( _
    ByRef shipNos As Object, _
    ByRef validationIssues() As ValidationIssue)

    Dim issueCount As Long
    On Error Resume Next
    issueCount = UBound(validationIssues) - LBound(validationIssues) + 1
    On Error GoTo 0
    If issueCount <= 0 Then Exit Sub

    Dim i As Long
    For i = LBound(validationIssues) To UBound(validationIssues)
        Dim sno As String
        sno = validationIssues(i).ShipmentNo
        If Len(sno) > 0 And shipNos.Exists(sno) Then
            shipNos.Remove sno
        End If
    Next i
End Sub


' 合并两组 FieldNormalizeIssue 数组（退单表 + 库存表各自的标准化问题）。
' 两次 Normalize 调用的 outIssues 变量是独立的，需在此合并后再传给 ValidatePre。
Private Function RN_MergeFieldIssues( _
    ByRef issuesA() As FieldNormalizeIssue, _
    ByRef issuesB() As FieldNormalizeIssue) As FieldNormalizeIssue()

    Dim countA As Long
    Dim countB As Long
    On Error Resume Next
    countA = UBound(issuesA) - LBound(issuesA) + 1
    countB = UBound(issuesB) - LBound(issuesB) + 1
    On Error GoTo 0

    If countA <= 0 And countB <= 0 Then Exit Function

    Dim merged() As FieldNormalizeIssue
    ReDim merged(1 To countA + countB)

    Dim idx As Long
    Dim i As Long

    If countA > 0 Then
        For i = LBound(issuesA) To UBound(issuesA)
            idx = idx + 1
            merged(idx) = issuesA(i)
        Next i
    End If

    If countB > 0 Then
        For i = LBound(issuesB) To UBound(issuesB)
            idx = idx + 1
            merged(idx) = issuesB(i)
        Next i
    End If

    RN_MergeFieldIssues = merged
End Function


' 统一写入所有输出工作表（M14 WriteSheet / WriteDebugLog / AppendRunHistory）。
' 汇总表、明细表、异常明细表的列顺序和表头在此集中定义，便于后续维护。
Private Sub RN_WriteAllOutput( _
    ByVal wb As Workbook, _
    ByRef cfg As ConfigStruct, _
    ByRef statusMap() As WMSStatusEntry, _
    ByVal finalResult As Object, _
    ByRef anomalyRows() As AnomalyRow, _
    ByRef events() As AllocationEvent, _
    ByRef runHistoryRow As OutputRow, _
    ByVal dryRunMode As Boolean)

    ' 分配状态汇总表
    Dim summaryRows() As OutputRow
    summaryRows = BuildSummaryRows(statusMap, dryRunMode)
    Dim summaryHeaders(1 To 4) As String
    summaryHeaders(1) = "物流单号"
    summaryHeaders(2) = "WMS退单号"
    summaryHeaders(3) = "退单号状态"
    summaryHeaders(4) = "原因"
    WriteSheet wb.Worksheets("分配状态汇总表"), summaryRows, summaryHeaders

    ' 成功分配明细表
    Dim detailRows() As OutputRow
    detailRows = BuildDetailRows(finalResult)
    Dim detailHeaders(1 To 11) As String
    detailHeaders(1) = "物流单号"
    detailHeaders(2) = "WMS退单号"
    detailHeaders(3) = "SKU"
    detailHeaders(4) = "行号"
    detailHeaders(5) = "退单数量"
    detailHeaders(6) = "QC情况"
    detailHeaders(7) = "批号"
    detailHeaders(8) = "效期"
    detailHeaders(9) = "分配数量"
    detailHeaders(10) = "行状态"
    detailHeaders(11) = "退单号状态"
    WriteSheet wb.Worksheets("成功分配明细表"), detailRows, detailHeaders

    ' 数据异常明细表
    Dim anomalyOutputRows() As OutputRow
    anomalyOutputRows = BuildAnomalyOutputRows(anomalyRows)
    Dim anomalyHeaders(1 To 9) As String
    anomalyHeaders(1) = "来源表"
    anomalyHeaders(2) = "Excel行号"
    anomalyHeaders(3) = "物流单号"
    anomalyHeaders(4) = "WMS退单号"
    anomalyHeaders(5) = "SKU"
    anomalyHeaders(6) = "字段名"
    anomalyHeaders(7) = "原始值"
    anomalyHeaders(8) = "错误码"
    anomalyHeaders(9) = "原因说明"
    WriteSheet wb.Worksheets("数据异常明细表"), anomalyOutputRows, anomalyHeaders

    ' 调试日志（WriteDebugLog 内部按 DetailedLogLimit 自动分表）
    Dim debugRows() As OutputRow
    debugRows = BuildDebugLogRows(events, cfg)
    WriteDebugLog wb, debugRows, cfg

    ' 运行历史（每次追加一行，不覆盖历史记录）
    AppendRunHistory wb.Worksheets(SHEET_RUN_HISTORY), runHistoryRow
End Sub


' M12 失败代表系统内部结果不一致，文员改源数据通常无法解决。
' 这里统一转成 E99，让完整运行在写入成功明细前中止。
Private Sub RN_RaisePostValidationE99IfNeeded(ByVal postResult As Object)
    If postResult Is Nothing Then Exit Sub
    If Not postResult.Exists("HasFailures") Then Exit Sub
    If Not CBool(postResult("HasFailures")) Then Exit Sub

    Dim issueCount As Long
    issueCount = 0
    If postResult.Exists("IssueCount") Then issueCount = CLng(postResult("IssueCount"))

    Dim shipNo As String
    Dim sku As String
    shipNo = RN_GetPostIssueText(postResult, 1, "ShipmentNo")
    sku = RN_GetPostIssueText(postResult, 1, "SKU")
    If Len(shipNo) = 0 Then shipNo = "[POST]"
    If Len(sku) = 0 Then sku = "[POST]"

    RaiseE99 shipNo, sku, 0, issueCount, RN_FormatPostValidationFailure(postResult)
End Sub

Private Function RN_FormatPostValidationFailure(ByVal postResult As Object) As String
    If postResult Is Nothing Then
        RN_FormatPostValidationFailure = "ValidatePost 返回空结果"
        Exit Function
    End If

    Dim issueCount As Long
    If postResult.Exists("IssueCount") Then issueCount = CLng(postResult("IssueCount"))
    If issueCount <= 0 Then
        RN_FormatPostValidationFailure = "ValidatePost 标记失败但未提供明细"
        Exit Function
    End If

    Dim msg As String
    msg = "ValidatePost 失败，问题数=" & CStr(issueCount)

    Dim maxShown As Long
    maxShown = issueCount
    If maxShown > 3 Then maxShown = 3

    Dim i As Long
    For i = 1 To maxShown
        msg = msg & "；#" & CStr(i) & " " & _
              RN_GetPostIssueText(postResult, i, "Code") & " " & _
              "物流单号=" & RN_GetPostIssueText(postResult, i, "ShipmentNo") & " " & _
              "WMS退单号=" & RN_GetPostIssueText(postResult, i, "WMSOrderNo") & " " & _
              "SKU=" & RN_GetPostIssueText(postResult, i, "SKU") & " " & _
              "行号=" & RN_GetPostIssueText(postResult, i, "LineNo") & " " & _
              RN_GetPostIssueText(postResult, i, "Message")
    Next i

    If issueCount > maxShown Then msg = msg & "；其余问题请运行 M12 测试或开启开发排查"
    RN_FormatPostValidationFailure = msg
End Function

Private Function RN_GetPostIssueText(ByVal postResult As Object, ByVal issueIndex As Long, ByVal fieldName As String) As String
    If postResult Is Nothing Then Exit Function

    Dim keyName As String
    keyName = "Issue_" & CStr(issueIndex) & "_" & fieldName
    If postResult.Exists(keyName) Then RN_GetPostIssueText = CStr(postResult(keyName))
End Function


' 从所有物流单号分配结果中合并调试日志事件（M09 → M15 → M13）。
Private Function RN_CollectDebugEvents(ByRef shipmentResults() As Object) As AllocationEvent()
    Dim resultCount As Long
    resultCount = RN_ObjectArrayCount(shipmentResults)
    If resultCount = 0 Then Exit Function

    Dim total As Long
    Dim i As Long
    For i = 1 To resultCount
        Dim shipResult As Object
        Set shipResult = shipmentResults(LBound(shipmentResults) + i - 1)
        If shipResult Is Nothing Then GoTo CountNext
        If Not shipResult.Exists("GroupCount") Then GoTo CountNext

        Dim g As Long
        For g = 1 To CLng(shipResult("GroupCount"))
            Dim key As String
            key = "Group_" & g & "_DebugEventCount"
            If shipResult.Exists(key) Then
                total = total + CLng(shipResult(key))
            End If
        Next g
CountNext:
    Next i
    If total <= 0 Then Exit Function

    Dim allEvents() As AllocationEvent
    ReDim allEvents(1 To total)
    Dim outIdx As Long

    For i = 1 To resultCount
        Set shipResult = shipmentResults(LBound(shipmentResults) + i - 1)
        If shipResult Is Nothing Then GoTo MergeNext

        Dim batch() As AllocationEvent
        batch = ExtractDebugEventsFromShipment(shipResult)

        Dim j As Long
        On Error GoTo EmptyBatch
        For j = LBound(batch) To UBound(batch)
            outIdx = outIdx + 1
            allEvents(outIdx) = batch(j)
        Next j
EmptyBatch:
        On Error GoTo 0
MergeNext:
    Next i

    If outIdx = 0 Then Exit Function
    If outIdx < total Then ReDim Preserve allEvents(1 To outIdx)
    RN_CollectDebugEvents = allEvents
End Function


' =============================================================================
' 四、私有长度工具函数（处理 VBA 空数组时 UBound 报错的问题）
' =============================================================================

' 安全获取 NormalizedReturnLine 数组长度，空数组返回 0。
Private Function RN_ReturnLineCount(ByRef arr() As NormalizedReturnLine) As Long
    On Error GoTo EmptyArr
    RN_ReturnLineCount = UBound(arr) - LBound(arr) + 1
    Exit Function
EmptyArr:
    RN_ReturnLineCount = 0
End Function

' 安全获取 NormalizedInventoryLine 数组长度，空数组返回 0。
Private Function RN_InventoryLineCount(ByRef arr() As NormalizedInventoryLine) As Long
    On Error GoTo EmptyArr
    RN_InventoryLineCount = UBound(arr) - LBound(arr) + 1
    Exit Function
EmptyArr:
    RN_InventoryLineCount = 0
End Function

' 安全获取 Object 数组长度，空数组返回 0。
Private Function RN_ObjectArrayCount(ByRef arr() As Object) As Long
    On Error GoTo EmptyArr
    RN_ObjectArrayCount = UBound(arr) - LBound(arr) + 1
    Exit Function
EmptyArr:
    RN_ObjectArrayCount = 0
End Function

' 统计两表去重合并后的物流单号总数（需求 §5.6 “输入：物流单号数”）。
' 空数组安全：任一空表按 0 行处理。
Private Function RN_CountDistinctShipments( _
    ByRef orders() As NormalizedReturnLine, _
    ByRef inventory() As NormalizedInventoryLine) As Long

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim i As Long
    If RN_ReturnLineCount(orders) > 0 Then
        For i = LBound(orders) To UBound(orders)
            If Len(orders(i).ShipmentNo) > 0 Then seen(orders(i).ShipmentNo) = True
        Next i
    End If
    If RN_InventoryLineCount(inventory) > 0 Then
        For i = LBound(inventory) To UBound(inventory)
            If Len(inventory(i).ShipmentNo) > 0 Then seen(inventory(i).ShipmentNo) = True
        Next i
    End If

    RN_CountDistinctShipments = seen.Count
End Function

' 计算两个 Timer 读数之间的耗时（秒，保留 1 位小数），处理午夜回绕。
Private Function RN_ElapsedSecs(ByVal tStart As Single, ByVal tEnd As Single) As Single
    If tEnd < tStart Then tEnd = tEnd + 86400!
    RN_ElapsedSecs = Round(tEnd - tStart, 1)
End Function

' 汇总各错误码命中的物流单号数，格式 “E01:3; E04:2”（按错误码升序）。
' 来源：M05 校验问题（按 错误码+物流单号 去重）+ M09 分配失败组（E09/E10/E99）。
Private Function RN_BuildErrorCodeDistribution( _
    ByRef validationIssues() As ValidationIssue, _
    ByRef shipmentResults() As Object) As String

    ' 先按 错误码|物流单号 去重，再按错误码计数
    Dim pairSeen As Object
    Set pairSeen = CreateObject("Scripting.Dictionary")
    pairSeen.CompareMode = vbTextCompare
    Dim codeCounts As Object
    Set codeCounts = CreateObject("Scripting.Dictionary")
    codeCounts.CompareMode = vbTextCompare

    Dim i As Long
    Dim issueCount As Long
    issueCount = 0
    On Error Resume Next
    issueCount = UBound(validationIssues) - LBound(validationIssues) + 1
    On Error GoTo 0
    If issueCount > 0 Then
        For i = LBound(validationIssues) To UBound(validationIssues)
            Dim code As String
            code = Trim$(validationIssues(i).ErrorCode)
            Dim ship As String
            ship = Trim$(validationIssues(i).ShipmentNo)
            If Len(code) > 0 And Len(ship) > 0 Then
                Dim pairKey As String
                pairKey = code & "|" & ship
                If Not pairSeen.Exists(pairKey) Then
                    pairSeen.Add pairKey, True
                    codeCounts(code) = CLng(codeCounts(code)) + 1
                End If
            End If
        Next i
    End If

    Dim resultCount As Long
    resultCount = RN_ObjectArrayCount(shipmentResults)
    Dim r As Long
    For r = 1 To resultCount
        Dim result As Object
        Set result = shipmentResults(LBound(shipmentResults) + r - 1)
        If result Is Nothing Then GoTo NextResult
        If Not result.Exists("GroupCount") Then GoTo NextResult
        Dim shipNo As String
        shipNo = vbNullString
        If result.Exists("ShipmentNo") Then shipNo = CStr(result("ShipmentNo"))
        Dim g As Long
        For g = 1 To CLng(result("GroupCount"))
            Dim gErrKey As String
            gErrKey = "Group_" & g & "_ErrorCode"
            If result.Exists(gErrKey) Then
                Dim gCode As String
                gCode = Trim$(CStr(result(gErrKey)))
                If Len(gCode) > 0 And Len(shipNo) > 0 Then
                    Dim gPair As String
                    gPair = gCode & "|" & shipNo
                    If Not pairSeen.Exists(gPair) Then
                        pairSeen.Add gPair, True
                        codeCounts(gCode) = CLng(codeCounts(gCode)) + 1
                    End If
                End If
            End If
        Next g
NextResult:
    Next r

    If codeCounts.Count = 0 Then Exit Function

    ' 按错误码升序拼装
    Dim keys() As String
    ReDim keys(1 To codeCounts.Count)
    Dim k As Variant
    i = 0
    For Each k In codeCounts.Keys
        i = i + 1
        keys(i) = CStr(k)
    Next k
    Dim a As Long
    Dim b As Long
    For a = 1 To UBound(keys) - 1
        For b = a + 1 To UBound(keys)
            If keys(a) > keys(b) Then
                Dim tmp As String
                tmp = keys(a)
                keys(a) = keys(b)
                keys(b) = tmp
            End If
        Next b
    Next a

    Dim text As String
    For i = 1 To UBound(keys)
        If Len(text) > 0 Then text = text & "; "
        text = text & keys(i) & ":" & CStr(codeCounts(keys(i)))
    Next i
    RN_BuildErrorCodeDistribution = text
End Function
