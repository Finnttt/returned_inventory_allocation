Option Explicit

' =============================================================================
' M09_回溯分配引擎（modBacktracking）
' =============================================================================
' 职责：提供两层接口，以物流单号为粒度完成带回溯的分配计算。
'
' 给新手的解释：
'   "回溯"就像在迷宫里找路：
'     先走一条路（给退单行分配一个 QC），如果后面走不通（后续行无法分配），
'     就退回来换一条路（换一个 QC 再试）。
'   如果换来换去都走不通，就报 E09（无解）或 E10（超过允许尝试次数）。
'
' 公开函数（外部调用）：
'   AllocateShipment(shipNo, skuList(), planMap, precheckMap, ledger, cfg) → Object
'     以物流单号为单位，遍历 SKU 组，处理短路和连带回滚。
'
' 私有函数（内部实现细节）：
'   AllocateSKUGroup(plan, precheckArr, ledger, cfg) → Object
'     以单个 SKU 组为单位，管理 choiceStack 和回溯计数。
'
' 返回值说明（Scripting.Dictionary 结构）：
'   AllocateShipment 返回：
'     "ShipmentNo"          → String
'     "GroupCount"          → Long   （SKU 组数量）
'     "Group_g_SKU"         → String  （第 g 组 SKU）
'     "Group_g_Success"     → Boolean
'     "Group_g_ErrorCode"   → String  （"E09"/"E10"/"E99"/"连带回滚"/""）
'     "Group_g_BacktrackCount" → Long
'     "Group_g_PreCheckHit" → String  （""/"预检测A"/"预检测B"）
'     "Group_g_DetailCount" → Long    （成功时 ≥ 1；失败时 = 0）
'     对每个组的每条明细（d = 1..DetailCount）：
'       "Group_g_WMSOrderNo_d", "Group_g_LineNo_d", "Group_g_OrderQty_d"
'       "Group_g_QC_d", "Group_g_LotNo_d", "Group_g_Expiry_d"
'       "Group_g_AllocQty_d", "Group_g_LineStatus_d"
'
' 与其他模块的关系：
'   依赖 M06 TakeSnapshot / Deduct / Undo（账本操作）
'   依赖 M07 FilterCandidatePool（动态筛选候选池）
'   依赖 M08 TryAllocate（单行三级策略分配）
'   依赖 M10 AssertConservation / AssertUndoConsistency / RaiseE99（守卫断言）
'   结果传递给 M11 ApplyRollback 进行整单状态判定
'
' precheckMap 说明：
'   键 = SKU 字符串；值 = Array(PrecheckAHit As Boolean, PrecheckBHit As Boolean)
'   用 Variant Array 而非 PrecheckResult Type，是因为 VBA Type 无法直接存入 Dictionary。
' =============================================================================


' =============================================================================
' 一、公开函数：AllocateShipment
' =============================================================================

' 以物流单号为单位，遍历所有 SKU 组，逐组调用私有 AllocateSKUGroup 完成分配。
'
' 短路逻辑：某 SKU 组返回 E09/E10/E99 时立即停止遍历，
' 为后续未处理的 SKU 组生成 ErrorCode="连带回滚" 的 GroupAllocResult。
'
' 参数：
'   shipNo      - 当前处理的物流单号
'   skuList()   - 该物流单号下所有需要分配的 SKU 字符串数组
'   planMap     - Object（Dictionary），键=SKU，值=BuildStaticPlan 返回的 StaticPlan
'   precheckMap - Object（Dictionary），键=SKU，值=Array(PrecheckAHit, PrecheckBHit)
'   ledger      - 当前账本（Object，来自 BuildLedger）
'   cfg         - 配置参数（含 MaxBacktrackCount、DebugLogLevel 等）
'
' 返回：Scripting.Dictionary，结构见模块头注释
Public Function AllocateShipment(ByVal shipNo As String, _
                                   ByRef skuList() As String, _
                                   ByVal planMap As Object, _
                                   ByVal precheckMap As Object, _
                                   ByVal ledger As Object, _
                                   ByRef cfg As ConfigStruct) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "ShipmentNo", shipNo

    ' 统计 SKU 列表长度
    Dim skuCount As Long
    skuCount = BT_SafeStringArrayCount(skuList)
    result.Add "GroupCount", skuCount

    If skuCount = 0 Then
        Set AllocateShipment = result
        Exit Function
    End If

    Dim g As Long           ' 组序号（1..skuCount）
    Dim shortCircuit As Boolean
    shortCircuit = False    ' 是否已触发短路

    For g = 1 To skuCount
        Dim sku As String
        sku = skuList(LBound(skuList) + g - 1)

        ' 写入组 SKU（无论成功还是连带回滚都需要这个字段）
        result.Add "Group_" & g & "_SKU", sku

        If shortCircuit Then
            ' 前面某组已失败，本组标记"连带回滚"，不再分配
            result.Add "Group_" & g & "_Success", False
            result.Add "Group_" & g & "_ErrorCode", "连带回滚"
            result.Add "Group_" & g & "_BacktrackCount", CLng(0)
            result.Add "Group_" & g & "_PreCheckHit", ""
            result.Add "Group_" & g & "_DetailCount", CLng(0)
            If planMap.Exists(sku) Then
                Dim scPlan As Object
                Set scPlan = planMap(sku)
                Dim scResult As Object
                Set scResult = CreateObject("Scripting.Dictionary")
                scResult.Add "ShipmentNo", shipNo
                scResult.Add "SKU", sku
                scResult.Add "Success", False
                scResult.Add "ErrorCode", "连带回滚"
                scResult.Add "BacktrackCount", CLng(0)
                scResult.Add "PreCheckHit", ""
                scResult.Add "DetailCount", CLng(0)
                BT_AttachDebugEvents scResult, scPlan, ledger, cfg, True
                BT_MergeGroupDebugEvents result, scResult, g
            End If
        Else
            ' 正常分配：取该 SKU 的 plan 和 precheck 结论
            Dim plan As Object
            Dim precheckArr As Variant

            If planMap.Exists(sku) Then
                Set plan = planMap(sku)
            Else
                ' 找不到计划，视为 E09
                Set plan = CreateObject("Scripting.Dictionary")
            End If

            If precheckMap.Exists(sku) Then
                precheckArr = precheckMap(sku)
            Else
                precheckArr = Array(False, False)
            End If

            ' 调用私有函数分配该 SKU 组
            Dim groupResult As Object
            Set groupResult = AllocateSKUGroup(plan, precheckArr, ledger, cfg)

            ' 将组结果展开写入 result 字典
            result.Add "Group_" & g & "_Success",       CBool(groupResult("Success"))
            result.Add "Group_" & g & "_ErrorCode",     CStr(groupResult("ErrorCode"))
            result.Add "Group_" & g & "_BacktrackCount", CLng(groupResult("BacktrackCount"))
            result.Add "Group_" & g & "_PreCheckHit",   CStr(groupResult("PreCheckHit"))

            Dim dc As Long
            dc = CLng(groupResult("DetailCount"))
            result.Add "Group_" & g & "_DetailCount", dc

            ' 明细字段（仅成功时有值）
            Dim d As Long
            For d = 1 To dc
                result.Add "Group_" & g & "_WMSOrderNo_" & d, CStr(groupResult("WMSOrderNo_" & d))
                result.Add "Group_" & g & "_LineNo_"     & d, CStr(groupResult("LineNo_" & d))
                result.Add "Group_" & g & "_OrderQty_"   & d, CLng(groupResult("OrderQty_" & d))
                result.Add "Group_" & g & "_QC_"         & d, CStr(groupResult("QC_" & d))
                result.Add "Group_" & g & "_LotNo_"      & d, CStr(groupResult("LotNo_" & d))
                result.Add "Group_" & g & "_Expiry_"     & d, CStr(groupResult("Expiry_" & d))
                result.Add "Group_" & g & "_AllocQty_"   & d, CLng(groupResult("AllocQty_" & d))
                result.Add "Group_" & g & "_LineStatus_" & d, CStr(groupResult("LineStatus_" & d))
                If groupResult.Exists("StrategyUsed_" & d) Then
                    result.Add "Group_" & g & "_StrategyUsed_" & d, CStr(groupResult("StrategyUsed_" & d))
                End If
            Next d

            BT_MergeGroupDebugEvents result, groupResult, g

            ' 若本组失败（E09/E10/E99），触发短路，后续组全部连带回滚
            Dim ec As String
            ec = CStr(groupResult("ErrorCode"))
            If ec = ERR_E09 Or ec = ERR_E10 Or ec = ERR_E99 Then
                shortCircuit = True
                ' E99 需要向上层（M15）抛出，调用 RaiseE99 传递错误
                ' M15 通过 On Error GoTo 捕获；当前阶段仅记录，不影响已完成的组结果
                If ec = ERR_E99 Then
                    ' 从 groupResult 中提取 E99 消息（AllocateSKUGroup 已写入 Debug.Print）
                    ' 此处 re-raise，由 M15 的错误处理器统一处理
                    Err.Raise E99_ERROR_NUMBER, "modBacktracking.AllocateShipment", _
                        "[E99] 物流单号=" & shipNo & " SKU=" & sku & " 库存守恒异常"
                End If
            End If
        End If
    Next g

    Set AllocateShipment = result
End Function


' =============================================================================
' 二、私有函数：AllocateSKUGroup
' =============================================================================

' 以单个 SKU 组为单位，使用深度优先回溯算法完成分配。
'
' 算法流程（每次循环处理 currentRow）：
'   1. 排除已尝试的 QC，动态构建候选池（FilterCandidatePool）
'   2. 调用 TryAllocate 尝试分配
'   3. 成功 → 提交到 choiceStack，前进到下一行
'   4. 失败 → 回溯：撤销上一行的提交，将其使用的 QC 加入已尝试列表，重试
'   5. 若第1行就失败（无法回溯）→ E09
'   6. 若回溯次数超过上限 → E10（先撤销所有已提交行）
'
' 成功后调用 AssertConservation 验证库存守恒，失败时返回 E99。
'
' 参数：
'   plan         - BuildStaticPlan 返回的 StaticPlan Object（含排序行序列）
'   precheckArr  - Variant Array(PrecheckAHit, PrecheckBHit)
'   ledger       - 当前账本（Object）
'   cfg          - 配置参数
'
' 返回：Scripting.Dictionary，字段说明见模块头注释
Private Function AllocateSKUGroup(ByVal plan As Object, _
                                   ByRef precheckArr As Variant, _
                                   ByVal ledger As Object, _
                                   ByRef cfg As ConfigStruct) As Object

    ' --- 初始化返回字典 ---
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim shipNo As String
    Dim sku As String

    ' 防御：plan 为空或不含 ShipmentNo 时快速失败
    If plan Is Nothing Then
        result.Add "ShipmentNo", ""
        result.Add "SKU", ""
        result.Add "Success", False
        result.Add "ErrorCode", ERR_E09
        result.Add "BacktrackCount", CLng(0)
        result.Add "PreCheckHit", ""
        result.Add "DetailCount", CLng(0)
        Set AllocateSKUGroup = result
        Exit Function
    End If

    If plan.Exists("ShipmentNo") Then
        shipNo = CStr(plan("ShipmentNo"))
    End If
    If plan.Exists("SKU") Then
        sku = CStr(plan("SKU"))
    End If

    result.Add "ShipmentNo", shipNo
    result.Add "SKU", sku
    result.Add "Success", False
    result.Add "ErrorCode", ""
    result.Add "BacktrackCount", CLng(0)
    result.Add "PreCheckHit", ""
    result.Add "DetailCount", CLng(0)

    ' --- 检查预检测结论：直接失败则跳过分配循环 ---
    ' precheckArr(0) = PrecheckAHit，precheckArr(1) = PrecheckBHit
    Dim precheckAHit As Boolean
    Dim precheckBHit As Boolean
    precheckAHit = CBool(precheckArr(0))
    precheckBHit = CBool(precheckArr(1))

    If precheckAHit Then
        result("ErrorCode") = ERR_E09
        result("PreCheckHit") = "预检测A"
        BT_AttachDebugEvents result, plan, ledger, cfg
        Set AllocateSKUGroup = result
        Exit Function
    End If
    If precheckBHit Then
        result("ErrorCode") = ERR_E09
        result("PreCheckHit") = "预检测B"
        BT_AttachDebugEvents result, plan, ledger, cfg
        Set AllocateSKUGroup = result
        Exit Function
    End If

    ' --- 获取行数 ---
    Dim rowCount As Long
    If Not plan.Exists("RowCount") Then
        result("ErrorCode") = ERR_E09
        Set AllocateSKUGroup = result
        Exit Function
    End If
    rowCount = CLng(plan("RowCount"))
    If rowCount = 0 Then
        result("ErrorCode") = ERR_E09
        Set AllocateSKUGroup = result
        Exit Function
    End If

    ' --- 拍摄快照（分配前的库存状态，供 AssertConservation 使用）---
    Dim snapshot As Object
    Set snapshot = TakeSnapshot(ledger, shipNo, sku)

    ' --- 初始化回溯状态 ---
    Dim backtrackCount As Long
    backtrackCount = 0

    ' choiceStack(i)：存储第 i 行分配成功的 TryAllocate 返回字典（Object）
    ' VBA 不支持 Object() 直接 ReDim，改用 Variant() 存放 Object 引用
    Dim choiceStack() As Object
    ReDim choiceStack(1 To rowCount)

    ' triedQCsByRow(i)：第 i 行已经尝试过的 QC，逗号分隔（如 "ZP,QC"）
    ' 每次回溯到第 i 行时，把之前使用的 QC 追加进来
    Dim triedQCsByRow() As String
    ReDim triedQCsByRow(1 To rowCount)
    Dim ri As Long
    For ri = 1 To rowCount
        triedQCsByRow(ri) = ""
    Next ri

    ' 详细日志使用的过程事件：简版不输出这些行，只保留最终结果行。
    Dim processEvents() As AllocationEvent
    Dim processEventCount As Long

    ' --- 主分配/回溯循环 ---
    Dim currentRow As Long
    currentRow = 1

    Do While currentRow >= 1 And currentRow <= rowCount

        ' 将逗号串解析为字符串数组，传给 FilterCandidatePool
        Dim triedArr() As String
        triedArr = BT_ParseTriedQCs(triedQCsByRow(currentRow))

        ' 获取当前行的行键和需求量
        Dim currentLineKey As String
        Dim demand As Long
        currentLineKey = CStr(plan("LineKey_" & currentRow))
        demand         = CLng(plan("Qty_" & currentRow))

        ' 动态筛选候选池（已尝试的 QC 被排除在外）
        Dim pool() As CandidateRow
        pool = FilterCandidatePool(currentLineKey, plan, ledger, triedArr)

        ' 尝试用三级策略分配当前行
        Dim attempt As Object
        Set attempt = TryAllocate(pool, demand, ledger)

        If CBool(attempt("Success")) Then
            BT_AppendProcessAttemptEvent processEvents, processEventCount, cfg, plan, _
                currentRow, pool, triedQCsByRow(currentRow), attempt, backtrackCount, _
                "过程-尝试成功", "", ""

            ' 分配成功：推入 choiceStack，前进到下一行
            Set choiceStack(currentRow) = attempt
            currentRow = currentRow + 1

        Else
            BT_AppendProcessAttemptEvent processEvents, processEventCount, cfg, plan, _
                currentRow, pool, triedQCsByRow(currentRow), attempt, backtrackCount, _
                "过程-尝试失败", ERR_E09, "动态分配无可用QC"

            ' 分配失败（候选池为空或三级策略均不满足）
            If currentRow = 1 Then
                ' 第一行就失败，没有可回溯的历史行 → 彻底无解，E09
                currentRow = 0   ' 置为 0 作为"耗尽"标志，退出循环后检测
            Else
                ' 回溯：撤销上一行的提交，换一个 QC 重试
                backtrackCount = backtrackCount + 1

                If backtrackCount > cfg.MaxBacktrackCount Then
                    ' 超过回溯上限 → E10
                    ' 必须先撤销所有已提交行，保证账本恢复原状
                    For ri = currentRow - 1 To 1 Step -1
                        If Not (choiceStack(ri) Is Nothing) Then
                            BT_AppendProcessRevokeEvent processEvents, processEventCount, cfg, plan, _
                                ri, choiceStack(ri), backtrackCount
                            Undo ledger, choiceStack(ri)("UndoLog")
                            Set choiceStack(ri) = Nothing
                        End If
                    Next ri

                    result("ErrorCode")     = ERR_E10
                    result("BacktrackCount") = backtrackCount

                    ' 守卫检查：E10 回滚后账本应与快照完全一致
                    If Not AssertUndoConsistency(snapshot, Nothing, ledger) Then
                        ' 回滚本身出了问题，立即升级为 E99，不能继续输出任何成功结果。
                        RaiseE99 shipNo, sku, _
                            BT_CalcObjectTotal(snapshot), _
                            BT_CalcLedgerRangeTotal(snapshot, ledger), _
                            "AssertUndoConsistency after E10 rollback"
                    End If

                    BT_AttachDebugEventsWithProcess result, plan, ledger, cfg, processEvents, processEventCount
                    Set AllocateSKUGroup = result
                    Exit Function
                End If

                ' 撤销上一行（prevRow = currentRow - 1）的已提交分配
                Dim prevRow As Long
                prevRow = currentRow - 1

                ' 获取上一行使用的 QC（strategy 一/二/三都只用同一个 QC）
                Dim usedQC As String
                usedQC = CStr(choiceStack(prevRow)("QC_1"))

                BT_AppendProcessRevokeEvent processEvents, processEventCount, cfg, plan, _
                    prevRow, choiceStack(prevRow), backtrackCount
                Undo ledger, choiceStack(prevRow)("UndoLog")
                Set choiceStack(prevRow) = Nothing

                ' 把刚用过的 QC 加入该行的已尝试列表，下次跳过它
                BT_AddTriedQC triedQCsByRow, prevRow, usedQC

                ' 清除 prevRow 之后所有行的已尝试记录
                ' （因为回溯回来后，后面的行将重新从头尝试）
                For ri = prevRow + 1 To rowCount
                    triedQCsByRow(ri) = ""
                Next ri

                ' 回到 prevRow 重新尝试
                currentRow = prevRow
            End If
        End If
    Loop

    ' --- 循环结束后判断结果 ---

    If currentRow > rowCount Then
        ' 所有行均成功分配，构建 AllocationDetail 数组
        Dim details() As AllocationDetail
        Dim detailCount As Long
        detailCount = BT_BuildDetails(plan, choiceStack, rowCount, details)

        ' 守卫断言：验证库存守恒等式
        Dim guardOK As Boolean
        guardOK = AssertConservation(snapshot, ledger, details)

        If Not guardOK Then
            ' 守恒失败 → E99
            ' 计算期望值（快照总量）和实际值（账本剩余 + 已分配），用于 E99 消息
            Dim expVal As Long
            Dim actLedger As Long
            Dim actAlloc As Long
            expVal    = BT_CalcObjectTotal(snapshot)
            actLedger = BT_CalcLedgerRangeTotal(snapshot, ledger)
            actAlloc  = BT_CalcDetailTotal(details, detailCount)

            RaiseE99 shipNo, sku, expVal, actLedger + actAlloc, "AssertConservation"
        End If

        ' 守恒通过：填写成功结果
        result("Success")        = True
        result("BacktrackCount") = backtrackCount
        result("DetailCount")    = detailCount

        ' 将 AllocationDetail 数组展开写入返回字典
        BT_WriteDetailsToResult result, details, detailCount

    Else
        ' currentRow <= 0：第一行就失败，彻底无解 → E09
        result("ErrorCode")      = ERR_E09
        result("BacktrackCount") = backtrackCount
    End If

    BT_AttachDebugEventsWithProcess result, plan, ledger, cfg, processEvents, processEventCount
    Set AllocateSKUGroup = result
End Function


' =============================================================================
' 三、私有辅助函数（前缀 BT_ 避免与其他模块同名函数冲突）
' =============================================================================

' BT_ParseTriedQCs：将逗号分隔的已尝试QC字符串解析为字符串数组。
' 供 FilterCandidatePool 的 triedQCs 参数使用。
' 示例："ZP,QC" → String array {"ZP", "QC"}
' 空字符串 → 未初始化数组（M07_SafeStringArrCount 对未初始化数组返回 0，效果等同"无已尝试"）
Private Function BT_ParseTriedQCs(ByVal triedStr As String) As String()
    If Len(Trim(triedStr)) = 0 Then
        ' 返回未初始化数组：M07_IsInTriedQCs 内部 SafeCount 返回 0，跳过所有 QC 过滤
        Dim emptyArr() As String
        BT_ParseTriedQCs = emptyArr
        Exit Function
    End If
    BT_ParseTriedQCs = Split(triedStr, ",")
End Function

' BT_AddTriedQC：向 triedQCsByRow(pos) 追加一个已尝试的 QC。
' 若列表非空则用逗号隔开；避免重复追加同一 QC。
Private Sub BT_AddTriedQC(ByRef triedQCsByRow() As String, _
                            ByVal pos As Long, _
                            ByVal qc As String)
    Dim current As String
    current = triedQCsByRow(pos)

    ' 防止重复追加（虽然正常流程下不会重复，加此检查更健壮）
    If Len(current) = 0 Then
        triedQCsByRow(pos) = qc
    ElseIf InStr("," & current & ",", "," & qc & ",") = 0 Then
        triedQCsByRow(pos) = current & "," & qc
    End If
End Sub

' BT_BuildDetails：从 choiceStack 提取所有行的分配结果，合并为 AllocationDetail 数组。
' 同时用 plan 中的 WMSOrderNo/LineNo/Qty 补全上层业务字段，用策略名决定 LineStatus。
' 返回值：总明细条数
Private Function BT_BuildDetails(ByVal plan As Object, _
                                   ByRef choiceStack() As Object, _
                                   ByVal rowCount As Long, _
                                   ByRef outDetails() As AllocationDetail) As Long

    Dim shipNo As String
    Dim sku As String
    shipNo = CStr(plan("ShipmentNo"))
    sku    = CStr(plan("SKU"))

    ' 第一遍：统计总明细数（每行可能有 1 条或多条明细，策略三跨批号时有多条）
    Dim total As Long
    total = 0
    Dim r As Long
    For r = 1 To rowCount
        If Not (choiceStack(r) Is Nothing) Then
            total = total + CLng(choiceStack(r)("DetailCount"))
        End If
    Next r

    If total = 0 Then
        BT_BuildDetails = 0
        Exit Function
    End If

    ReDim outDetails(1 To total)
    Dim idx As Long
    idx = 1

    ' 第二遍：逐行逐明细填充
    For r = 1 To rowCount
        If choiceStack(r) Is Nothing Then GoTo NextRow

        Dim wmsOrderNo As String
        Dim lineNo As String
        Dim orderQty As Long
        wmsOrderNo = CStr(plan("WMSOrderNo_" & r))
        lineNo     = CStr(plan("LineNo_" & r))
        orderQty   = CLng(plan("Qty_" & r))

        Dim strategyUsed As String
        strategyUsed = CStr(choiceStack(r)("StrategyUsed"))
        Dim lineStatus As String
        lineStatus = BT_GetLineStatus(strategyUsed)

        Dim dc As Long
        dc = CLng(choiceStack(r)("DetailCount"))

        Dim d As Long
        For d = 1 To dc
            outDetails(idx).ShipmentNo = shipNo
            outDetails(idx).WMSOrderNo = wmsOrderNo
            outDetails(idx).SKU        = sku
            outDetails(idx).LineNo     = lineNo
            outDetails(idx).OrderQty   = orderQty
            outDetails(idx).QC         = CStr(choiceStack(r)("QC_" & d))
            outDetails(idx).LotNo      = CStr(choiceStack(r)("LotNo_" & d))
            outDetails(idx).Expiry     = CStr(choiceStack(r)("Expiry_" & d))
            outDetails(idx).AllocQty   = CLng(choiceStack(r)("AllocQty_" & d))
            outDetails(idx).LineStatus = lineStatus
            outDetails(idx).StrategyUsed = strategyUsed
            idx = idx + 1
        Next d

NextRow:
    Next r

    BT_BuildDetails = total
End Function

' BT_GetLineStatus：根据策略名返回行状态。
' 策略一/二都只使用一个五元组，文员可批量导入；策略三跨批号/效期拼凑，需要手工操作。
Private Function BT_GetLineStatus(ByVal strategyUsed As String) As String
    Select Case strategyUsed
        Case "策略一", "策略二"
            BT_GetLineStatus = "批量导入"
        Case "策略三"
            BT_GetLineStatus = "手工操作"
        Case Else
            BT_GetLineStatus = "分配失败"
    End Select
End Function

' BT_WriteDetailsToResult：将 AllocationDetail 数组逐字段写入返回字典 result。
' 键名格式：WMSOrderNo_d, LineNo_d, OrderQty_d, QC_d, LotNo_d, Expiry_d, AllocQty_d, LineStatus_d
Private Sub BT_WriteDetailsToResult(ByVal result As Object, _
                                     ByRef details() As AllocationDetail, _
                                     ByVal detailCount As Long)
    Dim d As Long
    For d = 1 To detailCount
        result.Add "WMSOrderNo_" & d, details(LBound(details) + d - 1).WMSOrderNo
        result.Add "LineNo_"     & d, details(LBound(details) + d - 1).LineNo
        result.Add "OrderQty_"   & d, details(LBound(details) + d - 1).OrderQty
        result.Add "QC_"         & d, details(LBound(details) + d - 1).QC
        result.Add "LotNo_"      & d, details(LBound(details) + d - 1).LotNo
        result.Add "Expiry_"     & d, details(LBound(details) + d - 1).Expiry
        result.Add "AllocQty_"   & d, details(LBound(details) + d - 1).AllocQty
        result.Add "LineStatus_" & d, details(LBound(details) + d - 1).LineStatus
        result.Add "StrategyUsed_" & d, details(LBound(details) + d - 1).StrategyUsed
    Next d
End Sub

' BT_CalcObjectTotal：计算快照字典中所有条目的 CurrentQty 之和（index=1）。
' 用于在 AssertConservation 失败时构造 E99 消息的"期望值"。
Private Function BT_CalcObjectTotal(ByVal snapDict As Object) As Long
    Dim total As Long
    total = 0
    If snapDict Is Nothing Then
        BT_CalcObjectTotal = 0
        Exit Function
    End If
    Dim key As Variant
    For Each key In snapDict.Keys
        Dim arr As Variant
        arr = snapDict(key)
        total = total + CLng(arr(1))
    Next key
    BT_CalcObjectTotal = total
End Function

' BT_CalcLedgerRangeTotal：计算账本在快照范围内的当前 CurrentQty 之和（index=1）。
' 用于在 AssertConservation 失败时构造 E99 消息的"实际值（账本部分）"。
Private Function BT_CalcLedgerRangeTotal(ByVal snapDict As Object, _
                                          ByVal ledger As Object) As Long
    Dim total As Long
    total = 0
    If snapDict Is Nothing Or ledger Is Nothing Then
        BT_CalcLedgerRangeTotal = 0
        Exit Function
    End If
    Dim key As Variant
    For Each key In snapDict.Keys
        If ledger.Exists(key) Then
            Dim arr As Variant
            arr = ledger(key)
            total = total + CLng(arr(1))
        End If
    Next key
    BT_CalcLedgerRangeTotal = total
End Function

' BT_CalcDetailTotal：计算 AllocationDetail 数组中 AllocQty 之和。
' 用于在 AssertConservation 失败时构造 E99 消息的"实际值（明细部分）"。
Private Function BT_CalcDetailTotal(ByRef details() As AllocationDetail, _
                                     ByVal detailCount As Long) As Long
    Dim total As Long
    total = 0
    Dim i As Long
    For i = 1 To detailCount
        total = total + details(LBound(details) + i - 1).AllocQty
    Next i
    BT_CalcDetailTotal = total
End Function

' BT_SafeStringArrayCount：安全获取字符串数组长度（未初始化时返回 0）。
Private Function BT_SafeStringArrayCount(ByRef arr() As String) As Long
    On Error GoTo NotInit
    BT_SafeStringArrayCount = UBound(arr) - LBound(arr) + 1
    Exit Function
NotInit:
    BT_SafeStringArrayCount = 0
End Function


' =============================================================================
' 四、调试日志事件（19 列，M09 → M13）
' =============================================================================

' 从单次物流单号分配结果中提取全部调试事件（跨 SKU 组）。
Public Function ExtractDebugEventsFromShipment(ByVal shipResult As Object) As AllocationEvent()
    If shipResult Is Nothing Then Exit Function
    If Not shipResult.Exists("GroupCount") Then Exit Function

    Dim groupCount As Long
    groupCount = CLng(shipResult("GroupCount"))
    If groupCount <= 0 Then Exit Function

    Dim total As Long
    Dim g As Long
    For g = 1 To groupCount
        Dim key As String
        key = "Group_" & g & "_DebugEventCount"
        If shipResult.Exists(key) Then
            total = total + CLng(shipResult(key))
        End If
    Next g
    If total <= 0 Then Exit Function

    Dim events() As AllocationEvent
    ReDim events(1 To total)
    Dim outIdx As Long
    For g = 1 To groupCount
        BT_ReadGroupDebugEvents shipResult, g, events, outIdx
    Next g

    ExtractDebugEventsFromShipment = events
End Function

' 带过程事件的调试日志写入：详细模式输出过程+最终，简版仍只输出最终。
Private Sub BT_AttachDebugEventsWithProcess( _
    ByVal groupResult As Object, _
    ByVal plan As Object, _
    ByVal ledger As Object, _
    ByRef cfg As ConfigStruct, _
    ByRef processEvents() As AllocationEvent, _
    ByVal processEventCount As Long, _
    Optional ByVal crossSkuShort As Boolean = False)

    BT_AttachDebugEvents groupResult, plan, ledger, cfg, crossSkuShort
    If cfg.DebugLogLevel <> DEBUG_LEVEL_DETAIL Then Exit Sub
    If processEventCount <= 0 Then Exit Sub
    If Not groupResult.Exists("DebugEventCount") Then Exit Sub

    Dim finalCount As Long
    finalCount = CLng(groupResult("DebugEventCount"))
    If finalCount <= 0 Then Exit Sub

    Dim combined() As AllocationEvent
    ReDim combined(1 To processEventCount + finalCount)

    Dim outIdx As Long
    Dim i As Long
    For i = 1 To processEventCount
        outIdx = outIdx + 1
        combined(outIdx) = processEvents(i)
    Next i

    BT_ReadLocalDebugEvents groupResult, combined, outIdx
    BT_ClearDebugEvents groupResult
    BT_StoreDebugEventsInDict groupResult, combined
End Sub

' 构建并写入 groupResult 的 DebugEvent_* 字段。
Private Sub BT_AttachDebugEvents( _
    ByVal groupResult As Object, _
    ByVal plan As Object, _
    ByVal ledger As Object, _
    ByRef cfg As ConfigStruct, _
    Optional ByVal crossSkuShort As Boolean = False)

    If cfg.DebugLogLevel = DEBUG_LEVEL_OFF Then Exit Sub
    If plan Is Nothing Then Exit Sub
    If Not plan.Exists("RowCount") Then Exit Sub

    Dim rowCount As Long
    rowCount = CLng(plan("RowCount"))
    If rowCount <= 0 Then Exit Sub

    Dim shipNo As String
    Dim sku As String
    shipNo = CStr(plan("ShipmentNo"))
    sku = CStr(plan("SKU"))

    Dim success As Boolean
    Dim errorCode As String
    Dim backtrackCount As Long
    Dim preCheckHit As String
    success = CBool(groupResult("Success"))
    errorCode = CStr(groupResult("ErrorCode"))
    backtrackCount = CLng(groupResult("BacktrackCount"))
    If groupResult.Exists("PreCheckHit") Then preCheckHit = CStr(groupResult("PreCheckHit"))

    Dim failSubType As String
    failSubType = BT_ResolveFailSubType(preCheckHit, errorCode, crossSkuShort)

    Dim firstFailRow As Long
    firstFailRow = BT_FindFirstFailRow(plan, rowCount, success, errorCode, preCheckHit)

    Dim events() As AllocationEvent
    ReDim events(1 To rowCount)

    Dim r As Long
    For r = 1 To rowCount
        With events(r)
            .ShipmentNo = shipNo
            .SKU = sku
            .WMSOrderNo = CStr(plan("WMSOrderNo_" & r))
            .LineNo = CStr(plan("LineNo_" & r))
            .DemandD = CLng(plan("Qty_" & r))
            .ProcessOrder = CStr(r)
            .DynamicNextMinQty = BT_FormatNextMinQty(plan, r)
            .CandidateQCCount = CStr(plan("InitQCCount_" & r))
            .ExcludedQCList = ""
            .BacktrackNo = backtrackCount
            .IsFinalResult = True
            .IsRevoked = False

            If success Then
                BT_FillSuccessDebugFields groupResult, r, plan, ledger, events(r)
            Else
                BT_FillFailureDebugFields events(r), errorCode, failSubType, r, firstFailRow, backtrackCount, crossSkuShort
            End If
        End With
    Next r

    BT_StoreDebugEventsInDict groupResult, events
End Sub

' 记录一次分配尝试过程；只在详细模式下追加，简版不受影响。
Private Sub BT_AppendProcessAttemptEvent( _
    ByRef events() As AllocationEvent, _
    ByRef eventCount As Long, _
    ByRef cfg As ConfigStruct, _
    ByVal plan As Object, _
    ByVal rowIndex As Long, _
    ByRef pool() As CandidateRow, _
    ByVal triedQCList As String, _
    ByVal attempt As Object, _
    ByVal backtrackCount As Long, _
    ByVal lineStatus As String, _
    ByVal errorCode As String, _
    ByVal failSubType As String)

    If cfg.DebugLogLevel <> DEBUG_LEVEL_DETAIL Then Exit Sub
    If plan Is Nothing Then Exit Sub

    Dim evt As AllocationEvent
    BT_FillCommonProcessEvent evt, plan, rowIndex, backtrackCount
    evt.CandidateQCCount = CStr(BT_CountDistinctQCsInPool(pool))
    evt.ExcludedQCList = triedQCList
    evt.LineStatus = lineStatus
    evt.ErrorCode = errorCode
    evt.FailSubType = failSubType
    evt.IsFinalResult = False
    evt.IsRevoked = False

    If attempt Is Nothing Then
        evt.StrategyUsed = "-"
        evt.UsedQC = "-"
        evt.LotExpiryComboCount = "-"
    Else
        evt.StrategyUsed = CStr(attempt("StrategyUsed"))
        evt.UsedQC = BT_AttemptQCList(attempt)
        evt.LotExpiryComboCount = CStr(CLng(attempt("DetailCount")))
    End If

    BT_AppendDebugEvent events, eventCount, evt
End Sub

' 记录一次回溯撤销过程。撤销前记录，便于详细日志看到被撤回的选择。
Private Sub BT_AppendProcessRevokeEvent( _
    ByRef events() As AllocationEvent, _
    ByRef eventCount As Long, _
    ByRef cfg As ConfigStruct, _
    ByVal plan As Object, _
    ByVal rowIndex As Long, _
    ByVal attempt As Object, _
    ByVal backtrackCount As Long)

    If cfg.DebugLogLevel <> DEBUG_LEVEL_DETAIL Then Exit Sub
    If plan Is Nothing Then Exit Sub
    If attempt Is Nothing Then Exit Sub

    Dim evt As AllocationEvent
    BT_FillCommonProcessEvent evt, plan, rowIndex, backtrackCount
    evt.CandidateQCCount = "-"
    evt.ExcludedQCList = "-"
    evt.StrategyUsed = CStr(attempt("StrategyUsed"))
    evt.UsedQC = BT_AttemptQCList(attempt)
    evt.QCBefore = "-"
    evt.QCAfter = "-"
    evt.LotExpiryComboCount = CStr(CLng(attempt("DetailCount")))
    evt.LineStatus = "过程-回溯撤销"
    evt.ErrorCode = ""
    evt.FailSubType = ""
    evt.IsFinalResult = False
    evt.IsRevoked = True

    BT_AppendDebugEvent events, eventCount, evt
End Sub

Private Sub BT_FillCommonProcessEvent( _
    ByRef evt As AllocationEvent, _
    ByVal plan As Object, _
    ByVal rowIndex As Long, _
    ByVal backtrackCount As Long)

    evt.ShipmentNo = CStr(plan("ShipmentNo"))
    evt.SKU = CStr(plan("SKU"))
    evt.WMSOrderNo = CStr(plan("WMSOrderNo_" & rowIndex))
    evt.LineNo = CStr(plan("LineNo_" & rowIndex))
    evt.DemandD = CLng(plan("Qty_" & rowIndex))
    evt.ProcessOrder = CStr(rowIndex)
    evt.DynamicNextMinQty = BT_FormatNextMinQty(plan, rowIndex)
    evt.QCBefore = "-"
    evt.QCAfter = "-"
    evt.BacktrackNo = backtrackCount
    If backtrackCount > 0 Then
        evt.IsBacktrackRetry = "是"
    Else
        evt.IsBacktrackRetry = "否"
    End If
End Sub

Private Sub BT_AppendDebugEvent( _
    ByRef events() As AllocationEvent, _
    ByRef eventCount As Long, _
    ByRef evt As AllocationEvent)

    eventCount = eventCount + 1
    If eventCount = 1 Then
        ReDim events(1 To 1)
    Else
        ReDim Preserve events(1 To eventCount)
    End If
    events(eventCount) = evt
End Sub

Private Sub BT_MergeGroupDebugEvents(ByVal shipResult As Object, ByVal groupResult As Object, ByVal groupIndex As Long)
    If groupResult Is Nothing Then Exit Sub
    If Not groupResult.Exists("DebugEventCount") Then Exit Sub

    Dim count As Long
    count = CLng(groupResult("DebugEventCount"))
    shipResult.Add "Group_" & groupIndex & "_DebugEventCount", count

    Dim i As Long
    For i = 1 To count
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_ShipmentNo", groupResult("Dbg_" & i & "_ShipmentNo")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_SKU", groupResult("Dbg_" & i & "_SKU")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_WMSOrderNo", groupResult("Dbg_" & i & "_WMSOrderNo")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_LineNo", groupResult("Dbg_" & i & "_LineNo")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_DemandD", groupResult("Dbg_" & i & "_DemandD")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_ProcessOrder", groupResult("Dbg_" & i & "_ProcessOrder")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_DynamicNextMinQty", groupResult("Dbg_" & i & "_DynamicNextMinQty")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_CandidateQCCount", groupResult("Dbg_" & i & "_CandidateQCCount")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_ExcludedQCList", groupResult("Dbg_" & i & "_ExcludedQCList")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_StrategyUsed", groupResult("Dbg_" & i & "_StrategyUsed")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_UsedQC", groupResult("Dbg_" & i & "_UsedQC")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_QCBefore", groupResult("Dbg_" & i & "_QCBefore")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_QCAfter", groupResult("Dbg_" & i & "_QCAfter")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_LotExpiryComboCount", groupResult("Dbg_" & i & "_LotExpiryComboCount")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_IsBacktrackRetry", groupResult("Dbg_" & i & "_IsBacktrackRetry")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_BacktrackNo", groupResult("Dbg_" & i & "_BacktrackNo")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_LineStatus", groupResult("Dbg_" & i & "_LineStatus")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_ErrorCode", groupResult("Dbg_" & i & "_ErrorCode")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_FailSubType", groupResult("Dbg_" & i & "_FailSubType")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_IsFinalResult", groupResult("Dbg_" & i & "_IsFinalResult")
        shipResult.Add "Group_" & groupIndex & "_Dbg_" & i & "_IsRevoked", groupResult("Dbg_" & i & "_IsRevoked")
    Next i
End Sub

Private Sub BT_ReadGroupDebugEvents( _
    ByVal shipResult As Object, _
    ByVal groupIndex As Long, _
    ByRef events() As AllocationEvent, _
    ByRef outIdx As Long)

    Dim key As String
    key = "Group_" & groupIndex & "_DebugEventCount"
    If Not shipResult.Exists(key) Then Exit Sub

    Dim count As Long
    Dim i As Long
    count = CLng(shipResult(key))
    For i = 1 To count
        outIdx = outIdx + 1
        events(outIdx).ShipmentNo = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_ShipmentNo"))
        events(outIdx).SKU = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_SKU"))
        events(outIdx).WMSOrderNo = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_WMSOrderNo"))
        events(outIdx).LineNo = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_LineNo"))
        events(outIdx).DemandD = CLng(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_DemandD"))
        events(outIdx).ProcessOrder = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_ProcessOrder"))
        events(outIdx).DynamicNextMinQty = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_DynamicNextMinQty"))
        events(outIdx).CandidateQCCount = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_CandidateQCCount"))
        events(outIdx).ExcludedQCList = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_ExcludedQCList"))
        events(outIdx).StrategyUsed = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_StrategyUsed"))
        events(outIdx).UsedQC = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_UsedQC"))
        events(outIdx).QCBefore = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_QCBefore"))
        events(outIdx).QCAfter = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_QCAfter"))
        events(outIdx).LotExpiryComboCount = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_LotExpiryComboCount"))
        events(outIdx).IsBacktrackRetry = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_IsBacktrackRetry"))
        events(outIdx).BacktrackNo = CLng(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_BacktrackNo"))
        events(outIdx).LineStatus = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_LineStatus"))
        events(outIdx).ErrorCode = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_ErrorCode"))
        events(outIdx).FailSubType = CStr(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_FailSubType"))
        events(outIdx).IsFinalResult = CBool(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_IsFinalResult"))
        events(outIdx).IsRevoked = CBool(shipResult("Group_" & groupIndex & "_Dbg_" & i & "_IsRevoked"))
    Next i
End Sub

Private Sub BT_ReadLocalDebugEvents( _
    ByVal groupResult As Object, _
    ByRef events() As AllocationEvent, _
    ByRef outIdx As Long)

    If Not groupResult.Exists("DebugEventCount") Then Exit Sub

    Dim count As Long
    Dim i As Long
    count = CLng(groupResult("DebugEventCount"))
    For i = 1 To count
        outIdx = outIdx + 1
        events(outIdx).ShipmentNo = CStr(groupResult("Dbg_" & i & "_ShipmentNo"))
        events(outIdx).SKU = CStr(groupResult("Dbg_" & i & "_SKU"))
        events(outIdx).WMSOrderNo = CStr(groupResult("Dbg_" & i & "_WMSOrderNo"))
        events(outIdx).LineNo = CStr(groupResult("Dbg_" & i & "_LineNo"))
        events(outIdx).DemandD = CLng(groupResult("Dbg_" & i & "_DemandD"))
        events(outIdx).ProcessOrder = CStr(groupResult("Dbg_" & i & "_ProcessOrder"))
        events(outIdx).DynamicNextMinQty = CStr(groupResult("Dbg_" & i & "_DynamicNextMinQty"))
        events(outIdx).CandidateQCCount = CStr(groupResult("Dbg_" & i & "_CandidateQCCount"))
        events(outIdx).ExcludedQCList = CStr(groupResult("Dbg_" & i & "_ExcludedQCList"))
        events(outIdx).StrategyUsed = CStr(groupResult("Dbg_" & i & "_StrategyUsed"))
        events(outIdx).UsedQC = CStr(groupResult("Dbg_" & i & "_UsedQC"))
        events(outIdx).QCBefore = CStr(groupResult("Dbg_" & i & "_QCBefore"))
        events(outIdx).QCAfter = CStr(groupResult("Dbg_" & i & "_QCAfter"))
        events(outIdx).LotExpiryComboCount = CStr(groupResult("Dbg_" & i & "_LotExpiryComboCount"))
        events(outIdx).IsBacktrackRetry = CStr(groupResult("Dbg_" & i & "_IsBacktrackRetry"))
        events(outIdx).BacktrackNo = CLng(groupResult("Dbg_" & i & "_BacktrackNo"))
        events(outIdx).LineStatus = CStr(groupResult("Dbg_" & i & "_LineStatus"))
        events(outIdx).ErrorCode = CStr(groupResult("Dbg_" & i & "_ErrorCode"))
        events(outIdx).FailSubType = CStr(groupResult("Dbg_" & i & "_FailSubType"))
        events(outIdx).IsFinalResult = CBool(groupResult("Dbg_" & i & "_IsFinalResult"))
        events(outIdx).IsRevoked = CBool(groupResult("Dbg_" & i & "_IsRevoked"))
    Next i
End Sub

Private Sub BT_ClearDebugEvents(ByVal dict As Object)
    If dict Is Nothing Then Exit Sub
    If Not dict.Exists("DebugEventCount") Then Exit Sub

    Dim count As Long
    count = CLng(dict("DebugEventCount"))

    On Error Resume Next
    dict.Remove "DebugEventCount"

    Dim fields As Variant
    fields = Array("ShipmentNo", "SKU", "WMSOrderNo", "LineNo", "DemandD", _
        "ProcessOrder", "DynamicNextMinQty", "CandidateQCCount", "ExcludedQCList", _
        "StrategyUsed", "UsedQC", "QCBefore", "QCAfter", "LotExpiryComboCount", _
        "IsBacktrackRetry", "BacktrackNo", "LineStatus", "ErrorCode", "FailSubType", _
        "IsFinalResult", "IsRevoked")

    Dim i As Long
    Dim f As Long
    For i = 1 To count
        For f = LBound(fields) To UBound(fields)
            dict.Remove "Dbg_" & i & "_" & CStr(fields(f))
        Next f
    Next i
    On Error GoTo 0
End Sub

Private Sub BT_StoreDebugEventsInDict(ByVal dict As Object, ByRef events() As AllocationEvent)
    Dim count As Long
    count = UBound(events) - LBound(events) + 1
    dict.Add "DebugEventCount", count

    Dim i As Long
    For i = 1 To count
        dict.Add "Dbg_" & i & "_ShipmentNo", events(i).ShipmentNo
        dict.Add "Dbg_" & i & "_SKU", events(i).SKU
        dict.Add "Dbg_" & i & "_WMSOrderNo", events(i).WMSOrderNo
        dict.Add "Dbg_" & i & "_LineNo", events(i).LineNo
        dict.Add "Dbg_" & i & "_DemandD", events(i).DemandD
        dict.Add "Dbg_" & i & "_ProcessOrder", events(i).ProcessOrder
        dict.Add "Dbg_" & i & "_DynamicNextMinQty", events(i).DynamicNextMinQty
        dict.Add "Dbg_" & i & "_CandidateQCCount", events(i).CandidateQCCount
        dict.Add "Dbg_" & i & "_ExcludedQCList", events(i).ExcludedQCList
        dict.Add "Dbg_" & i & "_StrategyUsed", events(i).StrategyUsed
        dict.Add "Dbg_" & i & "_UsedQC", events(i).UsedQC
        dict.Add "Dbg_" & i & "_QCBefore", events(i).QCBefore
        dict.Add "Dbg_" & i & "_QCAfter", events(i).QCAfter
        dict.Add "Dbg_" & i & "_LotExpiryComboCount", events(i).LotExpiryComboCount
        dict.Add "Dbg_" & i & "_IsBacktrackRetry", events(i).IsBacktrackRetry
        dict.Add "Dbg_" & i & "_BacktrackNo", events(i).BacktrackNo
        dict.Add "Dbg_" & i & "_LineStatus", events(i).LineStatus
        dict.Add "Dbg_" & i & "_ErrorCode", events(i).ErrorCode
        dict.Add "Dbg_" & i & "_FailSubType", events(i).FailSubType
        dict.Add "Dbg_" & i & "_IsFinalResult", events(i).IsFinalResult
        dict.Add "Dbg_" & i & "_IsRevoked", events(i).IsRevoked
    Next i
End Sub

Private Function BT_CountDistinctQCsInPool(ByRef pool() As CandidateRow) As Long
    On Error GoTo EmptyPool

    Dim hasZP As Boolean
    Dim hasQC As Boolean
    Dim hasNG As Boolean
    Dim i As Long
    For i = LBound(pool) To UBound(pool)
        Select Case pool(i).QC
            Case QC_ZP: hasZP = True
            Case QC_QC: hasQC = True
            Case QC_NG: hasNG = True
        End Select
    Next i

    If hasZP Then BT_CountDistinctQCsInPool = BT_CountDistinctQCsInPool + 1
    If hasQC Then BT_CountDistinctQCsInPool = BT_CountDistinctQCsInPool + 1
    If hasNG Then BT_CountDistinctQCsInPool = BT_CountDistinctQCsInPool + 1
    Exit Function

EmptyPool:
    BT_CountDistinctQCsInPool = 0
End Function

Private Function BT_AttemptQCList(ByVal attempt As Object) As String
    If attempt Is Nothing Then
        BT_AttemptQCList = "-"
        Exit Function
    End If
    If Not attempt.Exists("DetailCount") Then
        BT_AttemptQCList = "-"
        Exit Function
    End If

    Dim detailCount As Long
    detailCount = CLng(attempt("DetailCount"))
    If detailCount <= 0 Then
        BT_AttemptQCList = "-"
        Exit Function
    End If

    Dim result As String
    Dim i As Long
    For i = 1 To detailCount
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(attempt("QC_" & i))
    Next i

    BT_AttemptQCList = result
End Function

Private Function BT_ResolveFailSubType( _
    ByVal preCheckHit As String, _
    ByVal errorCode As String, _
    ByVal crossSkuShort As Boolean) As String

    If crossSkuShort Then
        BT_ResolveFailSubType = "连带回滚—跨SKU短路"
        Exit Function
    End If
    If preCheckHit = "预检测A" Then
        BT_ResolveFailSubType = "预检测A（初始可用QC=0）"
        Exit Function
    End If
    If preCheckHit = "预检测B" Then
        BT_ResolveFailSubType = "预检测B（强制竞争库存不足）"
        Exit Function
    End If
    If errorCode = ERR_E10 Then
        BT_ResolveFailSubType = "回溯路径穷尽"
        Exit Function
    End If
    If errorCode = ERR_E09 Then
        BT_ResolveFailSubType = "回溯路径穷尽"
        Exit Function
    End If
    BT_ResolveFailSubType = ""
End Function

Private Function BT_FindFirstFailRow( _
    ByVal plan As Object, _
    ByVal rowCount As Long, _
    ByVal success As Boolean, _
    ByVal errorCode As String, _
    ByVal preCheckHit As String) As Long

    If success Then
        BT_FindFirstFailRow = rowCount + 1
        Exit Function
    End If

    If Len(preCheckHit) > 0 Then
        Dim r As Long
        For r = 1 To rowCount
            If CLng(plan("InitQCCount_" & r)) = 0 Then
                BT_FindFirstFailRow = r
                Exit Function
            End If
        Next r
        BT_FindFirstFailRow = 1
        Exit Function
    End If

    BT_FindFirstFailRow = 2
    If rowCount < 2 Then BT_FindFirstFailRow = 1
End Function

Private Sub BT_FillSuccessDebugFields( _
    ByVal groupResult As Object, _
    ByVal planRow As Long, _
    ByVal plan As Object, _
    ByVal ledger As Object, _
    ByRef evt As AllocationEvent)

    Dim lineNo As String
    lineNo = CStr(plan("LineNo_" & planRow))
    Dim detailIdx As Long
    detailIdx = BT_FindDetailIndexByLineNo(groupResult, lineNo)

    If detailIdx > 0 Then
        If groupResult.Exists("StrategyUsed_" & detailIdx) Then
            evt.StrategyUsed = CStr(groupResult("StrategyUsed_" & detailIdx))
        Else
            evt.StrategyUsed = "策略一"
        End If
        evt.UsedQC = CStr(groupResult("QC_" & detailIdx))
        evt.QCBefore = "-"
        evt.QCAfter = "-"
        evt.LotExpiryComboCount = "1"
        evt.LineStatus = CStr(groupResult("LineStatus_" & detailIdx))
        evt.ErrorCode = ""
        evt.FailSubType = ""
    Else
        evt.StrategyUsed = "-"
        evt.UsedQC = "-"
        evt.QCBefore = "-"
        evt.QCAfter = "-"
        evt.LotExpiryComboCount = "-"
        evt.LineStatus = STATUS_BATCH_IMPORT
    End If

    If evt.BacktrackNo > 0 Then
        evt.IsBacktrackRetry = "是"
    Else
        evt.IsBacktrackRetry = "否"
    End If
End Sub

Private Sub BT_FillFailureDebugFields( _
    ByRef evt As AllocationEvent, _
    ByVal errorCode As String, _
    ByVal failSubType As String, _
    ByVal rowIndex As Long, _
    ByVal firstFailRow As Long, _
    ByVal backtrackCount As Long, _
    ByVal crossSkuShort As Boolean)

    evt.LineStatus = LINE_STATUS_FAILED
    evt.StrategyUsed = "-"
    evt.UsedQC = "-"
    evt.QCBefore = "-"
    evt.QCAfter = "-"
    evt.LotExpiryComboCount = "-"

    If crossSkuShort Then
        evt.ErrorCode = errorCode
        evt.FailSubType = failSubType
        evt.ProcessOrder = "-"
        evt.CandidateQCCount = "-"
        evt.IsBacktrackRetry = "-"
        Exit Sub
    End If

    If Len(failSubType) > 0 And (InStr(failSubType, "预检测") > 0) Then
        evt.ErrorCode = ERR_E09
        evt.FailSubType = failSubType
        If CLng(evt.CandidateQCCount) = 0 Then
            evt.FailSubType = "预检测A（初始可用QC=0）"
        End If
    ElseIf rowIndex >= firstFailRow And rowIndex > 1 Then
        evt.ErrorCode = ERR_E09
        evt.FailSubType = "连带回滚—同SKU未到达行"
        evt.ProcessOrder = "-"
        evt.CandidateQCCount = "-"
    ElseIf rowIndex = firstFailRow Then
        evt.ErrorCode = errorCode
        If Len(errorCode) = 0 Then evt.ErrorCode = ERR_E09
        evt.FailSubType = failSubType
        If Len(evt.FailSubType) = 0 Then evt.FailSubType = "动态分配无可用QC"
    Else
        evt.ErrorCode = errorCode
        evt.FailSubType = failSubType
    End If

    If backtrackCount > 0 Then
        evt.IsBacktrackRetry = "是"
    Else
        evt.IsBacktrackRetry = "否"
    End If
End Sub

Private Function BT_FindDetailIndexByLineNo(ByVal groupResult As Object, ByVal lineNo As String) As Long
    If Not groupResult.Exists("DetailCount") Then Exit Function
    Dim dc As Long
    Dim d As Long
    dc = CLng(groupResult("DetailCount"))
    For d = 1 To dc
        If CStr(groupResult("LineNo_" & d)) = lineNo Then
            BT_FindDetailIndexByLineNo = d
            Exit Function
        End If
    Next d
End Function

Private Function BT_FormatNextMinQty(ByVal plan As Object, ByVal rowIndex As Long) As String
    Dim rowCount As Long
    rowCount = CLng(plan("RowCount"))
    If rowIndex >= rowCount Then
        BT_FormatNextMinQty = ""
        Exit Function
    End If

    Dim minQty As Long
    minQty = 0
    Dim hasValue As Boolean
    Dim r As Long
    For r = rowIndex + 1 To rowCount
        If Not hasValue Then
            minQty = CLng(plan("Qty_" & r))
            hasValue = True
        ElseIf CLng(plan("Qty_" & r)) < minQty Then
            minQty = CLng(plan("Qty_" & r))
        End If
    Next r

    If hasValue Then
        BT_FormatNextMinQty = CStr(minQty)
    Else
        BT_FormatNextMinQty = ""
    End If
End Function
