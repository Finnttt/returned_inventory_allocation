Option Explicit

' =============================================================================
' M08_分配策略（modStrategies）
' =============================================================================
' 职责：在候选池内依次尝试三级策略，返回分配尝试结果（含 UndoLog）。
' 只解决"当前行如何从候选池拿到 D 件"——不做回溯，不记调试日志。
'
' 公开函数：
'   TryAllocate(pool, demand, ledger)   → AllocationAttempt（Object/Dictionary）
'   CompareByPriority(a, b)             → Integer
'
' AllocationAttempt 内部结构（Scripting.Dictionary）：
'   "Success"      → Boolean   是否分配成功
'   "StrategyUsed" → String    "策略一" / "策略二" / "策略三" / "失败"
'   "UndoLog"      → Object    Collection，M09 回溯时调用 Undo(ledger, attempt("UndoLog"))
'   "DetailCount"  → Long      成功时 >= 1；失败时 = 0
'   对每条明细 i（i = 1 .. DetailCount）：
'     "QC_i"       → String    库存行 QC
'     "LotNo_i"    → String    库存行批号
'     "Expiry_i"   → String    库存行效期（YYYY/MM/DD）
'     "AllocQty_i" → Long      本条明细分配数量
'
' 注：ShipmentNo/WMSOrderNo/SKU/LineNo/OrderQty/LineStatus 由 M09 在接收结果后补填，
'     M08 在 TryAllocate 时还不知道这些上层业务字段。
' =============================================================================

' -----------------------------------------------------------------------------
' 公开函数：TryAllocate
' -----------------------------------------------------------------------------

' 从候选池为当前退单行分配 demand 件库存。
' 执行顺序：策略一 → 策略二 → 策略三，首个成功立即返回。
' 全部失败时返回 Success=False 且账本不变（各策略内部已保证回滚完整）。
'
' 参数：
'   pool   - FilterCandidatePool 返回的候选库存行数组（CandidateRow[]）
'   demand - 当前退单行的需求数量
'   ledger - 库存账本（Object/Scripting.Dictionary，来自 BuildLedger）
' 返回：AllocationAttempt（Object），结构见模块头注释
Public Function TryAllocate(ByRef pool() As CandidateRow, _
                             ByVal demand As Long, _
                             ByVal ledger As Object) As Object

    ' 初始化为失败状态（若所有策略均不成功，直接返回此对象）
    Dim attempt As Object
    Set attempt = CreateObject("Scripting.Dictionary")
    attempt.Add "Success", False
    attempt.Add "StrategyUsed", "失败"
    attempt.Add "DetailCount", CLng(0)

    ' 边界检查：需求量无效或候选池为空时无需尝试
    If demand <= 0 Or M08_SafeCandidateCount(pool) = 0 Then
        attempt.Add "UndoLog", NewUndoLog()
        Set TryAllocate = attempt
        Exit Function
    End If

    ' 按平局优先级对候选池排序（ZP > QC > NG；相同QC则效期降序；再相同则批号升序）
    ' 排序目的：策略一/二可直接取排序后的第一个命中；策略三按同 QC 分组时顺序已正确
    Dim sortedPool() As CandidateRow
    sortedPool = M08_SortPoolByPriority(pool)

    ' 依次尝试三个策略，任一成功即结束
    If StrategyOne(sortedPool, demand, ledger, attempt) Then GoTo Done
    If StrategyTwo(sortedPool, demand, ledger, attempt) Then GoTo Done
    If StrategyThree(sortedPool, demand, ledger, attempt) Then GoTo Done

    ' 全部失败：补充空 UndoLog，确保 M09 可无条件调用 Undo 而不崩溃
    attempt.Add "UndoLog", NewUndoLog()

Done:
    Set TryAllocate = attempt
End Function

' -----------------------------------------------------------------------------
' 公开函数：CompareByPriority
' -----------------------------------------------------------------------------

' 比较两个候选行的排序优先级。
' 返回值约定（与冒泡排序"若 a > b 则交换"语义一致）：
'   < 0  → a 排在 b 前（a 优先）
'   = 0  → 优先级相同
'   > 0  → b 排在 a 前（b 优先）
'
' 三级比较规则：
'   1. QC 优先级：ZP=0 > QC=1 > NG=2（数值小者优先）
'   2. 效期降序：同 QC 时效期更晚者优先
'   3. 批号字母升序：同 QC 同效期时的平局兜底
Public Function CompareByPriority(ByRef a As CandidateRow, _
                                   ByRef b As CandidateRow) As Integer
    ' 规则1：QC 优先级
    Dim ra As Integer
    Dim rb As Integer
    ra = M08_QCPriorityRank(a.QC)
    rb = M08_QCPriorityRank(b.QC)

    If ra < rb Then
        CompareByPriority = -1
        Exit Function
    ElseIf ra > rb Then
        CompareByPriority = 1
        Exit Function
    End If

    ' 规则2：效期降序（标准化 YYYY/MM/DD 的字典序与日期顺序一致）
    If a.Expiry > b.Expiry Then
        CompareByPriority = -1
        Exit Function
    ElseIf a.Expiry < b.Expiry Then
        CompareByPriority = 1
        Exit Function
    End If

    ' 规则3：批号字母升序（平局兜底，确保排序结果可重复）
    If a.LotNo < b.LotNo Then
        CompareByPriority = -1
        Exit Function
    ElseIf a.LotNo > b.LotNo Then
        CompareByPriority = 1
        Exit Function
    End If

    CompareByPriority = 0
End Function

' =============================================================================
' 私有策略函数
' =============================================================================

' 策略一：精确匹配
' 在已排序的候选池中找 CurrentQty = demand 的 lot。
' 由于池已按 CompareByPriority 排序，第一个命中的即为最高优先级的精确匹配。
' 成功：执行扣减，填充 attempt，返回 True。
' 失败：账本不变（无任何扣减操作），返回 False。
Private Function StrategyOne(ByRef sortedPool() As CandidateRow, _
                              ByVal demand As Long, _
                              ByVal ledger As Object, _
                              ByRef attempt As Object) As Boolean
    Dim i As Long
    For i = LBound(sortedPool) To UBound(sortedPool)
        If sortedPool(i).CurrentQty = demand Then
            Dim undoLog As Object
            Set undoLog = NewUndoLog()

            Dim key As InventoryKey
            key = M08_MakeCandidateKey(sortedPool(i))

            If Deduct(ledger, key, demand, undoLog) Then
                attempt("Success") = True
                attempt("StrategyUsed") = "策略一"
                attempt.Add "UndoLog", undoLog
                attempt("DetailCount") = CLng(1)
                attempt.Add "QC_1", sortedPool(i).QC
                attempt.Add "LotNo_1", sortedPool(i).LotNo
                attempt.Add "Expiry_1", sortedPool(i).Expiry
                attempt.Add "AllocQty_1", CLng(demand)
                StrategyOne = True
                Exit Function
            End If
        End If
    Next i

    StrategyOne = False
End Function

' 策略二：最小 sufficient 匹配（找 CurrentQty > demand 中数量最小的 lot）
' 选"最小能覆盖"的 lot，而非最大的，目的是保留较大库存供后续行使用（§剩余库存保留原则）。
' 相同 CurrentQty 时：池已排序（高优先级在前），第一个遇到的即为胜者，无需再比较。
' 成功：执行扣减（扣减量 = demand，非 lot 全量），填充 attempt，返回 True。
' 失败：无 lot > demand，账本不变，返回 False。
Private Function StrategyTwo(ByRef sortedPool() As CandidateRow, _
                              ByVal demand As Long, _
                              ByVal ledger As Object, _
                              ByRef attempt As Object) As Boolean
    ' 初始化"最优"候选下标和数量（用 Long 最大值作哨兵）
    Dim bestIdx As Long
    Dim bestQty As Long
    bestIdx = -1
    bestQty = 2147483647

    Dim i As Long
    For i = LBound(sortedPool) To UBound(sortedPool)
        If sortedPool(i).CurrentQty > demand Then
            If sortedPool(i).CurrentQty < bestQty Then
                ' 发现更小的 sufficient lot（数量更接近 demand，浪费更少）
                bestIdx = i
                bestQty = sortedPool(i).CurrentQty
            End If
            ' qty 相同时不更新：已排序的池保证先遇到高优先级，不覆盖
        End If
    Next i

    If bestIdx = -1 Then
        StrategyTwo = False
        Exit Function
    End If

    Dim undoLog As Object
    Set undoLog = NewUndoLog()

    Dim key As InventoryKey
    key = M08_MakeCandidateKey(sortedPool(bestIdx))

    If Deduct(ledger, key, demand, undoLog) Then
        attempt("Success") = True
        attempt("StrategyUsed") = "策略二"
        attempt.Add "UndoLog", undoLog
        attempt("DetailCount") = CLng(1)
        attempt.Add "QC_1", sortedPool(bestIdx).QC
        attempt.Add "LotNo_1", sortedPool(bestIdx).LotNo
        attempt.Add "Expiry_1", sortedPool(bestIdx).Expiry
        attempt.Add "AllocQty_1", CLng(demand)
        StrategyTwo = True
        Exit Function
    End If

    StrategyTwo = False
End Function

' 策略三：同 QC 内按“最接近剩余需求量”逐步拼凑（严禁跨 QC 合并）
'   第一步：从全部候选行选 |qty-demand| 最小者；平局按 QC→晚效期→小批号。
'           选中后锁定该行 QC。
'   后续：只在锁定 QC 内选 |qty-remaining| 最小者；距离相同时优先
'         qty>=remaining（一步完成），再按晚效期→小批号。
' FilterCandidatePool 已保证进入候选池的每个 QC 总量均满足当前行可用规则；
' 若调用方传入未经筛选的池且锁定 QC 总量不足，则撤销本次扣减并返回失败。
Private Function StrategyThree(ByRef sortedPool() As CandidateRow, _
                                ByVal demand As Long, _
                                ByVal ledger As Object, _
                                ByRef attempt As Object) As Boolean
    Dim poolSize As Long
    poolSize = UBound(sortedPool) - LBound(sortedPool) + 1

    Dim firstIdx As Long
    firstIdx = M08_FindStrategyThreeFirst(sortedPool, demand)
    If firstIdx = -1 Then
        StrategyThree = False
        Exit Function
    End If

    Dim targetQC As String
    targetQC = sortedPool(firstIdx).QC

    Dim used() As Boolean
    ReDim used(LBound(sortedPool) To UBound(sortedPool))

    ' 最多使用 poolSize 个五元组，先暂存在数组中，成功后再写入 attempt。
    Dim tempQCs()      As String
    Dim tempLots()     As String
    Dim tempExpiries() As String
    Dim tempAllocQtys() As Long
    ReDim tempQCs(1 To poolSize)
    ReDim tempLots(1 To poolSize)
    ReDim tempExpiries(1 To poolSize)
    ReDim tempAllocQtys(1 To poolSize)

    Dim tempLog As Object
    Set tempLog = NewUndoLog()

    Dim remaining As Long
    remaining = demand

    Dim detailCount As Long
    detailCount = 0

    Do While remaining > 0
        Dim selectedIdx As Long
        If detailCount = 0 Then
            selectedIdx = firstIdx
        Else
            selectedIdx = M08_FindStrategyThreeNext( _
                sortedPool, used, targetQC, remaining)
        End If

        If selectedIdx = -1 Then
            Undo ledger, tempLog
            StrategyThree = False
            Exit Function
        End If

        Dim toDeduct As Long
        If sortedPool(selectedIdx).CurrentQty <= remaining Then
            toDeduct = sortedPool(selectedIdx).CurrentQty
        Else
            toDeduct = remaining
        End If

        Dim key As InventoryKey
        key = M08_MakeCandidateKey(sortedPool(selectedIdx))

        If Not Deduct(ledger, key, toDeduct, tempLog) Then
            Undo ledger, tempLog
            StrategyThree = False
            Exit Function
        End If

        used(selectedIdx) = True
        detailCount = detailCount + 1
        tempQCs(detailCount)       = targetQC
        tempLots(detailCount)      = sortedPool(selectedIdx).LotNo
        tempExpiries(detailCount)  = sortedPool(selectedIdx).Expiry
        tempAllocQtys(detailCount) = toDeduct
        remaining = remaining - toDeduct
    Loop

    attempt("Success") = True
    attempt("StrategyUsed") = "策略三"
    attempt.Add "UndoLog", tempLog
    attempt("DetailCount") = CLng(detailCount)

    Dim j As Long
    For j = 1 To detailCount
        attempt.Add "QC_" & j,       tempQCs(j)
        attempt.Add "LotNo_" & j,     tempLots(j)
        attempt.Add "Expiry_" & j,    tempExpiries(j)
        attempt.Add "AllocQty_" & j,  CLng(tempAllocQtys(j))
    Next j

    StrategyThree = True
End Function

' =============================================================================
' 私有辅助函数
' =============================================================================

' 对候选池按 CompareByPriority 排序，返回排好序的副本（原数组不变）。
' 使用冒泡排序：同一 SKU 组的候选 lot 数量通常极小（< 20），性能无忧。
Private Function M08_SortPoolByPriority(ByRef pool() As CandidateRow) As CandidateRow()
    Dim n As Long
    n = UBound(pool) - LBound(pool) + 1

    ' 复制到新数组（统一从 1 起始，简化后续索引运算）
    Dim sorted() As CandidateRow
    ReDim sorted(1 To n)

    Dim i As Long
    For i = 1 To n
        sorted(i) = pool(LBound(pool) + i - 1)
    Next i

    ' 冒泡排序（稳定，行数小时性能足够）
    Dim swapped As Boolean
    Do
        swapped = False
        For i = 1 To n - 1
            If CompareByPriority(sorted(i), sorted(i + 1)) > 0 Then
                Dim tmp As CandidateRow
                tmp = sorted(i)
                sorted(i) = sorted(i + 1)
                sorted(i + 1) = tmp
                swapped = True
            End If
        Next i
    Loop While swapped

    M08_SortPoolByPriority = sorted
End Function

' 策略三第一步：数量距离优先；距离相同才使用统一平局规则。
Private Function M08_FindStrategyThreeFirst( _
    ByRef pool() As CandidateRow, _
    ByVal demand As Long) As Long

    Dim bestIdx As Long
    bestIdx = -1

    Dim bestDiff As Long
    Dim i As Long
    For i = LBound(pool) To UBound(pool)
        If pool(i).CurrentQty > 0 Then
            Dim thisDiff As Long
            thisDiff = M08_QuantityDistance(pool(i).CurrentQty, demand)

            ' VBA 的 Or 不会短路求值，因此不能在同一表达式中先判断
            ' bestIdx=-1 再访问 pool(bestIdx)，必须分层判断以避免下标越界。
            Dim chooseThis As Boolean
            chooseThis = False
            If bestIdx = -1 Then
                chooseThis = True
            ElseIf thisDiff < bestDiff Then
                chooseThis = True
            ElseIf thisDiff = bestDiff Then
                chooseThis = (CompareByPriority(pool(i), pool(bestIdx)) < 0)
            End If

            If chooseThis Then
                bestIdx = i
                bestDiff = thisDiff
            End If
        End If
    Next i

    M08_FindStrategyThreeFirst = bestIdx
End Function

' 策略三后续步骤：锁定 QC 后按距离选行；等距时先选可一次完成者，
' 若覆盖能力也相同，再按晚效期、批号小的顺序确定结果。
Private Function M08_FindStrategyThreeNext( _
    ByRef pool() As CandidateRow, _
    ByRef used() As Boolean, _
    ByVal targetQC As String, _
    ByVal remaining As Long) As Long

    Dim bestIdx As Long
    bestIdx = -1

    Dim bestDiff As Long
    Dim i As Long
    For i = LBound(pool) To UBound(pool)
        If Not used(i) _
            And pool(i).QC = targetQC _
            And pool(i).CurrentQty > 0 Then

            Dim thisDiff As Long
            thisDiff = M08_QuantityDistance(pool(i).CurrentQty, remaining)

            Dim chooseThis As Boolean
            chooseThis = False

            If bestIdx = -1 Then
                chooseThis = True
            ElseIf thisDiff < bestDiff Then
                chooseThis = True
            ElseIf thisDiff = bestDiff Then
                Dim thisCovers As Boolean
                Dim bestCovers As Boolean
                thisCovers = (pool(i).CurrentQty >= remaining)
                bestCovers = (pool(bestIdx).CurrentQty >= remaining)

                If thisCovers And Not bestCovers Then
                    chooseThis = True
                ElseIf thisCovers = bestCovers Then
                    chooseThis = (CompareByPriority(pool(i), pool(bestIdx)) < 0)
                End If
            End If

            If chooseThis Then
                bestIdx = i
                bestDiff = thisDiff
            End If
        End If
    Next i

    M08_FindStrategyThreeNext = bestIdx
End Function

' 避免直接 Abs(a-b) 的符号处理，让数量距离保持 Long 且含义清楚。
Private Function M08_QuantityDistance(ByVal qty As Long, ByVal target As Long) As Long
    If qty >= target Then
        M08_QuantityDistance = qty - target
    Else
        M08_QuantityDistance = target - qty
    End If
End Function

' 从 CandidateRow 构建 InventoryKey，供 Deduct/Undo 调用
Private Function M08_MakeCandidateKey(ByRef cand As CandidateRow) As InventoryKey
    Dim key As InventoryKey
    key.ShipmentNo = cand.ShipmentNo
    key.SKU        = cand.SKU
    key.QC         = cand.QC
    key.LotNo      = cand.LotNo
    key.Expiry     = cand.Expiry
    M08_MakeCandidateKey = key
End Function

' QC 优先级数值（数值越小优先级越高，与 CompareByPriority 排序方向一致）
Private Function M08_QCPriorityRank(ByVal qc As String) As Integer
    Select Case qc
        Case QC_ZP: M08_QCPriorityRank = 0
        Case QC_QC: M08_QCPriorityRank = 1
        Case QC_NG: M08_QCPriorityRank = 2
        Case Else:  M08_QCPriorityRank = 99  ' 未知 QC 排到最后，不中断运行
    End Select
End Function

' 安全统计 CandidateRow 数组元素数
' 未初始化数组调用 UBound 会报错，此函数用 On Error 统一处理，返回 0
Private Function M08_SafeCandidateCount(ByRef pool() As CandidateRow) As Long
    On Error GoTo NotAllocated
    M08_SafeCandidateCount = UBound(pool) - LBound(pool) + 1
    Exit Function
NotAllocated:
    M08_SafeCandidateCount = 0
End Function
