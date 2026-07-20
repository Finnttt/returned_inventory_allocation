Attribute VB_Name = "modSortFilter"
Option Explicit

' =============================================================================
' M07_排序·预检测·QC筛选（modSortFilter）
' =============================================================================
' 职责：分配前的全部准备工作。三个接口严格分离，禁止共享隐藏状态。
'
' 给新手的解释：
'   把一批退单行分配给库存，就像把"任务"分配给"工人"。
'   本模块负责三件事：
'   1. BuildStaticPlan  ——先把任务按"难易程度"静态排好顺序，并给每个任务列出初始可用的库存种类数。
'   2. RunPrecheck      ——快速检查有没有明显不可能完成的任务（两种预检测场景），若有则直接报错，不用尝试。
'   3. FilterCandidatePool——每处理一个任务时，根据当前库存实况，实时筛选这个任务能用哪些库存。
'
' StaticPlan 内部结构（Scripting.Dictionary）：
'   "RowCount"       -> Long    排序后的总行数
'   "GroupMinQty"    -> Long    组内最小需求量（E11 校验中已算出）
'   "ShipmentNo"     -> String  物流单号
'   "SKU"            -> String  SKU编号
'   "WMSOrderNo_i"   -> String  排序后第 i 行的 WMS退单号
'   "LineNo_i"       -> String  排序后第 i 行的行号
'   "LineKey_i"      -> String  复合键 = "WMSOrderNo:LineNo"，供 FilterCandidatePool 唯一定位
'   "Qty_i"          -> Long    排序后第 i 行的需求量
'   "ExcelRowNum_i"  -> Long    排序后第 i 行的 Excel 源行号（供 M09 错误定位使用）
'   "InitQCCount_i"  -> Long    排序后第 i 行的初始可用QC种类数
'
' 关键规则（§4.2.1 R071 / §4.2.2 R072）：
'   可用QC判定：T = D  或  T >= D + nextMinQty（T=该QC当前库存总量，D=本行需求，
'               nextMinQty=剩余待处理行需求量的最小值，最后一行特殊：仅 T=D）
'   静态排序四级：可用QC数升序 → 需求量降序 → WMS退单号升序 → 行号升序
' =============================================================================

' 系统认可的合法 QC 种类数量（ZP / QC / NG），已在 modTypes.bas 中定义常量
Private Const M07_QC_COUNT As Integer = 3

' =============================================================================
' 一、公开函数
' =============================================================================

' BuildStaticPlan：构建静态分配计划（一次性调用，分配开始前执行）
'
' 算法步骤：
'   1. 提取各行需求量到本地数组
'   2. 计算 groupMinQty = min(所有行 Qty)
'   3. 对每行 i 计算：
'      nextMinQty_i = min{Qty_j : j ≠ i}（所有其他行的最小值，静态快照）
'      isOnlyRow = (n=1)（整组只有1行时，该行等同于"最后行"，只看 T=D）
'      initQCCount_i = 满足可用条件的QC种类数
'   4. 按四级规则排序（冒泡排序，行数通常 < 100，性能足够）
'   5. 写入 StaticPlan Dictionary 并返回
'
' 参数：
'   rows   - 同一物流单号+SKU下所有合法退单行（来自 M09 按组拆分后传入）
'   ledger - 当前库存账本（BuildStaticPlan 调用时账本尚未做任何扣减）
' 返回：StaticPlan Object，空输入时返回空 Dictionary（不抛错）
Public Function BuildStaticPlan(ByRef rows() As NormalizedReturnLine, _
                                 ByVal ledger As Object) As Object
    Dim plan As Object
    Set plan = CreateObject("Scripting.Dictionary")

    Dim n As Long
    n = M07_SafeReturnLineCount(rows)
    If n = 0 Then
        Set BuildStaticPlan = plan
        Exit Function
    End If

    Dim shipNo As String
    Dim sku As String
    shipNo = rows(LBound(rows)).ShipmentNo
    sku = rows(LBound(rows)).SKU

    ' --- 步骤1：提取各行关键字段到本地数组（避免重复访问数组元素语法） ---
    Dim qtys() As Long
    Dim wmsONos() As String
    Dim lineNos() As String
    Dim excelRowNums() As Long
    ReDim qtys(1 To n)
    ReDim wmsONos(1 To n)
    ReDim lineNos(1 To n)
    ReDim excelRowNums(1 To n)

    Dim i As Long
    For i = 1 To n
        With rows(LBound(rows) + i - 1)
            qtys(i)         = .Qty
            wmsONos(i)      = .WMSOrderNo
            lineNos(i)      = .LineNo
            excelRowNums(i) = .ExcelRowNum
        End With
    Next i

    ' --- 步骤2：groupMinQty = 全组最小需求量 ---
    Dim groupMinQty As Long
    groupMinQty = qtys(1)
    For i = 2 To n
        If qtys(i) < groupMinQty Then groupMinQty = qtys(i)
    Next i

    ' --- 步骤3：计算每行的 nextMinQty 和初始可用QC数 ---
    ' isOnlyRow = True 时：该组只有一行，该行就是"最后行"，可用条件变为仅 T=D
    Dim initQCCounts() As Long
    ReDim initQCCounts(1 To n)
    Dim isOnlyRow As Boolean
    isOnlyRow = (n = 1)

    For i = 1 To n
        ' nextMinQty_i = 所有其他行需求量的最小值（静态快照）
        ' 用途：筛选"分配当前行后，剩余库存能否覆盖后续最小需求"
        Dim nmq As Long
        nmq = M07_CalcNextMinQtyAll(i, n, qtys)
        initQCCounts(i) = M07_CalcAvailableQCCount(shipNo, sku, qtys(i), nmq, isOnlyRow, ledger)
    Next i

    ' --- 步骤4：确定排序顺序（稳定冒泡排序） ---
    Dim sortIdx() As Long
    ReDim sortIdx(1 To n)
    For i = 1 To n
        sortIdx(i) = i
    Next i

    Dim swapped As Boolean
    Dim tmp As Long
    Do
        swapped = False
        For i = 1 To n - 1
            Dim a As Long
            Dim b As Long
            a = sortIdx(i)
            b = sortIdx(i + 1)
            If M07_ShouldSwap(a, b, initQCCounts, qtys, wmsONos, lineNos) Then
                sortIdx(i) = b
                sortIdx(i + 1) = a
                swapped = True
            End If
        Next i
    Loop While swapped

    ' --- 步骤5：按排序结果写入 StaticPlan Dictionary ---
    plan.Add "RowCount", n
    plan.Add "GroupMinQty", groupMinQty
    plan.Add "ShipmentNo", shipNo
    plan.Add "SKU", sku

    For i = 1 To n
        Dim si As Long
        si = sortIdx(i)
        plan.Add "WMSOrderNo_" & i, wmsONos(si)
        plan.Add "LineNo_" & i, lineNos(si)
        plan.Add "LineKey_" & i, wmsONos(si) & ":" & lineNos(si)
        plan.Add "Qty_" & i, qtys(si)
        plan.Add "ExcelRowNum_" & i, excelRowNums(si)
        plan.Add "InitQCCount_" & i, initQCCounts(si)
    Next i

    Set BuildStaticPlan = plan
End Function


' RunPrecheck：执行预检测A / B（只读，不修改 plan 或 ledger）
'
' 预检测A（§4.2.3 R081）：
'   若排序后某行 initQCCount = 0，说明无论当前库存如何，该行找不到任何可用QC，
'   分配必然失败。直接触发 E09，不进入分配循环，不触发回溯。
'
' 预检测B：
'   若多行均只有同一种QC可用（被"锁定"），而该QC总量 < 这些行的合计需求，
'   说明这批锁定行无法全部得到满足——分配必然失败。
'   判断：forcedCount >= 2 且 QC总量 < 合计需求
'   不命中条件：T=S（总量恰好等于合计需求，精确可行）；T>=S+minQtyOther（总量足够满足
'   锁定行后还能覆盖其他行的最小需求）——两种情况均属 T >= forcedDemand，不触发。
'
' 任一命中立即返回，不继续检测另一项。
Public Function RunPrecheck(ByVal plan As Object, ByVal ledger As Object) As PrecheckResult
    Dim result As PrecheckResult
    result.PrecheckAHit = False
    result.PrecheckBHit = False

    ' 防御：输入为空或不完整时返回默认值（不命中）
    If plan Is Nothing Then
        RunPrecheck = result
        Exit Function
    End If
    If Not plan.Exists("RowCount") Then
        RunPrecheck = result
        Exit Function
    End If

    Dim n As Long
    n = CLng(plan("RowCount"))
    If n = 0 Then
        RunPrecheck = result
        Exit Function
    End If

    Dim shipNo As String
    Dim sku As String
    shipNo = CStr(plan("ShipmentNo"))
    sku = CStr(plan("SKU"))

    ' === 预检测A：任意行 initQCCount = 0 ===
    Dim i As Long
    For i = 1 To n
        If CLng(plan("InitQCCount_" & i)) = 0 Then
            result.PrecheckAHit = True
            RunPrecheck = result
            Exit Function
        End If
    Next i

    ' === 预检测B：多行竞争同一QC，合计需求 > 该QC供应量 ===
    ' isOnlyRow 与 BuildStaticPlan 保持一致：仅 n=1 时才是"最后行模式"
    Dim isOnlyRow As Boolean
    isOnlyRow = (n = 1)

    Dim qcList(1 To M07_QC_COUNT) As String
    qcList(1) = QC_ZP
    qcList(2) = QC_QC
    qcList(3) = QC_NG

    Dim q As Integer
    For q = 1 To M07_QC_COUNT
        Dim thisQC As String
        thisQC = qcList(q)

        Dim forcedDemand As Long
        Dim forcedCount As Long
        forcedDemand = 0
        forcedCount = 0

        For i = 1 To n
            ' 只处理"锁定行"（初始可用QC数恰好为1的行）
            If CLng(plan("InitQCCount_" & i)) = 1 Then
                Dim D As Long
                D = CLng(plan("Qty_" & i))

                ' 重算该行在静态计划中的 nextMinQty（与 BuildStaticPlan 使用相同公式）
                Dim nmq As Long
                nmq = M07_CalcNextMinQtyInPlan(i, n, plan)

                ' 若 thisQC 就是该行唯一可用的QC，将其需求计入竞争总量
                Dim T As Long
                T = QueryQCTotal(ledger, shipNo, sku, thisQC)
                If M07_IsQCAvailableByRule(T, D, nmq, isOnlyRow) Then
                    forcedDemand = forcedDemand + D
                    forcedCount = forcedCount + 1
                End If
            End If
        Next i

        ' 两行或以上都锁定到同一QC，且该QC供应量 < 合计需求 → 预检测B命中
        If forcedCount >= 2 Then
            Dim supply As Long
            supply = QueryQCTotal(ledger, shipNo, sku, thisQC)
            If supply < forcedDemand Then
                result.PrecheckBHit = True
                RunPrecheck = result
                Exit Function
            End If
        End If
    Next q

    RunPrecheck = result
End Function


' FilterCandidatePool：动态筛选当前行的候选QC池（每行分配前调用一次）
'
' 算法：
'   1. 在 StaticPlan 中按复合键 currentLineKey 定位当前行的位置 p
'   2. 计算动态 nextMinQty = min{Qty_i : i > p}（当前行之后所有未处理行的最小需求）
'      最后一行（p=n 或之后无行）：isLastRow=True，仅 T=D 才有效
'   3. 对 {ZP, QC, NG} 逐一检查：
'      a. 若该QC在 triedQCs 中 → 跳过
'      b. T = QueryQCTotal(ledger, ...)；若不满足可用规则 → 跳过
'      c. 获取五元组行列表，过滤掉 CurrentQty=0 的行
'   4. 两遍扫描：先统计总行数，再填充返回数组
'
' 参数：
'   currentLineKey - 复合键（由 MakePlanLineKey 构造，格式 "WMSOrderNo:LineNo"）
'   plan           - BuildStaticPlan 返回的 StaticPlan 对象
'   ledger         - 当前库存账本（实时状态，已反映本 SKU 组前几行的扣减）
'   triedQCs       - 本行已尝试但失败的QC列表（回溯时由 M09 传入）
' 返回：CandidateRow 数组；若无候选则返回未初始化空数组（调用方用 SafeCount 判断）
'
' 注意：本函数不依赖 BuildStaticPlan 的任何缓存状态，只使用 plan 参数中的排序信息。
Public Function FilterCandidatePool(ByVal currentLineKey As String, _
                                     ByVal plan As Object, _
                                     ByVal ledger As Object, _
                                     ByRef triedQCs() As String) As CandidateRow()
    ' noResult：未初始化数组，用于"无候选"时的返回值
    ' 说明：VBA 的 Empty 是保留关键字，不可用作变量名
    Dim noResult() As CandidateRow

    If plan Is Nothing Then
        FilterCandidatePool = noResult
        Exit Function
    End If
    If Not plan.Exists("RowCount") Then
        FilterCandidatePool = noResult
        Exit Function
    End If

    Dim n As Long
    n = CLng(plan("RowCount"))

    ' --- 步骤1：定位当前行在排序计划中的位置 ---
    Dim currentPos As Long
    currentPos = 0
    Dim i As Long
    For i = 1 To n
        If CStr(plan("LineKey_" & i)) = currentLineKey Then
            currentPos = i
            Exit For
        End If
    Next i

    If currentPos = 0 Then
        ' 找不到当前行（调用参数有误），安全返回空
        FilterCandidatePool = noResult
        Exit Function
    End If

    Dim D As Long
    D = CLng(plan("Qty_" & currentPos))

    Dim shipNo As String
    Dim sku As String
    shipNo = CStr(plan("ShipmentNo"))
    sku = CStr(plan("SKU"))

    ' --- 步骤2：动态计算 nextMinQty ---
    ' 当前行之后（排序中位置 > currentPos）的行 = 尚未被分配的行
    ' 与 BuildStaticPlan 使用相同的"剩余未处理行最小值"规则
    Dim hasNext As Boolean
    Dim nextMinQty As Long
    hasNext = False
    nextMinQty = 0

    For i = currentPos + 1 To n
        Dim qAfter As Long
        qAfter = CLng(plan("Qty_" & i))
        If Not hasNext Or qAfter < nextMinQty Then
            nextMinQty = qAfter
            hasNext = True
        End If
    Next i

    Dim isLastRow As Boolean
    isLastRow = Not hasNext

    ' --- 步骤3/4：两遍扫描，统计并填充候选行 ---
    Dim qcList(1 To M07_QC_COUNT) As String
    qcList(1) = QC_ZP
    qcList(2) = QC_QC
    qcList(3) = QC_NG

    Dim total As Long
    total = 0
    Dim q As Integer
    Dim T As Long
    Dim tupleRows() As InventoryRow
    Dim rc As Long
    Dim r As Long

    ' 第一遍：统计有效候选行数（CurrentQty > 0）
    For q = 1 To M07_QC_COUNT
        If Not M07_IsInTriedQCs(qcList(q), triedQCs) Then
            T = QueryQCTotal(ledger, shipNo, sku, qcList(q))
            If M07_IsQCAvailableByRule(T, D, nextMinQty, isLastRow) Then
                tupleRows = GetFiveTupleRows(ledger, shipNo, sku, qcList(q))
                rc = M07_SafeInventoryRowCount(tupleRows)
                For r = 1 To rc
                    If tupleRows(LBound(tupleRows) + r - 1).CurrentQty > 0 Then
                        total = total + 1
                    End If
                Next r
            End If
        End If
    Next q

    If total = 0 Then
        FilterCandidatePool = noResult
        Exit Function
    End If

    ' 第二遍：填充 CandidateRow 数组
    Dim result() As CandidateRow
    ReDim result(1 To total)
    Dim idx As Long
    idx = 1

    For q = 1 To M07_QC_COUNT
        If Not M07_IsInTriedQCs(qcList(q), triedQCs) Then
            T = QueryQCTotal(ledger, shipNo, sku, qcList(q))
            If M07_IsQCAvailableByRule(T, D, nextMinQty, isLastRow) Then
                tupleRows = GetFiveTupleRows(ledger, shipNo, sku, qcList(q))
                rc = M07_SafeInventoryRowCount(tupleRows)
                For r = 1 To rc
                    Dim invRow As InventoryRow
                    invRow = tupleRows(LBound(tupleRows) + r - 1)
                    If invRow.CurrentQty > 0 Then
                        result(idx).ShipmentNo  = invRow.ShipmentNo
                        result(idx).SKU         = invRow.SKU
                        result(idx).QC          = invRow.QC
                        result(idx).LotNo       = invRow.LotNo
                        result(idx).Expiry      = invRow.Expiry
                        result(idx).OriginalQty = invRow.OriginalQty
                        result(idx).CurrentQty  = invRow.CurrentQty
                        idx = idx + 1
                    End If
                Next r
            End If
        End If
    Next q

    FilterCandidatePool = result
End Function


' MakePlanLineKey：构建 FilterCandidatePool 所需的行唯一标识符（复合键）
'
' 背景：同一物流单号+SKU下，不同WMS退单号可能有相同的行号（如00001），
' 单独用行号无法唯一标识一行。使用"WMSOrderNo:LineNo"作为复合键可唯一定位。
' M09 在调用 FilterCandidatePool 之前调用此函数构造 currentLineKey。
'
' 分隔符选用冒号（:）：业务字段（WMSOrderNo 和 LineNo）中均不含冒号。
Public Function MakePlanLineKey(ByVal wmsOrderNo As String, ByVal lineNo As String) As String
    MakePlanLineKey = wmsOrderNo & ":" & lineNo
End Function


' =============================================================================
' 二、私有辅助函数（前缀 M07_ 避免与其他模块同名函数冲突）
' =============================================================================

' QC可用判定规则（§4.2.2 R072）：BuildStaticPlan 和 FilterCandidatePool 均调用此函数，
' 确保"nextMinQty 定义一致"的测试点成立（相同的判定逻辑，参数来源不同）。
'
'   最后行（isLastRow=True）：仅 T = D（库存恰好等于需求，精确耗尽）
'   其他行：T = D（精确匹配）或 T >= D + nextMinQty（分配后为后续行至少保留 nextMinQty）
'
' T <= 0 时直接判为不可用（不参与筛选）。
Private Function M07_IsQCAvailableByRule(ByVal T As Long, ByVal D As Long, _
                                          ByVal nextMinQty As Long, _
                                          ByVal isLastRow As Boolean) As Boolean
    If T <= 0 Then
        M07_IsQCAvailableByRule = False
        Exit Function
    End If

    If isLastRow Then
        ' 最后一行：库存必须恰好等于需求，不能多也不能少
        M07_IsQCAvailableByRule = (T = D)
    Else
        ' 非最后行：精确匹配，或扣除需求后仍能覆盖后续行的最小需求
        M07_IsQCAvailableByRule = (T = D) Or (T >= D + nextMinQty)
    End If
End Function

' 计算行 i 的初始可用QC数：遍历 {ZP, QC, NG} 统计满足可用规则的种类数
Private Function M07_CalcAvailableQCCount(ByVal shipNo As String, ByVal sku As String, _
                                           ByVal D As Long, ByVal nextMinQty As Long, _
                                           ByVal isLastRow As Boolean, _
                                           ByVal ledger As Object) As Long
    Dim cnt As Long
    cnt = 0

    Dim qcArr(1 To M07_QC_COUNT) As String
    qcArr(1) = QC_ZP
    qcArr(2) = QC_QC
    qcArr(3) = QC_NG

    Dim q As Integer
    For q = 1 To M07_QC_COUNT
        Dim T As Long
        T = QueryQCTotal(ledger, shipNo, sku, qcArr(q))
        If M07_IsQCAvailableByRule(T, D, nextMinQty, isLastRow) Then
            cnt = cnt + 1
        End If
    Next q

    M07_CalcAvailableQCCount = cnt
End Function

' 计算行 i 在原始未排序数组中的 nextMinQty：所有其他行（j≠i）需求量的最小值
' 用于 BuildStaticPlan 的初始静态计算（无行已分配，"其他行"= "全组其余行"）
Private Function M07_CalcNextMinQtyAll(ByVal i As Long, ByVal n As Long, _
                                        ByRef qtys() As Long) As Long
    Dim minVal As Long
    Dim hasOther As Boolean
    hasOther = False
    Dim j As Long
    For j = 1 To n
        If j <> i Then
            If Not hasOther Or qtys(j) < minVal Then
                minVal = qtys(j)
                hasOther = True
            End If
        End If
    Next j

    If hasOther Then
        M07_CalcNextMinQtyAll = minVal
    Else
        M07_CalcNextMinQtyAll = 0  ' 只有1行时返回0（外层以 isOnlyRow 标记处理）
    End If
End Function

' 从已排序的 StaticPlan 中取第 pos 行的 nextMinQty（其他所有行的最小值）
' 用于 RunPrecheck B：验证锁定行时需要重算 nextMinQty，与 BuildStaticPlan 保持一致
Private Function M07_CalcNextMinQtyInPlan(ByVal pos As Long, ByVal n As Long, _
                                           ByVal plan As Object) As Long
    Dim minVal As Long
    Dim hasOther As Boolean
    hasOther = False
    Dim i As Long
    For i = 1 To n
        If i <> pos Then
            Dim qty As Long
            qty = CLng(plan("Qty_" & i))
            If Not hasOther Or qty < minVal Then
                minVal = qty
                hasOther = True
            End If
        End If
    Next i

    If hasOther Then
        M07_CalcNextMinQtyInPlan = minVal
    Else
        M07_CalcNextMinQtyInPlan = 0
    End If
End Function

' 排序比较：行 ai 是否应排在行 bi 之后（返回 True 则交换位置）
' 四级规则：
'   1. initQCCount 升序——选择少的行优先，防止被后续行"抢走"唯一可用QC
'   2. Qty 降序——大需求行先处理，尽早占用大块库存，减少碎片风险
'   3. WMSOrderNo 升序——保证相同前两级时处理顺序可复现（字典序）
'   4. LineNo 升序——最终兜底，完全确定顺序（字典序）
Private Function M07_ShouldSwap(ByVal ai As Long, ByVal bi As Long, _
                                  ByRef initQCC() As Long, ByRef qtys() As Long, _
                                  ByRef wmsONos() As String, _
                                  ByRef lineNos() As String) As Boolean
    If initQCC(ai) <> initQCC(bi) Then
        M07_ShouldSwap = (initQCC(ai) > initQCC(bi))   ' 升序：大值排后
        Exit Function
    End If

    If qtys(ai) <> qtys(bi) Then
        M07_ShouldSwap = (qtys(ai) < qtys(bi))          ' 降序：小值排后
        Exit Function
    End If

    If wmsONos(ai) <> wmsONos(bi) Then
        M07_ShouldSwap = (wmsONos(ai) > wmsONos(bi))   ' 升序：大值排后
        Exit Function
    End If

    M07_ShouldSwap = (lineNos(ai) > lineNos(bi))        ' 升序：大值排后
End Function

' 判断指定QC是否在已尝试列表中（大小写敏感，业务数据已标准化为大写）
Private Function M07_IsInTriedQCs(ByVal qc As String, ByRef triedQCs() As String) As Boolean
    Dim cnt As Long
    cnt = M07_SafeStringArrCount(triedQCs)
    If cnt = 0 Then
        M07_IsInTriedQCs = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(triedQCs) To UBound(triedQCs)
        If triedQCs(i) = qc Then
            M07_IsInTriedQCs = True
            Exit Function
        End If
    Next i

    M07_IsInTriedQCs = False
End Function

' 安全统计 NormalizedReturnLine 数组元素数（未初始化时返回0，不抛错）
Private Function M07_SafeReturnLineCount(ByRef arr() As NormalizedReturnLine) As Long
    On Error GoTo NotInit
    M07_SafeReturnLineCount = UBound(arr) - LBound(arr) + 1
    Exit Function
NotInit:
    M07_SafeReturnLineCount = 0
End Function

' 安全统计 InventoryRow 数组元素数（未初始化时返回0，不抛错）
Private Function M07_SafeInventoryRowCount(ByRef arr() As InventoryRow) As Long
    On Error GoTo NotInit
    M07_SafeInventoryRowCount = UBound(arr) - LBound(arr) + 1
    Exit Function
NotInit:
    M07_SafeInventoryRowCount = 0
End Function

' 安全统计 String 数组元素数（未初始化时返回0，不抛错）
Private Function M07_SafeStringArrCount(ByRef arr() As String) As Long
    On Error GoTo NotInit
    M07_SafeStringArrCount = UBound(arr) - LBound(arr) + 1
    Exit Function
NotInit:
    M07_SafeStringArrCount = 0
End Function
