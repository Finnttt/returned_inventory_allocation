Option Explicit

' =============================================================================
' M10_工程守卫（modGuards）
' =============================================================================
' 职责：提供运行时强制断言，检验库存账目数学等式成立。
'       任一断言失败时可通过 RaiseE99 立即终止当前 SKU 组的分配。
'
' 给新手的解释：
'   这里的"守卫"类似程序里的"安全检查站"。
'   每次完成一组分配后，守卫会验证：
'     扣减前的库存总量 == 当前剩余总量 + 已分配总量
'   如果这个等式不成立，说明程序记账出错了（有 bug），
'   立即报 E99 并停止，防止错误传播到最终输出。
'
' 公开函数：
'   AssertConservation(snapshot, ledger, details()) → Boolean
'   AssertUndoConsistency(snapshot, choiceStack, ledger) → Boolean
'   RaiseE99(shipNo, sku, expected, actual, context)
'
' 与 M09 的协作关系：
'   M09 的 AllocateSKUGroup 在所有行分配成功后调用 AssertConservation；
'   在 E09/E10 回滚后调用 AssertUndoConsistency 确认账本已还原。
'   两函数返回 False 时，M09 调用 RaiseE99 终止并标记 E99。
'
' 内部实现说明：
'   snapshot 和 ledger 均为 Scripting.Dictionary（Object），
'   内部 Value 格式为 Array(OriginalQty, CurrentQty)，与 M06 保持一致。
'   本模块不需要知道 KEY_SEP 分隔符，只迭代快照已有的键。
' =============================================================================

' E99 使用的 VBA 自定义错误号（基于 vbObjectError 偏移，确保不与系统错误号冲突）
Public Const E99_ERROR_NUMBER As Long = vbObjectError + 99


' =============================================================================
' 一、断言：库存守恒（AssertConservation）
' =============================================================================

' 验证"分配前快照总量 == 账本当前剩余 + 本次已分配"的全局守恒等式。
'
' 快照（snapshot）在该 SKU 组分配开始前由 TakeSnapshot 拍摄，
' 只覆盖该物流单号 + SKU 范围内的所有五元组。
'
' 守恒等式（全局求和版）：
'   ∑ snapshot[key].CurrentQty  ==  ∑ ledger[key].CurrentQty  +  ∑ details[i].AllocQty
'
' 为何用全局求和而非逐 key 匹配：
'   账本的 key 格式（vbNullChar 分隔）是 M06 内部实现细节，本模块不暴露。
'   全局求和只需迭代快照已有的键，与账本通过同一键集合对齐，等价于逐 key 检验。
'
' 参数：
'   snapshot - TakeSnapshot 返回的快照字典，value = Array(OrigQty, CurrQty_before)
'   ledger   - 分配后的当前账本字典，value = Array(OrigQty, CurrQty_after)
'   details  - AllocationDetail 数组，记录本次 SKU 组共分配了哪些数量
'
' 返回：True = 守恒成立；False = 守恒被破坏（漏记或多记）
Public Function AssertConservation(ByVal snapshot As Object, _
                                    ByVal ledger As Object, _
                                    ByRef details() As AllocationDetail) As Boolean

    ' 快照或账本为空时无法核验，不能视为通过。
    ' 这里返回 False，让上层统一升级为 E99，避免内部账本缺失被静默放过。
    If snapshot Is Nothing Or ledger Is Nothing Then
        AssertConservation = False
        Exit Function
    End If

    ' --- 步骤1：从快照中累加"分配前总可用量" ---
    ' snapshot(key) = Array(OriginalQty, CurrentQty_at_snapshot_time)
    ' index=1 是快照时刻的 CurrentQty
    Dim totalBefore As Long
    totalBefore = 0
    Dim key As Variant
    For Each key In snapshot.Keys
        Dim snapArr As Variant
        snapArr = snapshot(key)
        totalBefore = totalBefore + CLng(snapArr(1))
    Next key

    ' --- 步骤2：从账本中累加"分配后剩余量"（只统计快照范围内的键）---
    ' ledger(key) = Array(OriginalQty, CurrentQty_now)
    ' index=1 是当前可用数量
    Dim totalAfter As Long
    totalAfter = 0
    For Each key In snapshot.Keys
        If ledger.Exists(key) Then
            Dim ledgerArr As Variant
            ledgerArr = ledger(key)
            totalAfter = totalAfter + CLng(ledgerArr(1))
        End If
    Next key

    ' --- 步骤3：累加本次分配明细的总分配量 ---
    Dim totalAllocated As Long
    totalAllocated = 0
    Dim n As Long
    n = GD_SafeDetailCount(details)
    Dim i As Long
    For i = 1 To n
        totalAllocated = totalAllocated + details(LBound(details) + i - 1).AllocQty
    Next i

    ' --- 步骤4：验证守恒等式 ---
    AssertConservation = (totalBefore = totalAfter + totalAllocated)
End Function


' =============================================================================
' 二、断言：撤销一致性（AssertUndoConsistency）
' =============================================================================

' 验证执行完所有 Undo 后，账本与快照完全一致（账本已完整还原）。
'
' 在 AllocateSKUGroup 因 E09/E10 触发回滚后调用，
' 确认 Undo 操作没有遗漏任何已扣减的库存行。
'
' 参数：
'   snapshot    - 分配前的快照（Object）
'   choiceStack - 备用参数，预留供未来验证栈是否已清空；当前传 Nothing 也可工作
'   ledger      - 回滚后的当前账本（Object）
'
' 返回：True = 账本与快照一致（Undo 完整）；False = 有差异（Undo 不完整）
Public Function AssertUndoConsistency(ByVal snapshot As Object, _
                                       ByVal choiceStack As Object, _
                                       ByVal ledger As Object) As Boolean

    ' 快照或账本为空时无法证明 Undo 已完整还原，按失败处理。
    If snapshot Is Nothing Or ledger Is Nothing Then
        AssertUndoConsistency = False
        Exit Function
    End If

    ' 逐条比对：快照中每个五元组的 CurrentQty 应与账本完全相同
    Dim key As Variant
    For Each key In snapshot.Keys
        ' 账本中找不到该键本身就是异常
        If Not ledger.Exists(key) Then
            AssertUndoConsistency = False
            Exit Function
        End If

        Dim snapArr As Variant
        Dim ledgerArr As Variant
        snapArr   = snapshot(key)
        ledgerArr = ledger(key)

        ' snapshot(key)(1) = 快照时的 CurrentQty（Undo 后应恢复到该值）
        ' ledger(key)(1)   = 当前账本的 CurrentQty
        If CLng(snapArr(1)) <> CLng(ledgerArr(1)) Then
            AssertUndoConsistency = False
            Exit Function
        End If
    Next key

    AssertUndoConsistency = True
End Function


' =============================================================================
' 三、触发 E99（RaiseE99）
' =============================================================================

' 触发 E99 工程异常，用 Err.Raise 中断调用链。
'
' 调用方（M09）在捕获到 AssertConservation=False 或 AssertUndoConsistency=False 后
' 调用本函数。M15 在最外层用 On Error GoTo 统一捕获 E99。
'
' 错误消息格式（便于运维人员快速定位）：
'   [E99] 库存守恒异常：物流单号=XXX SKU=XXX 期望=N 实际=N 上下文=XXX
'
' 参数：
'   shipNo   - 发生异常的物流单号
'   sku      - 发生异常的 SKU
'   expected - 守恒等式左侧期望值（快照总量）
'   actual   - 守恒等式右侧实际值（账本剩余 + 已分配）
'   context  - 描述发生在哪个检查点，如 "AssertConservation" 或 "AssertUndoConsistency"
Public Sub RaiseE99(ByVal shipNo As String, ByVal sku As String, _
                     ByVal expected As Long, ByVal actual As Long, _
                     ByVal context As String)

    Dim msg As String
    msg = "[E99] 库存守恒异常：" & _
          "物流单号=" & shipNo & _
          " SKU=" & sku & _
          " 期望=" & CStr(expected) & _
          " 实际=" & CStr(actual) & _
          " 上下文=" & context

    ' 同时写入即时窗口，便于开发阶段调试
    Debug.Print msg

    ' 用自定义错误号抛出，调用方用 On Error Goto / Err.Number 捕获
    Err.Raise E99_ERROR_NUMBER, "modGuards.RaiseE99", msg
End Sub


' =============================================================================
' 四、私有辅助函数
' =============================================================================

' 安全统计 AllocationDetail 数组元素数（未初始化时返回 0，不抛错）
Private Function GD_SafeDetailCount(ByRef details() As AllocationDetail) As Long
    On Error GoTo NotInit
    GD_SafeDetailCount = UBound(details) - LBound(details) + 1
    Exit Function
NotInit:
    GD_SafeDetailCount = 0
End Function


' =============================================================================
' 五、独立自检入口（不依赖 modTestRunner，用于排查编译/导入问题）
' =============================================================================

' 在 VBA 即时窗口输入 RunGuardsSelfTest 并回车即可运行。
' 若本过程能跑通，说明 modGuards + modTypes + modInventoryLedger 已就绪；
' 若 RunGuardsTests 仍失败，问题多半在 modTestRunner 未更新或工程里有重复模块。
Public Sub RunGuardsSelfTest()
    Dim passCount As Long
    Dim failCount As Long
    passCount = 0
    failCount = 0

    ' --- 用例1：守恒成立 ---
    Dim shipNo As String
    Dim sku As String
    shipNo = "SF_GD_SELF01"
    sku = "H_GD_SELF01"

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1).ShipmentNo = shipNo
    inv(1).SKU = sku
    inv(1).QC = QC_ZP
    inv(1).LotNo = "LA01"
    inv(1).Expiry = "2029/01/01"
    inv(1).Qty = 10
    inv(1).QCValid = True
    inv(1).ExpiryValid = True
    inv(1).QtyValid = True

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim snapshot As Object
    Set snapshot = TakeSnapshot(ledger, shipNo, sku)

    Dim key As InventoryKey
    key.ShipmentNo = shipNo
    key.SKU = sku
    key.QC = QC_ZP
    key.LotNo = "LA01"
    key.Expiry = "2029/01/01"

    Dim undoLog As Object
    Set undoLog = NewUndoLog()
    If Not Deduct(ledger, key, 5, undoLog) Then
        Debug.Print "[FAIL] RunGuardsSelfTest 扣减失败"
        failCount = failCount + 1
    Else
        Dim details(1 To 1) As AllocationDetail
        details(1).AllocQty = 5
        If AssertConservation(snapshot, ledger, details) Then
            Debug.Print "[PASS] 守恒成立"
            passCount = passCount + 1
        Else
            Debug.Print "[FAIL] 守恒应成立却失败"
            failCount = failCount + 1
        End If
    End If

    ' --- 用例2：RaiseE99 消息格式 ---
    On Error Resume Next
    RaiseE99 "SF_GD_SELF02", "H_GD_SELF02", 100, 90, "自检"
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum = E99_ERROR_NUMBER _
       And InStr(errDesc, "SF_GD_SELF02") > 0 _
       And InStr(errDesc, "H_GD_SELF02") > 0 _
       And InStr(errDesc, "100") > 0 _
       And InStr(errDesc, "90") > 0 Then
        Debug.Print "[PASS] RaiseE99 消息格式正确"
        passCount = passCount + 1
    Else
        Debug.Print "[FAIL] RaiseE99 消息格式不正确，Err=" & errNum & " Desc=" & errDesc
        failCount = failCount + 1
    End If

    Debug.Print "RunGuardsSelfTest 完成：PASS=" & passCount & " FAIL=" & failCount
End Sub
