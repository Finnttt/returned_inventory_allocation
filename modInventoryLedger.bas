Option Explicit

' =============================================================================
' M06_库存账本（modInventoryLedger）
' =============================================================================
' 职责：按五元组（物流单号+SKU+QC+批号+效期）汇总质检库存，
'       提供统一的库存状态管理接口，是整个系统中库存状态的唯一持有者。
'
' 给新手的解释：
'   "五元组"就是确定一批货的5个维度——同一物流单号下，
'   SKU相同、QC等级相同、批号相同、效期相同的货才算"同一格库存"。
'   账本（Ledger）记录每一格库存的原始数量和当前可用数量。
'
' 内部数据结构：
'   InventoryLedger = Scripting.Dictionary
'     Key:   五元组拼接字符串（字段间用 vbNullChar 分隔，该字符不会出现在业务数据中）
'     Value: Variant 数组 Array(原始数量, 当前可用数量)
'
'   InventorySnapshot = Scripting.Dictionary（结构与账本相同，但只含指定范围的副本）
'
'   UndoLog = VBA Collection（每次 Deduct 追加一条记录，Undo 时按 LIFO 顺序还原）
'
' 设计原因：VBA 的 Type 不支持内嵌动态数组或 Dictionary 引用，
' 因此账本和快照在运行时均为 Object（Scripting.Dictionary）。
' 其他模块通过本模块的公开函数操作账本，无需知道内部分隔符等细节。
' =============================================================================

' 五元组键的分隔符：使用 vbNullChar（Chr(0)，ASCII 空字符）。
' 原因：VBA 的 Const 不支持 Chr() 函数调用，只能引用内置常量或字面值。
' vbNullChar 不会出现在任何业务字段（物流单号、SKU、QC、批号、效期）中，
' 且所有 VBA 字符串函数（Split/Left$/Len 等）均可正确处理它。
Private Const KEY_SEP  As String = vbNullChar
Private Const IDX_ORIG As Integer = 0   ' Value 数组中"原始数量"的下标
Private Const IDX_CURR As Integer = 1   ' Value 数组中"当前可用数量"的下标

' =============================================================================
' 一、核心公开函数
' =============================================================================

' 建立库存账本：将 NormalizedInventoryLine 数组按五元组汇总。
' 只有三个合法性标记（QCValid/ExpiryValid/QtyValid）全部为 True 的行才计入账本。
' 非法行已由 M05 报错，这里只纳入"可用库存"。
Public Function BuildLedger(ByRef inventory() As NormalizedInventoryLine) As Object
    Dim ledger As Object
    Set ledger = CreateObject("Scripting.Dictionary")

    ' 防御：数组未初始化时直接返回空账本
    If SafeInventoryLineCount(inventory) = 0 Then
        Set BuildLedger = ledger
        Exit Function
    End If

    Dim i As Long
    For i = LBound(inventory) To UBound(inventory)
        ' 三个合法性标记任一为 False，跳过该行
        If Not (inventory(i).QCValid And inventory(i).ExpiryValid And inventory(i).QtyValid) Then
            GoTo NextRow
        End If

        Dim k As String
        k = MakeLedgerKey(inventory(i).ShipmentNo, inventory(i).SKU, _
                           inventory(i).QC, inventory(i).LotNo, inventory(i).Expiry)

        If ledger.Exists(k) Then
            ' 相同五元组：累加数量（同一批货可能分多行录入）
            Dim arr As Variant
            arr = ledger(k)
            arr(IDX_ORIG) = arr(IDX_ORIG) + CLng(inventory(i).Qty)
            arr(IDX_CURR) = arr(IDX_CURR) + CLng(inventory(i).Qty)
            ' 注意：VBA Dictionary 的 Value 是值类型副本，必须显式写回才能生效
            ledger(k) = arr
        Else
            ledger.Add k, Array(CLng(inventory(i).Qty), CLng(inventory(i).Qty))
        End If

NextRow:
    Next i

    Set BuildLedger = ledger
End Function

' 查询特定（物流单号, SKU, QC）三元组下的当前可用总量。
' 该三元组下可能有多个批号/效期（不同五元组），本函数将它们全部加总。
' M07/M08 用此值判断某 QC 是否有足够库存满足退单需求。
Public Function QueryQCTotal(ByVal ledger As Object, _
                              ByVal shipNo As String, _
                              ByVal sku As String, _
                              ByVal qc As String) As Long
    Dim total As Long
    total = 0

    ' 前缀匹配：只要 key 以"物流单号|SKU|QC|"开头，就属于该三元组
    Dim prefix As String
    prefix = shipNo & KEY_SEP & sku & KEY_SEP & qc & KEY_SEP

    Dim key As Variant
    For Each key In ledger.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            Dim arr As Variant
            arr = ledger(key)
            total = total + arr(IDX_CURR)
        End If
    Next key

    QueryQCTotal = total
End Function

' 获取特定（物流单号, SKU, QC）下所有五元组行（含当前可用数量）。
' 返回 InventoryRow 数组；候选池为空时返回未初始化数组（安全，调用方用 SafeCount 判断）。
' M07/M08 用此结果构建候选批号/效期列表，每行对应一个可供选择的库存格。
Public Function GetFiveTupleRows(ByVal ledger As Object, _
                                  ByVal shipNo As String, _
                                  ByVal sku As String, _
                                  ByVal qc As String) As InventoryRow()
    Dim prefix As String
    prefix = shipNo & KEY_SEP & sku & KEY_SEP & qc & KEY_SEP

    ' 第一遍：统计匹配行数（VBA 动态数组需预先知道大小）
    Dim matchCount As Long
    matchCount = 0
    Dim key As Variant
    For Each key In ledger.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            matchCount = matchCount + 1
        End If
    Next key

    If matchCount = 0 Then
        Dim emptyResult() As InventoryRow
        GetFiveTupleRows = emptyResult
        Exit Function
    End If

    ' 第二遍：填充结果数组
    Dim result() As InventoryRow
    ReDim result(1 To matchCount)
    Dim idx As Long
    idx = 1

    For Each key In ledger.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            ' 解析键字符串还原各字段（顺序与 MakeLedgerKey 一致）
            Dim parts() As String
            parts = Split(CStr(key), KEY_SEP)
            result(idx).ShipmentNo = parts(0)
            result(idx).SKU        = parts(1)
            result(idx).QC         = parts(2)
            result(idx).LotNo      = parts(3)
            result(idx).Expiry     = parts(4)

            Dim arr As Variant
            arr = ledger(key)
            result(idx).OriginalQty = arr(IDX_ORIG)
            result(idx).CurrentQty  = arr(IDX_CURR)
            idx = idx + 1
        End If
    Next key

    GetFiveTupleRows = result
End Function

' 从账本中扣减指定五元组的库存数量。
' 成功：返回 True，账本 CurrentQty 减少，undoLog 追加一条还原记录。
' 失败（库存不足或 key 不存在）：返回 False，账本和 undoLog 均不变。
' 这是库存守恒的关键入口——所有分配操作都必须通过此函数修改账本。
Public Function Deduct(ByVal ledger As Object, _
                        ByRef key As InventoryKey, _
                        ByVal qty As Long, _
                        ByVal undoLog As Object) As Boolean
    Dim k As String
    k = MakeLedgerKey(key.ShipmentNo, key.SKU, key.QC, key.LotNo, key.Expiry)

    If Not ledger.Exists(k) Then
        ' key 不存在，直接拒绝
        Deduct = False
        Exit Function
    End If

    Dim arr As Variant
    arr = ledger(k)

    If arr(IDX_CURR) < qty Then
        ' 可用数量不足，拒绝扣减，账本保持原状（防止超发）
        Deduct = False
        Exit Function
    End If

    ' 执行扣减并写回 Dictionary（值类型副本修改后必须显式赋回）
    arr(IDX_CURR) = arr(IDX_CURR) - qty
    ledger(k) = arr

    ' 向撤销日志追加此次扣减的记录，格式：五元组键 & vbNullChar & 扣减数量
    ' 这样 Undo 函数只需要这一条字符串就能完整还原
    undoLog.Add k & KEY_SEP & CStr(qty)

    Deduct = True
End Function

' 撤销 undoLog 中记录的所有扣减操作。
' 按 LIFO（后进先出）顺序还原，确保与扣减顺序完全对称。
' 撤销完成后清空 undoLog，防止被重复调用导致数量异常增加。
Public Sub Undo(ByVal ledger As Object, ByVal undoLog As Object)
    ' 从最后一条向前撤销（LIFO 顺序）
    Dim i As Long
    For i = undoLog.Count To 1 Step -1
        Dim entry As String
        entry = CStr(undoLog(i))

        ' 解析 entry：格式为"shipNo|SKU|QC|LotNo|Expiry|qty"（共6段）
        Dim parts() As String
        parts = Split(entry, KEY_SEP)

        ' 重建键（前5段）和还原数量（第6段）
        Dim keyStr As String
        keyStr = parts(0) & KEY_SEP & parts(1) & KEY_SEP & parts(2) & _
                 KEY_SEP & parts(3) & KEY_SEP & parts(4)
        Dim undoQty As Long
        undoQty = CLng(parts(5))

        If ledger.Exists(keyStr) Then
            Dim arr As Variant
            arr = ledger(keyStr)
            arr(IDX_CURR) = arr(IDX_CURR) + undoQty
            ledger(keyStr) = arr
        End If
    Next i

    ' 清空撤销日志（从末尾逐条删除，避免索引偏移）
    Do While undoLog.Count > 0
        undoLog.Remove undoLog.Count
    Loop
End Sub

' 快照：将指定（物流单号, SKU）范围内所有五元组的当前状态复制为独立副本。
' 快照是深拷贝——账本后续的任何 Deduct/Undo 都不会影响快照内容。
' M10 守卫在分配一个 SKU 组之前先取快照，分配完成后用快照验证守恒等式。
Public Function TakeSnapshot(ByVal ledger As Object, _
                               ByVal shipNo As String, _
                               ByVal sku As String) As Object
    Dim snap As Object
    Set snap = CreateObject("Scripting.Dictionary")

    ' 前缀匹配：匹配该物流单号+SKU 下的所有五元组
    Dim prefix As String
    prefix = shipNo & KEY_SEP & sku & KEY_SEP

    Dim key As Variant
    For Each key In ledger.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            Dim arr As Variant
            arr = ledger(key)
            ' Array(...) 创建新的 Variant 数组，与账本中的数组完全独立
            snap.Add key, Array(arr(IDX_ORIG), arr(IDX_CURR))
        End If
    Next key

    Set TakeSnapshot = snap
End Function

' =============================================================================
' 二、辅助公开函数（供 M10 守卫和测试调用）
' =============================================================================

' 创建新的空撤销日志（VBA Collection）。
' M08/M09 在每次分配尝试前调用此函数，不需要知道日志的内部实现。
Public Function NewUndoLog() As Object
    Set NewUndoLog = New Collection
End Function

' 读取账本中指定五元组的当前可用数量。
' M10 守卫遍历快照键时，用此函数取得分配后的账本状态，与快照对比验证守恒。
Public Function GetCurrentQty(ByVal ledger As Object, ByRef key As InventoryKey) As Long
    Dim k As String
    k = MakeLedgerKey(key.ShipmentNo, key.SKU, key.QC, key.LotNo, key.Expiry)

    If Not ledger.Exists(k) Then
        GetCurrentQty = 0
        Exit Function
    End If

    Dim arr As Variant
    arr = ledger(k)
    GetCurrentQty = arr(IDX_CURR)
End Function

' 读取账本中指定五元组的原始数量（建账本时写入，此后不变）。
' M10 守卫用此验证守恒等式：原始量 = 当前量 + 本次分配量。
Public Function GetOriginalQty(ByVal ledger As Object, ByRef key As InventoryKey) As Long
    Dim k As String
    k = MakeLedgerKey(key.ShipmentNo, key.SKU, key.QC, key.LotNo, key.Expiry)

    If Not ledger.Exists(k) Then
        GetOriginalQty = 0
        Exit Function
    End If

    Dim arr As Variant
    arr = ledger(k)
    GetOriginalQty = arr(IDX_ORIG)
End Function

' 读取快照中指定五元组的当前可用数量（即分配前的可用量）。
' 与 GetCurrentQty（读账本）配合使用，可以计算本次分配消耗了多少库存。
Public Function GetSnapshotCurrentQty(ByVal snap As Object, ByRef key As InventoryKey) As Long
    Dim k As String
    k = MakeLedgerKey(key.ShipmentNo, key.SKU, key.QC, key.LotNo, key.Expiry)

    If Not snap.Exists(k) Then
        GetSnapshotCurrentQty = 0
        Exit Function
    End If

    Dim arr As Variant
    arr = snap(k)
    GetSnapshotCurrentQty = arr(IDX_CURR)
End Function

' 获取账本（或快照）中所有五元组的键字符串列表，供 M10 守卫遍历守恒验证使用。
Public Function GetLedgerKeys(ByVal ledger As Object) As String()
    Dim keys() As String
    Dim count As Long
    count = ledger.Count

    If count = 0 Then
        GetLedgerKeys = keys
        Exit Function
    End If

    ReDim keys(1 To count)
    Dim i As Long
    i = 1
    Dim k As Variant
    For Each k In ledger.Keys
        keys(i) = CStr(k)
        i = i + 1
    Next k
    GetLedgerKeys = keys
End Function

' =============================================================================
' 三、私有辅助函数
' =============================================================================

' 将五元组字段拼接为账本键字符串。
' 字段顺序固定为：物流单号 | SKU | QC | 批号 | 效期。
' 使用 vbNullChar 分隔，确保不与任何业务字段值冲突。
Private Function MakeLedgerKey(ByVal shipNo As String, ByVal sku As String, _
                                 ByVal qc As String, ByVal lotNo As String, _
                                 ByVal expiry As String) As String
    MakeLedgerKey = shipNo & KEY_SEP & sku & KEY_SEP & qc & KEY_SEP & lotNo & KEY_SEP & expiry
End Function

' 安全统计 NormalizedInventoryLine 数组元素数（未初始化时返回 0，避免 VBA 报错）。
Private Function SafeInventoryLineCount(ByRef arr() As NormalizedInventoryLine) As Long
    On Error GoTo NotInit
    SafeInventoryLineCount = UBound(arr) - LBound(arr) + 1
    Exit Function
NotInit:
    SafeInventoryLineCount = 0
End Function
