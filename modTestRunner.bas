Option Explicit

' =============================================================================
' M16_测试入口（modTestRunner）
' =============================================================================
' 本模块提供 M02～M15 的内存单元/集成测试、内存端到端测试及断言汇总。
' RunAllTests 不读取外部 DataSet；阶段 A 文件集成和批量测试器冒烟
' 分别通过 RunSingleTest(17) 与 RunSingleTest(18) 单独执行。
'
' 给新手的解释：
' “断言”就是让程序自动检查“实际结果”和“预期结果”是否一致。
' 如果一致，记为 PASS；如果不一致，记为 FAIL，并打印失败原因。
' =============================================================================

Private m_TotalCount As Long
Private m_PassCount As Long
Private m_FailCount As Long
Private m_CurrentSuite As String

' 失败明细缓存：供无头验收（PowerShell COM）在运行后读取具体失败项，
' 否则失败详情只进 VBA 即时窗口，隐藏 Excel 下无法诊断。
Private m_FailLog As String

' RunAllTests 专用：静默模式下不弹单套件弹窗，全部完成后输出一次总汇总
Private m_SilentMode  As Boolean
Private m_GlobalTotal As Long
Private m_GlobalPass  As Long
Private m_GlobalFail  As Long

' -----------------------------------------------------------------------------
' M16 公开接口：RunAllTests / RunSingleTest
' 规格对应：§M16（测试入口），这是本模块规格要求的两个公开函数
' -----------------------------------------------------------------------------

' 执行全部测试套件（共 17 个；不含需真实文件的阶段A），只在最后显示一次总汇总 MsgBox。
' 设计意图：以往各套件各弹一次窗口很烦人；RunAllTests 切换到静默模式，
'   把各套件的通过/失败计数累加到全局变量，全部执行完再统一输出。
' 注意：RunStageATests（编号17）依赖真实测试文件，RunAllTests 不包含它，
'   需要时请单独调用 RunSingleTest(17)。
Public Sub RunAllTests()
    RunAllTestsCore True
End Sub

' 自动化入口：运行与 RunAllTests 相同的 17 个内存套件，但不弹 MsgBox。
' PowerShell/Excel COM 可调用此函数并读取返回文本，避免隐藏 Excel 被弹窗卡住。
Public Function RunAllTestsSilent() As String
    RunAllTestsSilent = RunAllTestsCore(False)
End Function

Public Function RunSingleTestSilent(ByVal caseNo As Integer) As String
    m_SilentMode = True
    m_GlobalTotal = 0
    m_GlobalPass = 0
    m_GlobalFail = 0
    m_FailLog = vbNullString

    RunSingleTest caseNo

    m_SilentMode = False
    RunSingleTestSilent = "编号=" & CStr(caseNo) & _
                          "，总数=" & CStr(m_GlobalTotal) & _
                          "，通过=" & CStr(m_GlobalPass) & _
                          "，失败=" & CStr(m_GlobalFail)
End Function

Private Function RunAllTestsCore(ByVal showMessage As Boolean) As String
    m_SilentMode  = True
    m_GlobalTotal = 0
    m_GlobalPass  = 0
    m_GlobalFail  = 0
    m_FailLog     = vbNullString

    Debug.Print String(70, "*")
    Debug.Print "*** RunAllTests 开始（共 17 个套件）"
    Debug.Print String(70, "*")

    ' --- 框架自检 ---
    RunSmokeTests

    ' --- 基础/输入层（M02~M05）---
    RunConfigTests
    RunExcelInputTests
    RunNormalizeTests
    RunValidateTests

    ' --- 核心算法层（M06~M10）---
    RunLedgerTests
    RunSortFilterTests
    RunStrategyTests
    RunBacktrackingTests
    RunGuardsTests

    ' --- 结果/输出层（M11~M15）---
    RunStatusTests
    RunPostValidateTests
    RunOutputBuilderTests
    RunExcelOutputTests
    RunM13M14IntegrationTests
    RunRunnerTests

    ' --- E2E 验收（M05→M06→M07→M09→M11 完整链路，T01~T14）---
    RunAcceptanceTests

    ' 恢复普通模式，输出最终汇总
    m_SilentMode = False

    Debug.Print String(70, "*")
    Debug.Print "*** RunAllTests 完成：总数=" & m_GlobalTotal & _
                "，通过=" & m_GlobalPass & _
                "，失败=" & m_GlobalFail
    Debug.Print String(70, "*")

    RunAllTestsCore = "总数=" & CStr(m_GlobalTotal) & _
                      "，通过=" & CStr(m_GlobalPass) & _
                      "，失败=" & CStr(m_GlobalFail)

    If Not showMessage Then Exit Function

    If m_GlobalFail > 0 Then
        MsgBox "RunAllTests 完成，存在 " & m_GlobalFail & " 个失败用例。" & vbNewLine & _
               "请查看 VBA 即时窗口（Ctrl+G）。", vbExclamation
    Else
        MsgBox "RunAllTests 全部通过，共 " & m_GlobalPass & " 个用例。", vbInformation
    End If
End Function

' 按编号运行单个测试套件，方便定向调试。
' 编号说明：
'   0 = 冒烟测试        1 = M02 配置       2 = M03 读取
'   3 = M04 标准化      4 = M05 校验       5 = M06 账本
'   6 = M07 排序筛选    7 = M08 策略       8 = M09 回溯
'   9 = M10 守卫        10 = M11 状态      11 = M12 后校验
'   12 = M13 输出构建   13 = M14 写入      14 = M13+M14 集成
'   15 = M15 编排       16 = E2E 验收      17 = 阶段A集成（需真实文件）
'   18 = 批量测试运行器冒烟（内存临时工作簿）
Public Sub RunSingleTest(ByVal caseNo As Integer)
    Select Case caseNo
        Case 0:  RunSmokeTests
        Case 1:  RunConfigTests
        Case 2:  RunExcelInputTests
        Case 3:  RunNormalizeTests
        Case 4:  RunValidateTests
        Case 5:  RunLedgerTests
        Case 6:  RunSortFilterTests
        Case 7:  RunStrategyTests
        Case 8:  RunBacktrackingTests
        Case 9:  RunGuardsTests
        Case 10: RunStatusTests
        Case 11: RunPostValidateTests
        Case 12: RunOutputBuilderTests
        Case 13: RunExcelOutputTests
        Case 14: RunM13M14IntegrationTests
        Case 15: RunRunnerTests
        Case 16: RunAcceptanceTests
        Case 17: RunStageATests
        Case 18: RunBatchRunnerSmokeTest
        Case Else:
            MsgBox "无效编号 " & caseNo & "，有效范围：0~18。" & vbNewLine & _
                   "（编号 17 = 阶段A集成测试，需要真实测试文件；" & vbNewLine & _
                   " 编号 18 = 批量测试运行器冒烟测试）", vbExclamation
    End Select
End Sub

' -----------------------------------------------------------------------------
' 公开入口：测试框架自检
' -----------------------------------------------------------------------------

Public Sub RunSmokeTests()
    ' Smoke Test（冒烟测试）用于确认测试器本身可以正常工作。
    ' 这里故意只放会通过的断言，避免在项目刚开始时制造误报。
    BeginSuite "TestRunner Smoke Tests"

    AssertEqualString "字符串断言应通过", "ABC", "ABC"
    AssertEqualLong "数字断言应通过", 200, 200
    AssertEqualBool "布尔断言应通过", True, True

    FinishSuite
End Sub

Public Sub RunConfigTests()
    ' M02 配置模块测试。
    ' 测试器会临时创建一张工作表，跑完后自动删除，不会依赖正式输入表。
    BeginSuite "M02 Config Tests"

    Dim ws As Worksheet
    Set ws = CreateTempConfigSheet()

    On Error GoTo CleanFail

    WriteConfigHeaders ws
    WriteConfigRow ws, 2, "SF3190000000001", "TC-02", 200, DEBUG_LEVEL_OFF, LOT_MODE_INSENSITIVE, DEFAULT_NO_EXPIRY_SENTINEL
    WriteConfigRow ws, 3, "SF3190000000028", "TC-24", 10, DEBUG_LEVEL_SIMPLE, LOT_MODE_INSENSITIVE, DEFAULT_NO_EXPIRY_SENTINEL

    Dim cfg As ConfigStruct

    cfg = LoadConfig(ws)
    AssertEqualLong "LoadConfig 读取第2行最大回溯次数", 200, cfg.MaxBacktrackCount
    AssertEqualString "LoadConfig 读取第2行调试日志级别", DEBUG_LEVEL_OFF, cfg.DebugLogLevel
    AssertEqualBool "LoadConfig 批号不敏感转为 False", False, cfg.LotCaseSensitive
    AssertEqualLong "缺少详细日志单表上限时走默认值", DEFAULT_DETAILED_LOG_LIMIT, cfg.DetailedLogLimit

    cfg = LoadConfigForShipment(ws, "SF3190000000028")
    AssertEqualLong "按物流单号读取最大回溯次数", 10, cfg.MaxBacktrackCount
    AssertEqualString "按物流单号读取调试日志级别", DEBUG_LEVEL_SIMPLE, cfg.DebugLogLevel

    Dim wsKv As Worksheet
    Set wsKv = CreateTempSheet("__tmp_config_kv_test")
    WriteKeyValueConfig wsKv
    cfg = LoadConfig(wsKv)
    AssertEqualLong "键值配置读取最大回溯次数", 88, cfg.MaxBacktrackCount
    AssertEqualString "键值配置读取调试日志级别", DEBUG_LEVEL_DETAIL, cfg.DebugLogLevel
    AssertEqualLong "键值配置读取详细日志单表上限", 12345, cfg.DetailedLogLimit

CleanExit:
    DeleteTempSheet wsKv
    DeleteTempSheet ws
    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M02 Config Tests 执行异常：" & Err.Description
    DeleteTempSheet wsKv
    DeleteTempSheet ws
    FinishSuite
End Sub

Public Sub RunExcelInputTests()
    ' M03 数据加载模块测试。
    ' 重点验证：按固定列读取原始值，以及效期单元格类型 ExpiryCellKind。
    BeginSuite "M03 Excel Input Tests"

    Dim returnWs As Worksheet
    Dim inventoryWs As Worksheet

    Set returnWs = CreateTempSheet("__tmp_return_input_test")
    Set inventoryWs = CreateTempSheet("__tmp_inventory_input_test")

    On Error GoTo CleanFail

    WriteReturnInputFixture returnWs
    WriteInventoryInputFixture inventoryWs

    Dim orders() As RawReturnRow
    orders = ReadReturnOrders(returnWs)

    AssertEqualLong "退单表读取行数", 2, UBound(orders) - LBound(orders) + 1
    AssertEqualLong "退单表原始 Excel 行号", 2, orders(1).ExcelRowNum
    AssertEqualString "退单表物流单号列映射", "SF3190000000001", CStr(orders(1).ShipmentNo)
    AssertEqualString "退单表 WMS 退单号列映射", "TK00000011", CStr(orders(1).WMSOrderNo)
    AssertEqualString "退单表行号保留文本", "00001", CStr(orders(1).LineNo)

    Dim inventory() As RawInventoryRow
    inventory = ReadQCInventory(inventoryWs)

    AssertEqualLong "库存表读取行数", 3, UBound(inventory) - LBound(inventory) + 1
    AssertEqualString "Excel 日期型效期识别", CELL_KIND_EXCEL_DATE, inventory(1).ExpiryCellKind
    AssertEqualString "文本效期识别", CELL_KIND_TEXT, inventory(2).ExpiryCellKind
    AssertEqualString "空效期识别", CELL_KIND_BLANK, inventory(3).ExpiryCellKind

CleanExit:
    DeleteTempSheet returnWs
    DeleteTempSheet inventoryWs
    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M03 Excel Input Tests 执行异常：" & Err.Description
    DeleteTempSheet returnWs
    DeleteTempSheet inventoryWs
    FinishSuite
End Sub

Public Sub RunNormalizeTests()
    ' M04 数据标准化模块测试。
    BeginSuite "M04 Normalize Tests"

    On Error GoTo CleanFail

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim rawOrders(1 To 2) As RawReturnRow
    rawOrders(1).ExcelRowNum = 2
    rawOrders(1).ShipmentNo = " SF3190000000048 "
    rawOrders(1).WMSOrderNo = " TK10000480 "
    rawOrders(1).SKU = " H000000048 "
    rawOrders(1).LineNo = "00001"
    rawOrders(1).Qty = 5

    rawOrders(2).ExcelRowNum = 3
    rawOrders(2).ShipmentNo = "SF3190000000048"
    rawOrders(2).WMSOrderNo = "TK10000480"
    rawOrders(2).SKU = "H000000048"
    rawOrders(2).LineNo = 1          ' 数值型行号经 CStr 后为 "1"，应判非法
    rawOrders(2).Qty = "abc"         ' 非数字数量，应判非法

    Dim orderIssues() As FieldNormalizeIssue
    Dim orders() As NormalizedReturnLine
    orders = NormalizeReturnRows(rawOrders, cfg, orderIssues)

    AssertEqualString "退单物流单号 Trim", "SF3190000000048", orders(1).ShipmentNo
    AssertEqualBool "文本五位行号合法", True, orders(1).LineNoValid
    AssertEqualBool "数值型行号非法", False, orders(2).LineNoValid
    AssertEqualBool "非数字数量非法", False, orders(2).QtyValid
    AssertEqualLong "退单标准化问题数", 2, CountFieldIssues(orderIssues)

    Dim rawInventory(1 To 3) As RawInventoryRow
    rawInventory(1).ExcelRowNum = 2
    rawInventory(1).ShipmentNo = "SF3190000000046"
    rawInventory(1).SKU = "H000000046"
    rawInventory(1).QC = " zp "
    rawInventory(1).LotNo = " a01 "
    rawInventory(1).Expiry = DateSerial(2029, 6, 15)
    rawInventory(1).ExpiryCellKind = CELL_KIND_EXCEL_DATE
    rawInventory(1).Qty = 6

    rawInventory(2).ExcelRowNum = 3
    rawInventory(2).ShipmentNo = "SF3190000000049"
    rawInventory(2).SKU = "H000000049"
    rawInventory(2).QC = "QC"
    rawInventory(2).LotNo = "BA03"
    rawInventory(2).Expiry = "2029-12-31"
    rawInventory(2).ExpiryCellKind = CELL_KIND_TEXT
    rawInventory(2).Qty = 5

    rawInventory(3).ExcelRowNum = 4
    rawInventory(3).ShipmentNo = "SF3190000000049"
    rawInventory(3).SKU = "H000000049"
    rawInventory(3).QC = "NG"
    rawInventory(3).LotNo = "BA04"
    rawInventory(3).Expiry = "2029/02/29"    ' 2029 非闰年，2月29日非法
    rawInventory(3).ExpiryCellKind = CELL_KIND_TEXT
    rawInventory(3).Qty = 0                  ' 非正整数非法

    Dim inventoryIssues() As FieldNormalizeIssue
    Dim inventory() As NormalizedInventoryLine
    inventory = NormalizeInventoryRows(rawInventory, cfg, inventoryIssues)

    AssertEqualString "QC Trim+UCase", "ZP", inventory(1).QC
    AssertEqualString "批号默认不敏感转大写", "A01", inventory(1).LotNo
    AssertEqualString "ExcelDate 效期标准化", "2029/06/15", inventory(1).Expiry
    AssertEqualString "文本连字符效期标准化", "2029/12/31", inventory(2).Expiry
    AssertEqualBool "非闰年 2月29日非法", False, inventory(3).ExpiryValid
    AssertEqualBool "数量 0 非法", False, inventory(3).QtyValid
    AssertEqualLong "库存标准化问题数", 2, CountFieldIssues(inventoryIssues)

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M04 Normalize Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

Public Sub RunValidateTests()
    ' M05 分配前校验模块测试。
    ' 覆盖 E01~E08、E11 的主要路径，并复用 SF0051/SF0036 真实工作簿做集成验证。
    BeginSuite "M05 Validate Tests"

    On Error GoTo CleanFail

    TestValidateE01EmptyField
    TestValidateE02DuplicateLineNo
    TestValidateE02DiscontinuousLineNo
    TestValidateE03InvalidQc
    TestValidateE04InvalidQty
    TestValidateE05InvalidExpiry
    TestValidateE06ShipmentOnlyInOrders
    TestValidateE07ShipmentOnlyInInventory
    TestValidateE08QtyMismatch
    TestValidateE11FragmentInventory
    TestValidateMultiErrorRetention

    RunValidateWorkbook_SF0051
    RunValidateWorkbook_SF0036

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M05 Validate Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

Public Sub RunStageATests()
    ' 阶段 A 一键入口：
    ' 1) 真实文件读取链路
    ' 2) 异常路径最小集（临时构造）
    ' 3) 异常路径真实文件（SF0052/SF0053/SF0054/SF0055）
    RunIntegrationReadTests
    RunInputErrorTests
    RunInputErrorWorkbookTests
End Sub

Public Sub RunIntegrationReadTests()
    ' 真实文件集成测试：
    ' 1) 测试用例部分汇总.xlsm：按物流单号读取配置 + 读取输入 + 标准化
    ' 2) SF0032/SF0046/SF0047/SF0048/SF0051：逐个读取+标准化
    BeginSuite "Stage A Integration Read Tests"

    On Error GoTo CleanFail

    Dim rootPath As String
    rootPath = ResolveProjectRoot()

    Dim summaryPath As String
    summaryPath = rootPath & "\测试用例部分汇总.xlsm"
    AssertTrue "汇总工作簿存在", FileExists(summaryPath)

    Dim wbSummary As Workbook
    Dim summaryOpenedByTest As Boolean
    If StrComp(ThisWorkbook.FullName, summaryPath, vbTextCompare) = 0 Then
        ' 测试正由汇总工作簿自身承载时直接复用，避免重复打开同一文件触发隐藏对话框。
        Set wbSummary = ThisWorkbook
        summaryOpenedByTest = False
    Else
        Set wbSummary = Workbooks.Open(summaryPath, False, True)
        summaryOpenedByTest = True
    End If

    Dim wsCfg As Worksheet
    Dim wsReturn As Worksheet
    Dim wsInv As Worksheet
    Set wsCfg = wbSummary.Worksheets("输入_配置")
    Set wsReturn = wbSummary.Worksheets("输入_退单表")
    Set wsInv = wbSummary.Worksheets("输入_质检库存表")

    Dim cfg As ConfigStruct
    cfg = LoadConfigForShipment(wsCfg, "SF3190000000028")
    AssertEqualLong "汇总表按物流单号读取最大回溯次数", 10, cfg.MaxBacktrackCount
    AssertEqualString "汇总表按物流单号读取调试日志级别", DEBUG_LEVEL_SIMPLE, cfg.DebugLogLevel

    Dim rawOrders() As RawReturnRow
    Dim rawInventory() As RawInventoryRow
    rawOrders = ReadReturnOrders(wsReturn)
    rawInventory = ReadQCInventory(wsInv)

    AssertTrue "汇总表退单行数>0", CountRawReturnRows(rawOrders) > 0
    AssertTrue "汇总表库存行数>0", CountRawInventoryRows(rawInventory) > 0
    AssertEqualString "汇总表库存首行效期类型为ExcelDate", CELL_KIND_EXCEL_DATE, rawInventory(1).ExpiryCellKind

    Dim issuesA() As FieldNormalizeIssue
    Dim issuesB() As FieldNormalizeIssue
    Dim normOrders() As NormalizedReturnLine
    Dim normInv() As NormalizedInventoryLine
    normOrders = NormalizeReturnRows(rawOrders, cfg, issuesA)
    normInv = NormalizeInventoryRows(rawInventory, cfg, issuesB)

    AssertEqualLong "汇总表退单标准化行数一致", CountRawReturnRows(rawOrders), CountNormalizedReturnRows(normOrders)
    AssertEqualLong "汇总表库存标准化行数一致", CountRawInventoryRows(rawInventory), CountNormalizedInventoryRows(normInv)

    If summaryOpenedByTest Then wbSummary.Close False
    Set wbSummary = Nothing

    ' SF 文件逐个验证（读取 + 标准化 + 核心字段）
    RunSingleSfWorkbookReadTest rootPath, "SF0032_测试数据.xlsx", 2, 1, "SF3190000000032"
    RunSingleSfWorkbookReadTest rootPath, "SF0046_测试数据.xlsx", 2, 2, "SF3190000000046"
    RunSingleSfWorkbookReadTest rootPath, "SF0047_测试数据.xlsx", 2, 2, "SF3190000000047"
    RunSingleSfWorkbookReadTest rootPath, "SF0048_测试数据.xlsx", 3, 1, "SF3190000000048"
    RunSingleSfWorkbookReadTest rootPath, "SF0051_测试数据.xlsx", 4, 1, "SF3190000000051"

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] Stage A Integration Read Tests 执行异常：" & Err.Description
    On Error Resume Next
    If summaryOpenedByTest And Not wbSummary Is Nothing Then wbSummary.Close False
    On Error GoTo 0
    FinishSuite
End Sub

Public Sub RunInputErrorTests()
    ' 异常路径最小集：
    ' 1) 表头错位
    ' 2) 缺列
    ' 3) 非法配置值（调试日志级别）
    ' 4) 非法配置值（最大回溯次数）
    BeginSuite "Stage A Input Error Tests"

    On Error GoTo CleanFail

    TestHeaderMismatchError
    TestMissingColumnError
    TestWmsOrderCrossShipmentError
    TestInvalidConfigDebugLevelError
    TestInvalidConfigMaxBacktrackError

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] Stage A Input Error Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

Public Sub RunInputErrorWorkbookTests()
    ' 异常路径真实文件测试：
    ' - SF0052：E12 表头错位
    ' - SF0053：E12 缺列
    ' - SF0054A：最大回溯次数非法
    ' - SF0054B：调试日志级别非法
    BeginSuite "Stage A Input Error Workbook Tests"

    On Error GoTo CleanFail

    Dim rootPath As String
    rootPath = ResolveProjectRoot()

    RunSingleSfErrorWorkbookTest rootPath, "SF0052_测试数据.xlsx", "ReadReturnOrders", "E12"
    RunSingleSfErrorWorkbookTest rootPath, "SF0053_测试数据.xlsx", "ReadReturnOrders", "E12"
    RunSingleSfErrorWorkbookTest rootPath, "SF0054A_测试数据.xlsx", "LoadConfig", "配置读取错误"
    RunSingleSfErrorWorkbookTest rootPath, "SF0054B_测试数据.xlsx", "LoadConfig", "配置读取错误"
    RunSingleSfErrorWorkbookTest rootPath, "SF0055_测试数据.xlsx", "ReadReturnOrders", "E12"

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] Stage A Input Error Workbook Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' -----------------------------------------------------------------------------
' 测试生命周期
' -----------------------------------------------------------------------------

Public Sub BeginSuite(ByVal suiteName As String)
    m_TotalCount = 0
    m_PassCount = 0
    m_FailCount = 0
    m_CurrentSuite = suiteName

    Debug.Print String(70, "=")
    Debug.Print "开始测试：" & m_CurrentSuite
    Debug.Print String(70, "=")
End Sub

Public Sub FinishSuite()
    ' 各测试套件的 CleanFail 都会调用 FinishSuite。
    ' 若此时 Err 仍有值，说明套件发生了未预期运行错误；必须计为失败，
    ' 不能只写即时窗口后仍显示“全部通过”。
    If Err.Number <> 0 Then
        Dim unexpectedNumber As Long
        Dim unexpectedDescription As String
        unexpectedNumber = Err.Number
        unexpectedDescription = Err.Description
        Err.Clear

        m_TotalCount = m_TotalCount + 1
        RecordFail "未预期运行错误", _
                   "错误号=" & CStr(unexpectedNumber) & "；" & unexpectedDescription
    End If

    Debug.Print String(70, "-")
    Debug.Print "测试完成：" & m_CurrentSuite
    Debug.Print "总数=" & m_TotalCount & _
                "，通过=" & m_PassCount & _
                "，失败=" & m_FailCount
    Debug.Print String(70, "=")

    ' RunAllTests 启用静默模式时：将计数累加到全局变量，不显示单套件弹窗。
    ' 普通模式（单独调用某套件时）：保持原来的弹窗行为，不受影响。
    If m_SilentMode Then
        m_GlobalTotal = m_GlobalTotal + m_TotalCount
        m_GlobalPass  = m_GlobalPass  + m_PassCount
        m_GlobalFail  = m_GlobalFail  + m_FailCount
    Else
        If m_FailCount > 0 Then
            MsgBox "测试完成，但存在失败用例：" & m_FailCount & " 个。请查看 VBA 即时窗口。", vbExclamation
        Else
            MsgBox "测试全部通过：" & m_PassCount & " 个。", vbInformation
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' 断言工具
' -----------------------------------------------------------------------------

Public Sub AssertEqualString(ByVal caseName As String, ByVal expected As String, ByVal actual As String)
    m_TotalCount = m_TotalCount + 1

    If expected = actual Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期=[" & expected & "]，实际=[" & actual & "]"
    End If
End Sub

Public Sub AssertEqualLong(ByVal caseName As String, ByVal expected As Long, ByVal actual As Long)
    m_TotalCount = m_TotalCount + 1

    If expected = actual Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期=[" & CStr(expected) & "]，实际=[" & CStr(actual) & "]"
    End If
End Sub

Public Sub AssertEqualBool(ByVal caseName As String, ByVal expected As Boolean, ByVal actual As Boolean)
    m_TotalCount = m_TotalCount + 1

    If expected = actual Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期=[" & BoolToText(expected) & "]，实际=[" & BoolToText(actual) & "]"
    End If
End Sub

Public Sub AssertTrue(ByVal caseName As String, ByVal condition As Boolean)
    m_TotalCount = m_TotalCount + 1

    If condition Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期条件为 True，但实际为 False"
    End If
End Sub

Public Sub AssertFalse(ByVal caseName As String, ByVal condition As Boolean)
    m_TotalCount = m_TotalCount + 1

    If Not condition Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期条件为 False，但实际为 True"
    End If
End Sub

' -----------------------------------------------------------------------------
' 私有输出函数
' -----------------------------------------------------------------------------

Private Sub RecordPass(ByVal caseName As String)
    m_PassCount = m_PassCount + 1
    Debug.Print "[PASS] " & caseName
End Sub

Private Sub RecordFail(ByVal caseName As String, ByVal detail As String)
    m_FailCount = m_FailCount + 1
    m_FailLog = m_FailLog & "[FAIL] " & caseName & "；" & detail & vbLf
    Debug.Print "[FAIL] " & caseName & "；" & detail
End Sub

' 无头诊断：返回本次运行累计的失败明细（每次 RunAllTests/RunSingleTestSilent 开始时清零）。
Public Function GetFailLog() As String
    GetFailLog = m_FailLog
End Function

Private Function BoolToText(ByVal value As Boolean) As String
    If value Then
        BoolToText = "True"
    Else
        BoolToText = "False"
    End If
End Function

Private Sub RunSingleSfWorkbookReadTest( _
    ByVal rootPath As String, _
    ByVal fileName As String, _
    ByVal expectedReturnRows As Long, _
    ByVal expectedInventoryRows As Long, _
    ByVal expectedShipmentNo As String)

    Dim fullPath As String
    fullPath = rootPath & "\" & fileName
    AssertTrue fileName & " 存在", FileExists(fullPath)

    Dim wb As Workbook
    Set wb = Workbooks.Open(fullPath, False, True)

    Dim wsCfg As Worksheet
    Dim wsReturn As Worksheet
    Dim wsInv As Worksheet
    Set wsCfg = wb.Worksheets("输入_配置")
    Set wsReturn = wb.Worksheets("输入_退单表")
    Set wsInv = wb.Worksheets("输入_质检库存表")

    Dim cfg As ConfigStruct
    cfg = LoadConfig(wsCfg)

    Dim rawOrders() As RawReturnRow
    Dim rawInv() As RawInventoryRow
    rawOrders = ReadReturnOrders(wsReturn)
    rawInv = ReadQCInventory(wsInv)

    AssertEqualLong fileName & " 退单读取行数", expectedReturnRows, CountRawReturnRows(rawOrders)
    AssertEqualLong fileName & " 库存读取行数", expectedInventoryRows, CountRawInventoryRows(rawInv)
    AssertEqualString fileName & " 退单首行物流单号", expectedShipmentNo, CStr(rawOrders(1).ShipmentNo)
    AssertTrue fileName & " 库存首行效期类型合法", _
               (rawInv(1).ExpiryCellKind = CELL_KIND_EXCEL_DATE Or rawInv(1).ExpiryCellKind = CELL_KIND_TEXT)

    Dim issuesA() As FieldNormalizeIssue
    Dim issuesB() As FieldNormalizeIssue
    Dim normOrders() As NormalizedReturnLine
    Dim normInv() As NormalizedInventoryLine
    normOrders = NormalizeReturnRows(rawOrders, cfg, issuesA)
    normInv = NormalizeInventoryRows(rawInv, cfg, issuesB)

    AssertEqualLong fileName & " 退单标准化行数一致", CountRawReturnRows(rawOrders), CountNormalizedReturnRows(normOrders)
    AssertEqualLong fileName & " 库存标准化行数一致", CountRawInventoryRows(rawInv), CountNormalizedInventoryRows(normInv)

    wb.Close False
End Sub

Private Sub RunSingleSfErrorWorkbookTest( _
    ByVal rootPath As String, _
    ByVal fileName As String, _
    ByVal invokeAction As String, _
    ByVal expectedErrorLabel As String)

    Dim fullPath As String
    fullPath = rootPath & "\" & fileName
    AssertTrue fileName & " 存在", FileExists(fullPath)

    Dim wb As Workbook
    Set wb = Workbooks.Open(fullPath, False, True)

    Dim wsReturn As Worksheet
    Dim wsCfg As Worksheet
    Dim wsExpected As Worksheet
    Set wsReturn = wb.Worksheets("输入_退单表")
    Set wsCfg = wb.Worksheets("输入_配置")
    Set wsExpected = wb.Worksheets("预期_断言")

    Dim raised As Boolean
    Dim errorText As String
    raised = False

    On Error Resume Next
    Select Case invokeAction
        Case "ReadReturnOrders"
            Dim orders() As RawReturnRow
            orders = ReadReturnOrders(wsReturn)
        Case "LoadConfig"
            Dim cfg As ConfigStruct
            cfg = LoadConfig(wsCfg)
        Case Else
            Err.Raise vbObjectError + 5101, "modTestRunner", "未知测试动作：" & invokeAction
    End Select

    If Err.Number <> 0 Then
        raised = True
        errorText = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue fileName & " 应抛错", raised

    If raised Then
        AssertTrue fileName & " 错误信息非空", Len(Trim$(errorText)) > 0
        AssertErrorTextAgainstExpected fileName, errorText, wsExpected
    End If

    AssertTrue fileName & " 预期错误标签存在", HasExpectedLabel(wsExpected, expectedErrorLabel)

    wb.Close False
End Sub

Private Sub AssertErrorTextAgainstExpected( _
    ByVal casePrefix As String, _
    ByVal errorText As String, _
    ByVal wsExpected As Worksheet)

    Dim lastRow As Long
    lastRow = wsExpected.Cells(wsExpected.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim keyName As String
        Dim expectedSnippet As String
        keyName = Trim$(CStr(wsExpected.Cells(r, 1).Value))
        expectedSnippet = Trim$(CStr(wsExpected.Cells(r, 2).Value))

        If expectedSnippet <> vbNullString And InStr(1, keyName, "关键报错片段", vbTextCompare) > 0 Then
            AssertContains casePrefix & " " & keyName & " 命中", errorText, expectedSnippet
        End If
    Next r
End Sub

Private Function HasExpectedLabel(ByVal wsExpected As Worksheet, ByVal expectedLabel As String) As Boolean
    Dim lastRow As Long
    lastRow = wsExpected.Cells(wsExpected.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(wsExpected.Cells(r, 1).Value)) = "期望错误" Then
            If InStr(1, Trim$(CStr(wsExpected.Cells(r, 2).Value)), expectedLabel, vbTextCompare) > 0 Then
                HasExpectedLabel = True
                Exit Function
            End If
        End If
    Next r

    HasExpectedLabel = False
End Function

Private Sub TestHeaderMismatchError()
    Dim ws As Worksheet
    Set ws = CreateTempSheet("__tmp_error_header_mismatch")
    On Error GoTo CleanFail

    ws.Cells(1, 1).Value = "物流单号X" ' 故意错
    ws.Cells(1, 2).Value = "WMS退单号"
    ws.Cells(1, 3).Value = "SKU"
    ws.Cells(1, 4).Value = "行号"
    ws.Cells(1, 5).Value = "数量"
    ws.Cells(2, 1).Value = "SF3190000000001"

    Dim raised As Boolean
    raised = False

    On Error Resume Next
    Dim rows() As RawReturnRow
    rows = ReadReturnOrders(ws)
    If Err.Number <> 0 Then
        raised = True
        AssertContains "表头错位错误信息可定位", Err.Description, "第 1 列应为 [物流单号]"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue "表头错位应抛错", raised

CleanExit:
    DeleteTempSheet ws
    Exit Sub

CleanFail:
    DeleteTempSheet ws
End Sub

Private Sub TestMissingColumnError()
    Dim ws As Worksheet
    Set ws = CreateTempSheet("__tmp_error_missing_col")
    On Error GoTo CleanFail

    ws.Cells(1, 1).Value = "物流单号"
    ws.Cells(1, 2).Value = "WMS退单号"
    ws.Cells(1, 3).Value = "SKU"
    ws.Cells(1, 4).Value = "行号"
    ' 第5列故意缺失（为空）
    ws.Cells(2, 1).Value = "SF3190000000001"

    Dim raised As Boolean
    raised = False

    On Error Resume Next
    Dim rows() As RawReturnRow
    rows = ReadReturnOrders(ws)
    If Err.Number <> 0 Then
        raised = True
        AssertContains "缺列错误信息可定位", Err.Description, "第 5 列应为 [数量]"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue "缺列应抛错", raised

CleanExit:
    DeleteTempSheet ws
    Exit Sub

CleanFail:
    DeleteTempSheet ws
End Sub

Private Sub TestInvalidConfigDebugLevelError()
    Dim ws As Worksheet
    Set ws = CreateTempConfigSheet()
    On Error GoTo CleanFail

    WriteConfigHeaders ws
    WriteConfigRow ws, 2, "SF3190000000001", "TC-XX", 200, "开启", LOT_MODE_INSENSITIVE, DEFAULT_NO_EXPIRY_SENTINEL

    Dim raised As Boolean
    raised = False

    On Error Resume Next
    Dim cfg As ConfigStruct
    cfg = LoadConfig(ws)
    If Err.Number <> 0 Then
        raised = True
        AssertContains "非法调试日志级别信息可定位", Err.Description, "调试日志级别"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue "非法调试日志级别应抛错", raised

CleanExit:
    DeleteTempSheet ws
    Exit Sub

CleanFail:
    DeleteTempSheet ws
End Sub

Private Sub TestWmsOrderCrossShipmentError()
    ' E12-②：同一 WMS退单号出现在两个不同物流单号下
    Dim ws As Worksheet
    Set ws = CreateTempSheet("__tmp_error_wms_cross_ship")
    On Error GoTo CleanFail

    ws.Cells(1, 1).Value = "物流单号"
    ws.Cells(1, 2).Value = "WMS退单号"
    ws.Cells(1, 3).Value = "SKU"
    ws.Cells(1, 4).Value = "行号"
    ws.Cells(1, 5).Value = "数量"
    ws.Cells(2, 1).Value = "SF3190000000055"
    ws.Cells(2, 2).Value = "TK10000550"
    ws.Cells(2, 3).Value = "H000000055"
    ws.Cells(2, 4).NumberFormat = "@"
    ws.Cells(2, 4).Value = "00001"
    ws.Cells(2, 5).Value = 5
    ws.Cells(3, 1).Value = "SF3190000000056"
    ws.Cells(3, 2).Value = "TK10000550"
    ws.Cells(3, 3).Value = "H000000055"
    ws.Cells(3, 4).NumberFormat = "@"
    ws.Cells(3, 4).Value = "00001"
    ws.Cells(3, 5).Value = 3

    Dim raised As Boolean
    raised = False

    On Error Resume Next
    Dim rows() As RawReturnRow
    rows = ReadReturnOrders(ws)
    If Err.Number <> 0 Then
        raised = True
        AssertContains "E12-② 错误信息含退单号", Err.Description, "TK10000550"
        AssertContains "E12-② 错误信息含已出现物流单号", Err.Description, "SF3190000000055"
        AssertContains "E12-② 错误信息含冲突物流单号", Err.Description, "SF3190000000056"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue "WMS退单号跨物流单号应抛错", raised

CleanExit:
    DeleteTempSheet ws
    Exit Sub

CleanFail:
    DeleteTempSheet ws
End Sub

Private Sub TestInvalidConfigMaxBacktrackError()
    Dim ws As Worksheet
    Set ws = CreateTempConfigSheet()
    On Error GoTo CleanFail

    WriteConfigHeaders ws
    WriteConfigRow ws, 2, "SF3190000000001", "TC-XX", -1, DEBUG_LEVEL_OFF, LOT_MODE_INSENSITIVE, DEFAULT_NO_EXPIRY_SENTINEL

    Dim raised As Boolean
    raised = False

    On Error Resume Next
    Dim cfg As ConfigStruct
    cfg = LoadConfig(ws)
    If Err.Number <> 0 Then
        raised = True
        AssertContains "非法最大回溯次数信息可定位", Err.Description, "最大回溯次数"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue "非法最大回溯次数应抛错", raised

CleanExit:
    DeleteTempSheet ws
    Exit Sub

CleanFail:
    DeleteTempSheet ws
End Sub

Private Sub AssertContains(ByVal caseName As String, ByVal actual As String, ByVal expectedPart As String)
    m_TotalCount = m_TotalCount + 1

    If InStr(1, actual, expectedPart, vbTextCompare) > 0 Then
        RecordPass caseName
    Else
        RecordFail caseName, "预期包含=[" & expectedPart & "]，实际=[" & actual & "]"
    End If
End Sub

Private Function ResolveProjectRoot() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path

    If candidate <> vbNullString Then
        If FileExists(candidate & "\测试用例部分汇总.xlsm") Then
            ResolveProjectRoot = candidate
            Exit Function
        End If
    End If

    candidate = "D:\cursor_practice\returned_inventory_allocation"
    If FileExists(candidate & "\测试用例部分汇总.xlsm") Then
        ResolveProjectRoot = candidate
        Exit Function
    End If

    Err.Raise vbObjectError + 5001, "modTestRunner", _
              "未找到测试数据目录。请将测试驱动工作簿放在 returned_inventory_allocation 目录，" & _
              "或确认 D:\cursor_practice\returned_inventory_allocation\测试用例部分汇总.xlsm 存在。"
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath)) > 0)
End Function

' -----------------------------------------------------------------------------
' M02 测试辅助函数
' -----------------------------------------------------------------------------

Private Function CreateTempConfigSheet() As Worksheet
    Set CreateTempConfigSheet = CreateTempSheet("__tmp_config_test")
End Function

Private Function CreateTempSheet(ByVal sheetName As String) As Worksheet
    DeleteSheetIfExists sheetName

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = sheetName
    Set CreateTempSheet = ws
End Function

Private Sub WriteConfigHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "物流单号"
    ws.Cells(1, 2).Value = "TC编号"
    ws.Cells(1, 3).Value = "最大回溯次数"
    ws.Cells(1, 4).Value = "调试日志级别"
    ws.Cells(1, 5).Value = "批号比较模式"
    ws.Cells(1, 6).Value = "无保质期哨兵值"
    ws.Cells(1, 7).Value = "备注"
End Sub

Private Sub WriteKeyValueConfig(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "参数名"
    ws.Cells(1, 2).Value = "值"
    ws.Cells(1, 3).Value = "说明"

    ws.Cells(2, 1).Value = "最大回溯次数"
    ws.Cells(2, 2).Value = 88
    ws.Cells(2, 3).Value = "生产全局配置"

    ws.Cells(3, 1).Value = "调试日志级别"
    ws.Cells(3, 2).Value = DEBUG_LEVEL_DETAIL

    ws.Cells(4, 1).Value = "批号比较模式"
    ws.Cells(4, 2).Value = LOT_MODE_SENSITIVE

    ws.Cells(5, 1).Value = "无保质期哨兵值"
    ws.Cells(5, 2).Value = DEFAULT_NO_EXPIRY_SENTINEL

    ws.Cells(6, 1).Value = "详细日志单表上限"
    ws.Cells(6, 2).Value = 12345
End Sub

Private Sub WriteConfigRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal shipmentNo As String, _
    ByVal tcNo As String, _
    ByVal maxBacktrack As Long, _
    ByVal debugLevel As String, _
    ByVal lotMode As String, _
    ByVal sentinel As String)

    ws.Cells(rowIndex, 1).Value = shipmentNo
    ws.Cells(rowIndex, 2).Value = tcNo
    ws.Cells(rowIndex, 3).Value = maxBacktrack
    ws.Cells(rowIndex, 4).Value = debugLevel
    ws.Cells(rowIndex, 5).Value = lotMode
    ws.Cells(rowIndex, 6).Value = sentinel
    ws.Cells(rowIndex, 7).Value = vbNullString
End Sub

Private Sub DeleteTempSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        DeleteTempSheet ws
    End If
End Sub

Private Sub WriteReturnInputFixture(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "物流单号"
    ws.Cells(1, 2).Value = "WMS退单号"
    ws.Cells(1, 3).Value = "SKU"
    ws.Cells(1, 4).Value = "行号"
    ws.Cells(1, 5).Value = "数量"

    ws.Cells(2, 1).Value = "SF3190000000001"
    ws.Cells(2, 2).Value = "TK00000011"
    ws.Cells(2, 3).Value = "H000000001"
    ws.Cells(2, 4).NumberFormat = "@"
    ws.Cells(2, 4).Value = "00001"
    ws.Cells(2, 5).Value = 2

    ws.Cells(3, 1).Value = "SF3190000000001"
    ws.Cells(3, 2).Value = "TK00000011"
    ws.Cells(3, 3).Value = "H000000001"
    ws.Cells(3, 4).NumberFormat = "@"
    ws.Cells(3, 4).Value = "00002"
    ws.Cells(3, 5).Value = 1
End Sub

Private Sub WriteInventoryInputFixture(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "物流单号"
    ws.Cells(1, 2).Value = "SKU"
    ws.Cells(1, 3).Value = "QC情况"
    ws.Cells(1, 4).Value = "批号"
    ws.Cells(1, 5).Value = "效期"
    ws.Cells(1, 6).Value = "数量"
    ws.Cells(1, 7).Value = "备注"

    ws.Cells(2, 1).Value = "SF3190000000049"
    ws.Cells(2, 2).Value = "H000000049"
    ws.Cells(2, 3).Value = "ZP"
    ws.Cells(2, 4).Value = "BA01"
    ws.Cells(2, 5).Value = DateSerial(2029, 1, 1)
    ws.Cells(2, 6).Value = 5

    ws.Cells(3, 1).Value = "SF3190000000049"
    ws.Cells(3, 2).Value = "H000000049"
    ws.Cells(3, 3).Value = "ZP"
    ws.Cells(3, 4).Value = "BA02"
    ws.Cells(3, 5).NumberFormat = "@"
    ws.Cells(3, 5).Value = "2029/06/15"
    ws.Cells(3, 6).Value = 5

    ws.Cells(4, 1).Value = "SF3190000000049"
    ws.Cells(4, 2).Value = "H000000049"
    ws.Cells(4, 3).Value = "ZP"
    ws.Cells(4, 4).Value = "BA03"
    ws.Cells(4, 5).ClearContents
    ws.Cells(4, 6).Value = 5
End Sub

Private Function CountFieldIssues(ByRef issues() As FieldNormalizeIssue) As Long
    On Error GoTo NotAllocated
    CountFieldIssues = UBound(issues) - LBound(issues) + 1
    Exit Function

NotAllocated:
    CountFieldIssues = 0
End Function

Private Function CountRawReturnRows(ByRef rows() As RawReturnRow) As Long
    On Error GoTo NotAllocated
    CountRawReturnRows = UBound(rows) - LBound(rows) + 1
    Exit Function

NotAllocated:
    CountRawReturnRows = 0
End Function

Private Function CountRawInventoryRows(ByRef rows() As RawInventoryRow) As Long
    On Error GoTo NotAllocated
    CountRawInventoryRows = UBound(rows) - LBound(rows) + 1
    Exit Function

NotAllocated:
    CountRawInventoryRows = 0
End Function

Private Function CountNormalizedReturnRows(ByRef rows() As NormalizedReturnLine) As Long
    On Error GoTo NotAllocated
    CountNormalizedReturnRows = UBound(rows) - LBound(rows) + 1
    Exit Function

NotAllocated:
    CountNormalizedReturnRows = 0
End Function

Private Function CountNormalizedInventoryRows(ByRef rows() As NormalizedInventoryLine) As Long
    On Error GoTo NotAllocated
    CountNormalizedInventoryRows = UBound(rows) - LBound(rows) + 1
    Exit Function

NotAllocated:
    CountNormalizedInventoryRows = 0
End Function

' -----------------------------------------------------------------------------
' M05 测试用例
' -----------------------------------------------------------------------------

Private Sub TestValidateE01EmptyField()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000099"
    orders(1).WMSOrderNo = "TK10000099"
    orders(1).SKU = vbNullString
    orders(1).LineNo = "00001"
    orders(1).Qty = 5
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    ' 补一条同物流单号的库存行，避免第2层 E06 干扰本用例（只测 E01）。
    ' 数量无效且不计入汇总，避免第3层 E08 误触发。
    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000099"
    inventory(1).SKU = "H000000099"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 0
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = False

    Dim fieldIssues(1 To 1) As FieldNormalizeIssue
    fieldIssues(1).ExcelRowNum = 2
    fieldIssues(1).SourceTable = SOURCE_RETURN_TABLE
    fieldIssues(1).FieldName = "SKU"
    fieldIssues(1).RawValue = vbNullString
    fieldIssues(1).IssueKind = ISSUE_KIND_EMPTY

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, fieldIssues, cfg, validationIssues)

    AssertTrue "E01 应标记失败", result.HasFailures
    AssertEqualLong "E01 问题数", 1, CountIssuesByError(validationIssues, ERR_E01)
    AssertEqualString "E01 错误码", ERR_E01, validationIssues(1).ErrorCode
End Sub

Private Sub TestValidateE02DuplicateLineNo()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000098"
    orders(1).WMSOrderNo = "TK10000098"
    orders(1).SKU = "H000000098"
    orders(1).LineNo = "00001"
    orders(1).Qty = 3
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    orders(2).ExcelRowNum = 3
    orders(2).ShipmentNo = "SF3190000000098"
    orders(2).WMSOrderNo = "TK10000098"
    orders(2).SKU = "H000000098"
    orders(2).LineNo = "00001"
    orders(2).Qty = 2
    orders(2).LineNoValid = True
    orders(2).QtyValid = True

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000098"
    inventory(1).SKU = "H000000098"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, EmptyFieldIssueArray(), cfg, validationIssues)

    AssertTrue "E02 重复行号应失败", result.HasFailures
    AssertEqualLong "E02 重复至少2条记录", 2, CountIssuesByError(validationIssues, ERR_E02)
End Sub

Private Sub TestValidateE02DiscontinuousLineNo()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000097"
    orders(1).WMSOrderNo = "TK10000097"
    orders(1).SKU = "H000000097"
    orders(1).LineNo = "00001"
    orders(1).Qty = 3
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    orders(2).ExcelRowNum = 3
    orders(2).ShipmentNo = "SF3190000000097"
    orders(2).WMSOrderNo = "TK10000097"
    orders(2).SKU = "H000000097"
    orders(2).LineNo = "00003"
    orders(2).Qty = 2
    orders(2).LineNoValid = True
    orders(2).QtyValid = True

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000097"
    inventory(1).SKU = "H000000097"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, EmptyFieldIssueArray(), cfg, validationIssues)

    AssertTrue "E02 跳号应失败", result.HasFailures
    AssertEqualLong "E02 跳号记录数", 2, CountIssuesByError(validationIssues, ERR_E02)
    AssertContains "E02 跳号原因含不连续", validationIssues(1).Reason, "不连续"
End Sub

Private Sub TestValidateE03InvalidQc()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000096"
    inventory(1).SKU = "H000000096"
    inventory(1).QC = "XX"
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = False
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim fieldIssues(1 To 1) As FieldNormalizeIssue
    fieldIssues(1).ExcelRowNum = 2
    fieldIssues(1).SourceTable = SOURCE_INVENTORY_TABLE
    fieldIssues(1).FieldName = "QC情况"
    fieldIssues(1).RawValue = "XX"
    fieldIssues(1).IssueKind = ISSUE_KIND_FORMAT_ERROR

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000096"
    orders(1).WMSOrderNo = "TK10000096"
    orders(1).SKU = "H000000096"
    orders(1).LineNo = "00001"
    orders(1).Qty = 5
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, fieldIssues, cfg, validationIssues)

    AssertTrue "E03 应失败", result.HasFailures
    AssertEqualString "E03 错误码", ERR_E03, validationIssues(1).ErrorCode
End Sub

Private Sub TestValidateE04InvalidQty()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000095"
    orders(1).WMSOrderNo = "TK10000095"
    orders(1).SKU = "H000000095"
    orders(1).LineNo = "00001"
    orders(1).Qty = 0
    orders(1).LineNoValid = True
    orders(1).QtyValid = False

    Dim fieldIssues(1 To 1) As FieldNormalizeIssue
    fieldIssues(1).ExcelRowNum = 2
    fieldIssues(1).SourceTable = SOURCE_RETURN_TABLE
    fieldIssues(1).FieldName = "数量"
    fieldIssues(1).RawValue = "0"
    fieldIssues(1).IssueKind = ISSUE_KIND_RANGE_ERROR

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000095"
    inventory(1).SKU = "H000000095"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, fieldIssues, cfg, validationIssues)

    AssertTrue "E04 应失败", result.HasFailures
    AssertEqualString "E04 错误码", ERR_E04, validationIssues(1).ErrorCode
    AssertFalse "E04 后应跳过 E08", ShipmentHasError(validationIssues, "SF3190000000095", ERR_E08)
End Sub

Private Sub TestValidateE05InvalidExpiry()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000094"
    inventory(1).SKU = "H000000094"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/13/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = False
    inventory(1).QtyValid = True

    Dim fieldIssues(1 To 1) As FieldNormalizeIssue
    fieldIssues(1).ExcelRowNum = 2
    fieldIssues(1).SourceTable = SOURCE_INVENTORY_TABLE
    fieldIssues(1).FieldName = "效期"
    fieldIssues(1).RawValue = "2029/13/01"
    fieldIssues(1).IssueKind = ISSUE_KIND_FORMAT_ERROR

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000094"
    orders(1).WMSOrderNo = "TK10000094"
    orders(1).SKU = "H000000094"
    orders(1).LineNo = "00001"
    orders(1).Qty = 5
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, fieldIssues, cfg, validationIssues)

    AssertTrue "E05 应失败", result.HasFailures
    AssertEqualString "E05 错误码", ERR_E05, validationIssues(1).ErrorCode
End Sub

Private Sub TestValidateE06ShipmentOnlyInOrders()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000093"
    orders(1).WMSOrderNo = "TK10000093"
    orders(1).SKU = "H000000093"
    orders(1).LineNo = "00001"
    orders(1).Qty = 5
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, EmptyInventoryArray(), EmptyFieldIssueArray(), cfg, validationIssues)

    AssertTrue "E06 应失败", result.HasFailures
    AssertTrue "E06 命中", ShipmentHasError(validationIssues, "SF3190000000093", ERR_E06)
    AssertFalse "E06 单侧缺失时不应叠加 E08", _
                ShipmentHasError(validationIssues, "SF3190000000093", ERR_E08)

    ' 2026-07-20 起 E06 按退单行生成明细：Excel行号/WMS/SKU 均可定位（进数据异常明细表）
    Dim foundE06 As Boolean
    Dim i As Long
    For i = LBound(validationIssues) To UBound(validationIssues)
        If validationIssues(i).ErrorCode = ERR_E06 Then
            foundE06 = True
            AssertEqualLong "E06 明细Excel行号", 2, validationIssues(i).ExcelRowNum
            AssertEqualString "E06 明细WMS退单号", "TK10000093", validationIssues(i).WMSOrderNo
            AssertEqualString "E06 明细SKU", "H000000093", validationIssues(i).SKU
            AssertEqualString "E06 明细来源表", SOURCE_RETURN_TABLE, validationIssues(i).SourceTable
        End If
    Next i
    AssertTrue "E06 生成行级明细问题记录", foundE06
End Sub

Private Sub TestValidateE07ShipmentOnlyInInventory()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000092"
    inventory(1).SKU = "H000000092"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim anomalies() As AnomalyRow
    Dim result As ValidationResult
    result = ValidatePre(EmptyReturnArray(), inventory, EmptyFieldIssueArray(), cfg, validationIssues)
    anomalies = BuildAnomalyRows(validationIssues)

    AssertTrue "E07 应失败", result.HasFailures
    AssertEqualLong "E07 异常明细行数", 1, CountAnomalyRows(anomalies)
    AssertEqualString "E07 异常错误码", ERR_E07, anomalies(1).ErrorCode
    AssertFalse "E07 单侧缺失时不应叠加 E08", _
                ShipmentHasError(validationIssues, "SF3190000000092", ERR_E08)
End Sub

Private Sub TestValidateE08QtyMismatch()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000091"
    orders(1).WMSOrderNo = "TK10000091"
    orders(1).SKU = "H000000091"
    orders(1).LineNo = "00001"
    orders(1).Qty = 8
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    Dim inventory(1 To 1) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000091"
    inventory(1).SKU = "H000000091"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 5
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, EmptyFieldIssueArray(), cfg, validationIssues)

    AssertTrue "E08 应失败", result.HasFailures
    AssertTrue "E08 命中", ShipmentHasError(validationIssues, "SF3190000000091", ERR_E08)
End Sub

Private Sub TestValidateE11FragmentInventory()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000036"
    orders(1).WMSOrderNo = "TK00000036"
    orders(1).SKU = "H000000001"
    orders(1).LineNo = "00001"
    orders(1).Qty = 2
    orders(1).LineNoValid = True
    orders(1).QtyValid = True

    Dim inventory(1 To 2) As NormalizedInventoryLine
    inventory(1).ExcelRowNum = 2
    inventory(1).ShipmentNo = "SF3190000000036"
    inventory(1).SKU = "H000000001"
    inventory(1).QC = QC_ZP
    inventory(1).LotNo = "LA01"
    inventory(1).Expiry = "2029/01/01"
    inventory(1).Qty = 1
    inventory(1).QCValid = True
    inventory(1).ExpiryValid = True
    inventory(1).QtyValid = True

    inventory(2).ExcelRowNum = 3
    inventory(2).ShipmentNo = "SF3190000000036"
    inventory(2).SKU = "H000000001"
    inventory(2).QC = QC_QC
    inventory(2).LotNo = "LA01"
    inventory(2).Expiry = "2029/01/01"
    inventory(2).Qty = 1
    inventory(2).QCValid = True
    inventory(2).ExpiryValid = True
    inventory(2).QtyValid = True

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, EmptyFieldIssueArray(), cfg, validationIssues)

    AssertTrue "E11 应失败", result.HasFailures
    AssertTrue "E11 命中", ShipmentHasError(validationIssues, "SF3190000000036", ERR_E11)
End Sub

Private Sub TestValidateMultiErrorRetention()
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1).ExcelRowNum = 2
    orders(1).ShipmentNo = "SF3190000000090"
    orders(1).WMSOrderNo = "TK10000090"
    orders(1).SKU = "H000000090"
    orders(1).LineNo = "1"
    orders(1).Qty = 0
    orders(1).LineNoValid = False
    orders(1).QtyValid = False

    Dim fieldIssues(1 To 2) As FieldNormalizeIssue
    fieldIssues(1).ExcelRowNum = 2
    fieldIssues(1).SourceTable = SOURCE_RETURN_TABLE
    fieldIssues(1).FieldName = "行号"
    fieldIssues(1).RawValue = "1"
    fieldIssues(1).IssueKind = ISSUE_KIND_FORMAT_ERROR

    fieldIssues(2).ExcelRowNum = 2
    fieldIssues(2).SourceTable = SOURCE_RETURN_TABLE
    fieldIssues(2).FieldName = "数量"
    fieldIssues(2).RawValue = "abc"
    fieldIssues(2).IssueKind = ISSUE_KIND_FORMAT_ERROR

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, EmptyInventoryArray(), fieldIssues, cfg, validationIssues)

    AssertTrue "E01+E04 应失败", result.HasFailures
    AssertEqualLong "E01 保留", 1, CountIssuesByError(validationIssues, ERR_E01)
    AssertEqualLong "E04 保留", 1, CountIssuesByError(validationIssues, ERR_E04)
End Sub

Private Sub RunValidateWorkbook_SF0051()
    Dim rootPath As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim wsReturn As Worksheet
    Dim wsInv As Worksheet
    Dim wsCfg As Worksheet

    rootPath = ResolveProjectRoot()
    fullPath = rootPath & "\SF0051_测试数据.xlsx"
    AssertTrue "SF0051 工作簿存在", FileExists(fullPath)

    Set wb = Workbooks.Open(fullPath, False, True)
    Set wsReturn = wb.Worksheets("输入_退单表")
    Set wsInv = wb.Worksheets("输入_质检库存表")
    Set wsCfg = wb.Worksheets("输入_配置")

    Dim cfg As ConfigStruct
    cfg = LoadConfig(wsCfg)

    Dim rawOrders() As RawReturnRow
    Dim rawInv() As RawInventoryRow
    rawOrders = ReadReturnOrders(wsReturn)
    rawInv = ReadQCInventory(wsInv)

    Dim fieldIssuesA() As FieldNormalizeIssue
    Dim fieldIssuesB() As FieldNormalizeIssue
    Dim orders() As NormalizedReturnLine
    Dim inventory() As NormalizedInventoryLine
    orders = NormalizeReturnRows(rawOrders, cfg, fieldIssuesA)
    inventory = NormalizeInventoryRows(rawInv, cfg, fieldIssuesB)

    Dim validationIssues() As ValidationIssue
    Dim anomalies() As AnomalyRow
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, MergeFieldIssues(fieldIssuesA, fieldIssuesB), cfg, validationIssues)
    anomalies = BuildAnomalyRows(validationIssues)

    AssertTrue "SF0051 校验应失败", result.HasFailures
    AssertTrue "SF0051 命中 E02", ShipmentHasError(validationIssues, "SF3190000000051", ERR_E02)
    AssertEqualLong "SF0051 E02 异常明细至少4行", 4, CountAnomalyRows(anomalies)

    wb.Close False
End Sub

Private Sub RunValidateWorkbook_SF0036()
    Dim rootPath As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim wsReturn As Worksheet
    Dim wsInv As Worksheet
    Dim wsCfg As Worksheet

    rootPath = ResolveProjectRoot()
    fullPath = rootPath & "\SF0036_测试数据.xlsx"
    AssertTrue "SF0036 工作簿存在", FileExists(fullPath)

    Set wb = Workbooks.Open(fullPath, False, True)
    Set wsReturn = wb.Worksheets("输入_退单表")
    Set wsInv = wb.Worksheets("输入_质检库存表")
    Set wsCfg = wb.Worksheets("输入_配置")

    Dim cfg As ConfigStruct
    cfg = LoadConfig(wsCfg)

    Dim rawOrders() As RawReturnRow
    Dim rawInv() As RawInventoryRow
    rawOrders = ReadReturnOrders(wsReturn)
    rawInv = ReadQCInventory(wsInv)

    Dim fieldIssuesA() As FieldNormalizeIssue
    Dim fieldIssuesB() As FieldNormalizeIssue
    Dim orders() As NormalizedReturnLine
    Dim inventory() As NormalizedInventoryLine
    orders = NormalizeReturnRows(rawOrders, cfg, fieldIssuesA)
    inventory = NormalizeInventoryRows(rawInv, cfg, fieldIssuesB)

    Dim validationIssues() As ValidationIssue
    Dim result As ValidationResult
    result = ValidatePre(orders, inventory, MergeFieldIssues(fieldIssuesA, fieldIssuesB), cfg, validationIssues)

    AssertTrue "SF0036 校验应失败", result.HasFailures
    AssertTrue "SF0036 命中 E11", ShipmentHasError(validationIssues, "SF3190000000036", ERR_E11)

    wb.Close False
End Sub

Private Function EmptyReturnArray() As NormalizedReturnLine()
    Dim rows() As NormalizedReturnLine
    EmptyReturnArray = rows
End Function

Private Function EmptyInventoryArray() As NormalizedInventoryLine()
    Dim rows() As NormalizedInventoryLine
    EmptyInventoryArray = rows
End Function

Private Function EmptyFieldIssueArray() As FieldNormalizeIssue()
    Dim issues() As FieldNormalizeIssue
    EmptyFieldIssueArray = issues
End Function

' 返回空的分配结果数组。注意：不可直接传给 ByRef 参数，须先赋给局部变量。
Private Function EmptyShipmentResultArray() As Object()
    Dim results() As Object
    EmptyShipmentResultArray = results
End Function

Private Function EmptyValidationIssueArray() As ValidationIssue()
    Dim issues() As ValidationIssue
    EmptyValidationIssueArray = issues
End Function

' =============================================================================
' M06 测试用例（UT-Ledger）
' =============================================================================

Public Sub RunLedgerTests()
    ' M06 库存账本模块测试。
    ' 覆盖：建账本汇总、扣减正常路径、超量拒绝、边界值、Undo 恢复、快照独立性。
    BeginSuite "M06 Inventory Ledger Tests"

    On Error GoTo CleanFail

    TestLedgerBuildSameTupleMerge
    TestLedgerBuildDifferentTuplesSeparate
    TestLedgerBuildSkipsInvalidRows
    TestLedgerDeductSuccess
    TestLedgerDeductExceedBalance
    TestLedgerDeductExactBalance
    TestLedgerDeductKeyNotFound
    TestLedgerUndoRestoresBalance
    TestLedgerUndoClearsLog
    TestLedgerSnapshotIsIsolated
    TestLedgerQueryQCTotalCrossLot
    TestLedgerGetFiveTupleRowsFilter

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M06 Ledger Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-L01：相同五元组多行应合并汇总
Private Sub TestLedgerBuildSameTupleMerge()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF001", "H001", "ZP", "BA01", "2029/01/01", 3)
    inv(2) = MakeInventoryLine("SF001", "H001", "ZP", "BA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    AssertEqualLong "TC-L01 相同五元组汇总总量=8", 8, QueryQCTotal(ledger, "SF001", "H001", "ZP")

    Dim rows() As InventoryRow
    rows = GetFiveTupleRows(ledger, "SF001", "H001", "ZP")
    AssertEqualLong "TC-L01 相同五元组只有1行", 1, SafeInventoryRowCount(rows)
    AssertEqualLong "TC-L01 合并行 OriginalQty=8", 8, rows(1).OriginalQty
    AssertEqualLong "TC-L01 合并行 CurrentQty=8", 8, rows(1).CurrentQty
End Sub

' TC-L02：不同批号应独立存储，不合并
Private Sub TestLedgerBuildDifferentTuplesSeparate()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF002", "H002", "ZP", "BA01", "2029/01/01", 3)
    inv(2) = MakeInventoryLine("SF002", "H002", "ZP", "BA02", "2029/01/01", 4)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    AssertEqualLong "TC-L02 不同批号总量=7", 7, QueryQCTotal(ledger, "SF002", "H002", "ZP")

    Dim rows() As InventoryRow
    rows = GetFiveTupleRows(ledger, "SF002", "H002", "ZP")
    AssertEqualLong "TC-L02 不同批号有2行", 2, SafeInventoryRowCount(rows)
End Sub

' TC-L03：非法行（QtyValid/QCValid 为 False）不计入账本
Private Sub TestLedgerBuildSkipsInvalidRows()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF003", "H003", "ZP", "BA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF003", "H003", "ZP", "BA02", "2029/01/01", 3)
    inv(2).QtyValid = False   ' 数量非法，应跳过
    inv(3) = MakeInventoryLine("SF003", "H003", "ZP", "BA03", "2029/01/01", 4)
    inv(3).QCValid = False    ' QC 非法，应跳过

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    AssertEqualLong "TC-L03 非法行不计入，总量=5", 5, QueryQCTotal(ledger, "SF003", "H003", "ZP")

    Dim rows() As InventoryRow
    rows = GetFiveTupleRows(ledger, "SF003", "H003", "ZP")
    AssertEqualLong "TC-L03 只有1行合法行", 1, SafeInventoryRowCount(rows)
End Sub

' TC-L04：正常扣减成功，CurrentQty 减少，OriginalQty 不变
Private Sub TestLedgerDeductSuccess()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF004", "H004", "ZP", "BA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF004", "H004", "ZP", "BA01", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    AssertTrue "TC-L04 Deduct 返回 True", Deduct(ledger, key, 3, log)
    AssertEqualLong "TC-L04 扣减后 CurrentQty=2", 2, GetCurrentQty(ledger, key)
    AssertEqualLong "TC-L04 OriginalQty 不变=5", 5, GetOriginalQty(ledger, key)
    AssertEqualLong "TC-L04 undoLog 增加1条", 1, log.Count
End Sub

' TC-L05：扣减量超过余量，应返回 False，账本和 undoLog 均不变
Private Sub TestLedgerDeductExceedBalance()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF005", "H005", "ZP", "BA01", "2029/01/01", 3)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF005", "H005", "ZP", "BA01", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    AssertFalse "TC-L05 超量扣减返回 False", Deduct(ledger, key, 5, log)
    AssertEqualLong "TC-L05 账本不变 CurrentQty=3", 3, GetCurrentQty(ledger, key)
    AssertEqualLong "TC-L05 undoLog 无新记录", 0, log.Count
End Sub

' TC-L06：精确扣尽（边界值）应成功，CurrentQty 变为 0
Private Sub TestLedgerDeductExactBalance()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF006", "H006", "ZP", "BA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF006", "H006", "ZP", "BA01", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    AssertTrue "TC-L06 精确扣尽返回 True", Deduct(ledger, key, 5, log)
    AssertEqualLong "TC-L06 扣尽后 CurrentQty=0", 0, GetCurrentQty(ledger, key)
End Sub

' TC-L07：key 不存在于账本时，Deduct 应返回 False，不改变任何状态
Private Sub TestLedgerDeductKeyNotFound()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF007", "H007", "ZP", "BA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF007", "H007_NONEXIST", "ZP", "BA01", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    AssertFalse "TC-L07 key 不存在返回 False", Deduct(ledger, key, 3, log)
    AssertEqualLong "TC-L07 undoLog 无记录", 0, log.Count
End Sub

' TC-L08：Undo 后账本完全恢复至扣减前状态（含多步扣减 LIFO 还原）
Private Sub TestLedgerUndoRestoresBalance()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF008", "H008", "ZP", "BA01", "2029/01/01", 10)
    inv(2) = MakeInventoryLine("SF008", "H008", "ZP", "BA02", "2029/01/01", 8)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key1 As InventoryKey
    Dim key2 As InventoryKey
    key1 = MakeInventoryKey("SF008", "H008", "ZP", "BA01", "2029/01/01")
    key2 = MakeInventoryKey("SF008", "H008", "ZP", "BA02", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    ' 连续两次扣减
    Deduct ledger, key1, 4, log
    Deduct ledger, key2, 3, log

    AssertEqualLong "TC-L08 扣减后 key1=6", 6, GetCurrentQty(ledger, key1)
    AssertEqualLong "TC-L08 扣减后 key2=5", 5, GetCurrentQty(ledger, key2)

    Undo ledger, log

    AssertEqualLong "TC-L08 Undo后 key1 恢复=10", 10, GetCurrentQty(ledger, key1)
    AssertEqualLong "TC-L08 Undo后 key2 恢复=8", 8, GetCurrentQty(ledger, key2)
End Sub

' TC-L09：Undo 后 undoLog 应被清空，防止二次误用
Private Sub TestLedgerUndoClearsLog()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF009", "H009", "ZP", "BA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF009", "H009", "ZP", "BA01", "2029/01/01")

    Dim log As Object
    Set log = NewUndoLog()

    Deduct ledger, key, 3, log
    AssertEqualLong "TC-L09 Undo前日志有1条", 1, log.Count

    Undo ledger, log
    AssertEqualLong "TC-L09 Undo后日志清空=0", 0, log.Count
End Sub

' TC-L10：快照独立性——取快照后扣减账本，快照保持分配前状态
Private Sub TestLedgerSnapshotIsIsolated()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF010", "H010", "ZP", "BA01", "2029/01/01", 10)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim key As InventoryKey
    key = MakeInventoryKey("SF010", "H010", "ZP", "BA01", "2029/01/01")

    ' 取快照（此时 CurrentQty=10）
    Dim snap As Object
    Set snap = TakeSnapshot(ledger, "SF010", "H010")

    ' 扣减账本
    Dim log As Object
    Set log = NewUndoLog()
    Deduct ledger, key, 6, log

    AssertEqualLong "TC-L10 扣减后账本 CurrentQty=4", 4, GetCurrentQty(ledger, key)
    AssertEqualLong "TC-L10 快照中 CurrentQty 仍=10", 10, GetSnapshotCurrentQty(snap, key)
End Sub

' TC-L11：QueryQCTotal 跨批号/效期汇总同一 QC，不同 QC 各自独立
Private Sub TestLedgerQueryQCTotalCrossLot()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF011", "H011", "ZP", "BA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF011", "H011", "ZP", "BA02", "2029/06/01", 3)
    inv(3) = MakeInventoryLine("SF011", "H011", "QC", "BA03", "2029/01/01", 4)  ' 不同QC

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    AssertEqualLong "TC-L11 ZP 跨批号总量=8", 8, QueryQCTotal(ledger, "SF011", "H011", "ZP")
    AssertEqualLong "TC-L11 QC 总量=4", 4, QueryQCTotal(ledger, "SF011", "H011", "QC")
    AssertEqualLong "TC-L11 NG 总量=0（不存在）", 0, QueryQCTotal(ledger, "SF011", "H011", "NG")
End Sub

' TC-L12：GetFiveTupleRows 只返回指定 QC 的行，不混入其他 QC
Private Sub TestLedgerGetFiveTupleRowsFilter()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF012", "H012", "ZP", "BA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF012", "H012", "ZP", "BA02", "2029/01/01", 3)
    inv(3) = MakeInventoryLine("SF012", "H012", "NG", "BA01", "2029/01/01", 2)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows() As InventoryRow
    rows = GetFiveTupleRows(ledger, "SF012", "H012", "ZP")
    AssertEqualLong "TC-L12 ZP 返回2行", 2, SafeInventoryRowCount(rows)

    rows = GetFiveTupleRows(ledger, "SF012", "H012", "NG")
    AssertEqualLong "TC-L12 NG 返回1行", 1, SafeInventoryRowCount(rows)

    rows = GetFiveTupleRows(ledger, "SF012", "H012", "QC")
    AssertEqualLong "TC-L12 QC 返回0行（空数组）", 0, SafeInventoryRowCount(rows)
End Sub

' -----------------------------------------------------------------------------
' M06 测试辅助函数
' -----------------------------------------------------------------------------

' 快速构建一个合法的 NormalizedInventoryLine（默认三个 Valid 标记均为 True）。
Private Function MakeInventoryLine( _
    ByVal shipNo As String, ByVal sku As String, ByVal qc As String, _
    ByVal lotNo As String, ByVal expiry As String, ByVal qty As Long) As NormalizedInventoryLine

    Dim line As NormalizedInventoryLine
    line.ShipmentNo  = shipNo
    line.SKU         = sku
    line.QC          = qc
    line.LotNo       = lotNo
    line.Expiry      = expiry
    line.Qty         = qty
    line.QCValid     = True
    line.ExpiryValid = True
    line.QtyValid    = True
    MakeInventoryLine = line
End Function

' 快速构建 InventoryKey。
Private Function MakeInventoryKey( _
    ByVal shipNo As String, ByVal sku As String, ByVal qc As String, _
    ByVal lotNo As String, ByVal expiry As String) As InventoryKey

    Dim key As InventoryKey
    key.ShipmentNo = shipNo
    key.SKU        = sku
    key.QC         = qc
    key.LotNo      = lotNo
    key.Expiry     = expiry
    MakeInventoryKey = key
End Function

' 安全统计 InventoryRow 数组元素数（未初始化时返回 0，避免 VBA 报错）。
Private Function SafeInventoryRowCount(ByRef rows() As InventoryRow) As Long
    On Error GoTo NotAllocated
    SafeInventoryRowCount = UBound(rows) - LBound(rows) + 1
    Exit Function
NotAllocated:
    SafeInventoryRowCount = 0
End Function

Private Function MergeFieldIssues( _
    ByRef issuesA() As FieldNormalizeIssue, _
    ByRef issuesB() As FieldNormalizeIssue) As FieldNormalizeIssue()

    Dim countA As Long
    Dim countB As Long
    Dim i As Long
    Dim merged() As FieldNormalizeIssue

    countA = CountFieldIssues(issuesA)
    countB = CountFieldIssues(issuesB)

    If countA = 0 And countB = 0 Then
        MergeFieldIssues = merged
        Exit Function
    End If

    ReDim merged(1 To countA + countB)

    If countA > 0 Then
        For i = 1 To countA
            merged(i) = issuesA(LBound(issuesA) + i - 1)
        Next i
    End If

    If countB > 0 Then
        For i = 1 To countB
            merged(countA + i) = issuesB(LBound(issuesB) + i - 1)
        Next i
    End If

    MergeFieldIssues = merged
End Function

' =============================================================================
' M07 测试用例（UT-Candidate）
' =============================================================================

Public Sub RunSortFilterTests()
    ' M07 排序·预检测·QC筛选模块测试。
    ' 覆盖：
    '   TC-SF01 BuildStaticPlan 基本排序（TC-12数据：initQCCount升序）
    '   TC-SF02 BuildStaticPlan 二级排序（Qty降序平局处理）
    '   TC-SF03 RunPrecheck 预检测A命中（TC-06数据：initQCCount=0）
    '   TC-SF04 RunPrecheck 均不命中（TC-12数据：正常通过）
    '   TC-SF05 RunPrecheck 预检测B命中（两行竞争ZP，合计需求超供应量）
    '   TC-SF06 FilterCandidatePool 中间行筛选（TC-12行00001动态验证）
    '   TC-SF07 FilterCandidatePool 最后行仅T=D有效
    '   TC-SF08 FilterCandidatePool triedQCs排除已尝试QC
    '   TC-SF09 FilterCandidatePool 首行与BuildStaticPlan nextMinQty定义一致
    BeginSuite "M07 SortFilter Tests"

    On Error GoTo CleanFail

    TestBuildStaticPlan_TC12Sort
    TestBuildStaticPlan_MultiLevelSort
    TestRunPrecheck_A_Hit
    TestRunPrecheck_Neither
    TestRunPrecheck_B_Hit
    TestFilterPool_MiddleRow
    TestFilterPool_LastRow
    TestFilterPool_TriedQCExcluded
    TestFilterPool_SameNextMinQtyAsStatic

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M07 SortFilter Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-SF01：BuildStaticPlan 基本排序（TC-12数据）
' 库存：ZP:2, QC:5, NG:5；三行 D={2,5,5}
' 行00001（D=2, nextMinQty=5）：ZP(T=2=D→可用), QC(T=5<7→不可用), NG(T=5<7→不可用) → initQCCount=1
' 行00002/00003（D=5, nextMinQty=2）：QC(T=5=D→可用), NG(T=5=D→可用), ZP(T=2<7且T≠5→不可用) → initQCCount=2
' 预期：00001排第1（initQCCount=1最少），00002排第2（行号升序），00003排第3
Private Sub TestBuildStaticPlan_TC12Sort()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_TC12_ST", "H_TC12_ST", "ZP", "LA01", "2029/01/01", 2)
    inv(2) = MakeInventoryLine("SF_TC12_ST", "H_TC12_ST", "QC", "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine("SF_TC12_ST", "H_TC12_ST", "NG", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 3) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_TC12_ST", "WMS_TC12", "H_TC12_ST", "00001", 2)
    rows(2) = SF_MakeReturnLine("SF_TC12_ST", "WMS_TC12", "H_TC12_ST", "00002", 5)
    rows(3) = SF_MakeReturnLine("SF_TC12_ST", "WMS_TC12", "H_TC12_ST", "00003", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    AssertEqualLong "TC-SF01 行数=3", 3, CLng(plan("RowCount"))
    AssertEqualLong "TC-SF01 GroupMinQty=2", 2, CLng(plan("GroupMinQty"))
    AssertEqualString "TC-SF01 排序后第1行=00001（initQCCount=1最小）", "00001", CStr(plan("LineNo_1"))
    AssertEqualLong "TC-SF01 第1行initQCCount=1", 1, CLng(plan("InitQCCount_1"))
    AssertEqualLong "TC-SF01 第2行initQCCount=2", 2, CLng(plan("InitQCCount_2"))
    AssertEqualLong "TC-SF01 第3行initQCCount=2", 2, CLng(plan("InitQCCount_3"))
    ' 第2、3行 initQCCount相同、D相同、退单号相同 → 行号升序兜底：00002先于00003
    AssertEqualString "TC-SF01 排序后第2行=00002（行号升序）", "00002", CStr(plan("LineNo_2"))
    AssertEqualString "TC-SF01 排序后第3行=00003", "00003", CStr(plan("LineNo_3"))
End Sub

' TC-SF02：BuildStaticPlan 二级排序（Qty降序）
' 库存：ZP:20, QC:20；两行D=8和D=4（initQCCount均=2，触发Qty降序）
' 行00001（D=8, nextMinQty=4）：ZP(20>=12→可用), QC(20>=12→可用) → initQCCount=2
' 行00002（D=4, nextMinQty=8）：ZP(20>=12→可用), QC(20>=12→可用) → initQCCount=2
' 预期：D=8的行00001排第1（Qty降序，大需求优先）
Private Sub TestBuildStaticPlan_MultiLevelSort()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_SORT2", "H_SORT2", "ZP", "LA01", "2029/01/01", 20)
    inv(2) = MakeInventoryLine("SF_SORT2", "H_SORT2", "QC", "LA01", "2029/01/01", 20)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 2) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_SORT2", "WMS_SORT2", "H_SORT2", "00001", 8)
    rows(2) = SF_MakeReturnLine("SF_SORT2", "WMS_SORT2", "H_SORT2", "00002", 4)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    AssertEqualLong "TC-SF02 行数=2", 2, CLng(plan("RowCount"))
    AssertEqualString "TC-SF02 D=8的行排第1（Qty降序）", "00001", CStr(plan("LineNo_1"))
    AssertEqualString "TC-SF02 D=4的行排第2", "00002", CStr(plan("LineNo_2"))
    AssertEqualLong "TC-SF02 第1行Qty=8", 8, CLng(plan("Qty_1"))
    AssertEqualLong "TC-SF02 第2行Qty=4", 4, CLng(plan("Qty_2"))
    AssertEqualLong "TC-SF02 两行initQCCount均=2", 2, CLng(plan("InitQCCount_1"))
End Sub

' TC-SF03：RunPrecheck A命中（TC-06/TC-39场景）
' 库存：ZP:5, QC:6；两行D={10,1}
' 行D=10, nextMinQty=1：ZP(5)≠10且5<11→不可用；QC(6)≠10且6<11→不可用 → initQCCount=0
' 预期：PrecheckAHit=True（initQCCount=0直接命中，不检测B）
Private Sub TestRunPrecheck_A_Hit()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_PRE_A", "H_PRE_A", "ZP", "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF_PRE_A", "H_PRE_A", "QC", "LA01", "2029/01/01", 6)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 2) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_PRE_A", "TK_PRE_A", "H_PRE_A", "00001", 10)
    rows(2) = SF_MakeReturnLine("SF_PRE_A", "TK_PRE_A", "H_PRE_A", "00002", 1)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    ' 验证initQCCount=0是预检测A的根源（两行均应为0）
    AssertEqualLong "TC-SF03 第1行initQCCount=0", 0, CLng(plan("InitQCCount_1"))
    AssertEqualLong "TC-SF03 第2行initQCCount=0", 0, CLng(plan("InitQCCount_2"))

    Dim pr As PrecheckResult
    pr = RunPrecheck(plan, ledger)

    AssertTrue "TC-SF03 PrecheckAHit=True", pr.PrecheckAHit
    AssertFalse "TC-SF03 PrecheckBHit=False（A命中后不再判断B）", pr.PrecheckBHit
End Sub

' TC-SF04：RunPrecheck 均不命中（TC-12正常数据）
' 库存：ZP:2, QC:5, NG:5；三行D={2,5,5}（initQCCount={1,2,2}，无=0，无B竞争）
' 预期：PrecheckAHit=False, PrecheckBHit=False
Private Sub TestRunPrecheck_Neither()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_PRE_N", "H_PRE_N", "ZP", "LA01", "2029/01/01", 2)
    inv(2) = MakeInventoryLine("SF_PRE_N", "H_PRE_N", "QC", "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine("SF_PRE_N", "H_PRE_N", "NG", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 3) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_PRE_N", "WMS_PRE_N", "H_PRE_N", "00001", 2)
    rows(2) = SF_MakeReturnLine("SF_PRE_N", "WMS_PRE_N", "H_PRE_N", "00002", 5)
    rows(3) = SF_MakeReturnLine("SF_PRE_N", "WMS_PRE_N", "H_PRE_N", "00003", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    Dim pr As PrecheckResult
    pr = RunPrecheck(plan, ledger)

    AssertFalse "TC-SF04 PrecheckAHit=False", pr.PrecheckAHit
    AssertFalse "TC-SF04 PrecheckBHit=False", pr.PrecheckBHit
End Sub

' TC-SF05：RunPrecheck B命中（两行竞争ZP，合计需求超供应量）
' 库存：ZP:3（无QC/NG）；两行D均=3
' 行00001（D=3, nextMinQty=3）：ZP(T=3=D→可用)，其他T=0→不可用 → initQCCount=1，锁定ZP
' 行00002（D=3, nextMinQty=3）：同上 → initQCCount=1，锁定ZP
' forcedDemand=3+3=6 > ZP(T=3)=3 → PrecheckBHit=True
Private Sub TestRunPrecheck_B_Hit()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_PRE_B", "H_PRE_B", "ZP", "LA01", "2029/01/01", 3)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 2) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_PRE_B", "WMS_PRE_B", "H_PRE_B", "00001", 3)
    rows(2) = SF_MakeReturnLine("SF_PRE_B", "WMS_PRE_B", "H_PRE_B", "00002", 3)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    ' 两行均应锁定到ZP（initQCCount=1）
    AssertEqualLong "TC-SF05 行1 initQCCount=1（锁定ZP）", 1, CLng(plan("InitQCCount_1"))
    AssertEqualLong "TC-SF05 行2 initQCCount=1（锁定ZP）", 1, CLng(plan("InitQCCount_2"))

    Dim pr As PrecheckResult
    pr = RunPrecheck(plan, ledger)

    AssertFalse "TC-SF05 PrecheckAHit=False（两行均有1个可用QC）", pr.PrecheckAHit
    AssertTrue "TC-SF05 PrecheckBHit=True（ZP=3 < 合计需求6）", pr.PrecheckBHit
End Sub

' TC-SF06：FilterCandidatePool 中间行筛选（TC-12行00001的动态验证）
' 使用TC-12初始库存；当前行=WMS_TC12:00001（位置1，最先处理）
' D=2；rows after={D=5, D=5} → nextMinQty=5
' ZP(T=2): 2=D=2 → 可用
' QC(T=5): 5≠2 且 5<2+5=7 → 不可用（核心验证：动态nextMinQty正确为5而非groupMinQty=2）
' NG(T=5): 同QC → 不可用
' 预期：候选池中只有ZP（1行）
Private Sub TestFilterPool_MiddleRow()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_TC12_F", "H_TC12_F", "ZP", "LA01", "2029/01/01", 2)
    inv(2) = MakeInventoryLine("SF_TC12_F", "H_TC12_F", "QC", "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine("SF_TC12_F", "H_TC12_F", "NG", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 3) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_TC12_F", "WMS_TC12F", "H_TC12_F", "00001", 2)
    rows(2) = SF_MakeReturnLine("SF_TC12_F", "WMS_TC12F", "H_TC12_F", "00002", 5)
    rows(3) = SF_MakeReturnLine("SF_TC12_F", "WMS_TC12F", "H_TC12_F", "00003", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    Dim emptyTried() As String
    Dim lineKey As String
    lineKey = MakePlanLineKey("WMS_TC12F", "00001")

    Dim pool() As CandidateRow
    pool = FilterCandidatePool(lineKey, plan, ledger, emptyTried)

    AssertEqualLong "TC-SF06 候选池仅1行（动态nextMinQty=5正确筛除QC和NG）", 1, SF_SafeCandidateRowCount(pool)
    AssertEqualString "TC-SF06 唯一候选QC=ZP", QC_ZP, pool(1).QC
    AssertEqualLong "TC-SF06 候选CurrentQty=2", 2, pool(1).CurrentQty
End Sub

' TC-SF07：FilterCandidatePool 最后行——仅T=D有效
' 库存：只有QC:5和NG:5（ZP已耗尽）；只有一行D=5（最后行）
' isLastRow=True → 仅 T=D 有效
' QC(T=5): 5=D=5 → 可用；NG(T=5): 5=D=5 → 可用；ZP(T=0) → 不参与
' 预期：候选池共2行（QC和NG各1行）
Private Sub TestFilterPool_LastRow()
    ' 构造只含QC和NG的库存（模拟ZP在前几行已被分配完毕的状态）
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_LAST", "H_LAST", "QC", "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF_LAST", "H_LAST", "NG", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 1) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_LAST", "WMS_LAST", "H_LAST", "00001", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    Dim emptyTried() As String
    Dim lineKey As String
    lineKey = MakePlanLineKey("WMS_LAST", "00001")

    Dim pool() As CandidateRow
    pool = FilterCandidatePool(lineKey, plan, ledger, emptyTried)

    AssertEqualLong "TC-SF07 最后行候选=2行（QC+NG）", 2, SF_SafeCandidateRowCount(pool)
    AssertTrue "TC-SF07 包含QC", SF_PoolContainsQC(pool, QC_QC)
    AssertTrue "TC-SF07 包含NG", SF_PoolContainsQC(pool, QC_NG)
    AssertFalse "TC-SF07 不含ZP（T=0）", SF_PoolContainsQC(pool, QC_ZP)
End Sub

' TC-SF08：FilterCandidatePool 排除 triedQCs 中的QC
' 库存：ZP:5, QC:5（两者T=D=5均可用）；单行D=5（最后行）
' triedQCs = {"QC"}（QC已在本行尝试过并失败，需排除）
' 预期：候选池只有ZP（1行）
Private Sub TestFilterPool_TriedQCExcluded()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_TRIED", "H_TRIED", "ZP", "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF_TRIED", "H_TRIED", "QC", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 1) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_TRIED", "WMS_TRIED", "H_TRIED", "00001", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    Dim tried(1 To 1) As String
    tried(1) = QC_QC  ' QC 已被尝试过，应被排除

    Dim lineKey As String
    lineKey = MakePlanLineKey("WMS_TRIED", "00001")

    Dim pool() As CandidateRow
    pool = FilterCandidatePool(lineKey, plan, ledger, tried)

    AssertEqualLong "TC-SF08 候选=1行（QC被排除）", 1, SF_SafeCandidateRowCount(pool)
    AssertEqualString "TC-SF08 唯一候选=ZP", QC_ZP, pool(1).QC
    AssertFalse "TC-SF08 不含QC（已在triedQCs中）", SF_PoolContainsQC(pool, QC_QC)
End Sub

' TC-SF09：规格测试点验证——"BuildStaticPlan 和 FilterCandidatePool 使用完全相同的 nextMinQty 定义"
' 含义：对排序后第1行（首次调用FilterCandidatePool，账本状态与BuildStaticPlan时相同），
' BuildStaticPlan 计算的 initQCCount_1 与 FilterCandidatePool 返回的候选QC种类数应一致。
' 使用TC-12数据：initQCCount_1=1（只有ZP可用），FilterCandidatePool也应只返回ZP
Private Sub TestFilterPool_SameNextMinQtyAsStatic()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_TC12_C", "H_TC12_C", "ZP", "LA01", "2029/01/01", 2)
    inv(2) = MakeInventoryLine("SF_TC12_C", "H_TC12_C", "QC", "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine("SF_TC12_C", "H_TC12_C", "NG", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 3) As NormalizedReturnLine
    rows(1) = SF_MakeReturnLine("SF_TC12_C", "WMS_TC12C", "H_TC12_C", "00001", 2)
    rows(2) = SF_MakeReturnLine("SF_TC12_C", "WMS_TC12C", "H_TC12_C", "00002", 5)
    rows(3) = SF_MakeReturnLine("SF_TC12_C", "WMS_TC12C", "H_TC12_C", "00003", 5)

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)

    ' BuildStaticPlan 对首行（00001）计算 initQCCount=1
    Dim staticQCCount As Long
    staticQCCount = CLng(plan("InitQCCount_1"))
    AssertEqualLong "TC-SF09 BuildStaticPlan 首行initQCCount=1", 1, staticQCCount

    ' FilterCandidatePool 在初始状态（账本未改动）对首行筛选
    Dim emptyTried() As String
    Dim lineKey As String
    lineKey = MakePlanLineKey("WMS_TC12C", "00001")

    Dim pool() As CandidateRow
    pool = FilterCandidatePool(lineKey, plan, ledger, emptyTried)

    ' 候选池中的QC种类数应与 initQCCount 相同（同为1）
    Dim dynamicQCCount As Long
    dynamicQCCount = SF_CountDistinctQCsInPool(pool)

    AssertEqualLong "TC-SF09 FilterCandidatePool 首行候选QC种类=1", 1, dynamicQCCount
    AssertEqualLong "TC-SF09 静态initQCCount与动态候选QC种类数一致", staticQCCount, dynamicQCCount
    AssertEqualString "TC-SF09 唯一候选QC=ZP", QC_ZP, pool(1).QC
End Sub

' -----------------------------------------------------------------------------
' M07 测试辅助函数
' -----------------------------------------------------------------------------

' 快速构建一个合法的 NormalizedReturnLine（默认两个 Valid 标记均为 True）
Private Function SF_MakeReturnLine(ByVal shipNo As String, ByVal wmsOrderNo As String, _
                                    ByVal sku As String, ByVal lineNo As String, _
                                    ByVal qty As Long) As NormalizedReturnLine
    Dim line As NormalizedReturnLine
    line.ShipmentNo  = shipNo
    line.WMSOrderNo  = wmsOrderNo
    line.SKU         = sku
    line.LineNo      = lineNo
    line.Qty         = qty
    line.ExcelRowNum = 0
    line.LineNoValid = True
    line.QtyValid    = True
    SF_MakeReturnLine = line
End Function

' 安全统计 CandidateRow 数组元素数（未初始化时返回0，不抛错）
Private Function SF_SafeCandidateRowCount(ByRef rows() As CandidateRow) As Long
    On Error GoTo NotInit
    SF_SafeCandidateRowCount = UBound(rows) - LBound(rows) + 1
    Exit Function
NotInit:
    SF_SafeCandidateRowCount = 0
End Function

' 判断候选池中是否包含指定QC类型的行（至少有一行该QC）
Private Function SF_PoolContainsQC(ByRef pool() As CandidateRow, ByVal qc As String) As Boolean
    Dim cnt As Long
    cnt = SF_SafeCandidateRowCount(pool)
    If cnt = 0 Then
        SF_PoolContainsQC = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(pool) To UBound(pool)
        If pool(i).QC = qc Then
            SF_PoolContainsQC = True
            Exit Function
        End If
    Next i

    SF_PoolContainsQC = False
End Function

' 统计候选池中不同QC种类的数量（用于验证 initQCCount 定义一致性）
Private Function SF_CountDistinctQCsInPool(ByRef pool() As CandidateRow) As Long
    Dim cnt As Long
    cnt = SF_SafeCandidateRowCount(pool)
    If cnt = 0 Then
        SF_CountDistinctQCsInPool = 0
        Exit Function
    End If

    ' 只检查三种合法QC（ZP / QC / NG），与 M07 内部枚举一致
    Dim hasZP As Boolean
    Dim hasQC As Boolean
    Dim hasNG As Boolean
    hasZP = False
    hasQC = False
    hasNG = False

    Dim i As Long
    For i = LBound(pool) To UBound(pool)
        Select Case pool(i).QC
            Case QC_ZP: hasZP = True
            Case QC_QC: hasQC = True
            Case QC_NG: hasNG = True
        End Select
    Next i

    Dim total As Long
    total = 0
    If hasZP Then total = total + 1
    If hasQC Then total = total + 1
    If hasNG Then total = total + 1

    SF_CountDistinctQCsInPool = total
End Function

' =============================================================================
' M08 测试用例（UT-Strategy）
' =============================================================================

Public Sub RunStrategyTests()
    ' M08 分配策略模块测试。
    ' 覆盖：
    '   TC-ST01 策略一精确匹配（单 lot = demand）
    '   TC-ST02 策略一 QC 优先级（ZP/QC/NG 均精确匹配，选 ZP）
    '   TC-ST03 策略二最小 sufficient（单 lot > demand）
    '   TC-ST04 策略二选最小保留大库存（两个 sufficient lot，选较小的）
    '   TC-ST05 策略三同 QC 跨批号合并（单 lot 不足，同 QC 多 lot 之和足够）
    '   TC-ST06 策略三第一步按数量距离锁定 QC，不跨 QC
    '   TC-ST07 全策略失败，账本不变（完整回滚保证）
    '   TC-ST08 CompareByPriority QC 优先级顺序（ZP > QC > NG）
    '   TC-ST09 CompareByPriority 效期降序（同 QC 时晚效期优先）
    '   TC-ST10 CompareByPriority 批号升序（同 QC 同效期时的平局兜底）
    '   TC-ST11 策略一平局时晚效期优先
    '   TC-ST12 策略二等量平局时晚效期优先
    '   TC-ST13 策略三后续步骤等距时优先一步完成
    '   TC-ST14 策略三第一步找到最近行后不被后续较差候选覆盖
    BeginSuite "M08 Strategy Tests"

    On Error GoTo CleanFail

    TestStrategy_One_ExactMatch
    TestStrategy_One_QCPriority
    TestStrategy_Two_SmallestSufficient
    TestStrategy_Two_PreservesLargerLot
    TestStrategy_Three_SameQCMultiLot
    TestStrategy_Three_NoSpanQC
    TestStrategy_AllFail_LedgerUnchanged
    TestCompare_QCOrder
    TestCompare_ExpiryDesc
    TestCompare_LotNoTieBreak
    TestStrategy_One_LateExpiry
    TestStrategy_Two_LateExpiry
    TestStrategy_Three_NextClosest
    TestStrategy_Three_FirstKeepsBest

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M08 Strategy Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-ST01：策略一精确匹配
' 候选池：ZP/LA01/qty=10, QC/LA01/qty=5；demand=10
' ZP qty 恰好等于 demand → 策略一命中，QC 不动
Private Sub TestStrategy_One_ExactMatch()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST01", "H_ST01", "ZP", "LA01", "2029/01/01", 10)
    inv(2) = MakeInventoryLine("SF_ST01", "H_ST01", "QC", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 2) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST01", "H_ST01", "ZP", "LA01", "2029/01/01", 10)
    pool(2) = ST_MakeCandidateRow("SF_ST01", "H_ST01", "QC", "LA01", "2029/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 10, ledger)

    AssertTrue "TC-ST01 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST01 策略=策略一", "策略一", CStr(attempt("StrategyUsed"))
    AssertEqualLong "TC-ST01 DetailCount=1", 1, CLng(attempt("DetailCount"))
    AssertEqualString "TC-ST01 Detail_1 QC=ZP", QC_ZP, CStr(attempt("QC_1"))
    AssertEqualLong "TC-ST01 Detail_1 AllocQty=10", 10, CLng(attempt("AllocQty_1"))
    ' 账本验证：ZP 扣减至0，QC 未动
    AssertEqualLong "TC-ST01 ZP 已扣减至0", 0, _
        QueryQCTotal(ledger, "SF_ST01", "H_ST01", QC_ZP)
    AssertEqualLong "TC-ST01 QC 未动=5", 5, _
        QueryQCTotal(ledger, "SF_ST01", "H_ST01", QC_QC)
End Sub

' TC-ST02：策略一 QC 优先级（ZP/QC/NG 均有精确匹配，应选 ZP）
' 候选池传入顺序故意乱序（NG/QC/ZP），确认策略会按优先级选 ZP 而非第一个遇到的
Private Sub TestStrategy_One_QCPriority()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST02", "H_ST02", "NG", "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF_ST02", "H_ST02", "QC", "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine("SF_ST02", "H_ST02", "ZP", "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    ' pool 故意以 NG→QC→ZP 顺序传入，测试排序是否生效
    Dim pool(1 To 3) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST02", "H_ST02", "NG", "LA01", "2029/01/01", 5)
    pool(2) = ST_MakeCandidateRow("SF_ST02", "H_ST02", "QC", "LA01", "2029/01/01", 5)
    pool(3) = ST_MakeCandidateRow("SF_ST02", "H_ST02", "ZP", "LA01", "2029/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 5, ledger)

    AssertTrue "TC-ST02 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST02 策略=策略一", "策略一", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST02 选中 QC=ZP（最高优先级）", QC_ZP, CStr(attempt("QC_1"))
    AssertEqualLong "TC-ST02 ZP 已扣减至0", 0, _
        QueryQCTotal(ledger, "SF_ST02", "H_ST02", QC_ZP)
    AssertEqualLong "TC-ST02 QC 未动=5", 5, _
        QueryQCTotal(ledger, "SF_ST02", "H_ST02", QC_QC)
    AssertEqualLong "TC-ST02 NG 未动=5", 5, _
        QueryQCTotal(ledger, "SF_ST02", "H_ST02", QC_NG)
End Sub

' TC-ST03：策略二最小 sufficient（单 lot > demand）
' 候选池：ZP/LA01/qty=15；demand=10（15≠10 且 15>10）
' 策略一失败，策略二命中：扣减10件，剩余5件
Private Sub TestStrategy_Two_SmallestSufficient()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST03", "H_ST03", "ZP", "LA01", "2029/01/01", 15)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 1) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST03", "H_ST03", "ZP", "LA01", "2029/01/01", 15)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 10, ledger)

    AssertTrue "TC-ST03 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST03 策略=策略二", "策略二", CStr(attempt("StrategyUsed"))
    AssertEqualLong "TC-ST03 DetailCount=1", 1, CLng(attempt("DetailCount"))
    AssertEqualLong "TC-ST03 AllocQty=10", 10, CLng(attempt("AllocQty_1"))
    ' 账本验证：15-10=5
    AssertEqualLong "TC-ST03 ZP 剩余=5", 5, _
        QueryQCTotal(ledger, "SF_ST03", "H_ST03", QC_ZP)
End Sub

' TC-ST04：策略二选最小，保留较大库存
' 候选池：ZP/LA01/qty=15, ZP/LA02/qty=20；demand=10
' 两个 lot 均 > 10，应优先选较小的（LA01=15），保留大库存（LA02=20）供后续行使用
Private Sub TestStrategy_Two_PreservesLargerLot()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST04", "H_ST04", "ZP", "LA01", "2029/01/01", 15)
    inv(2) = MakeInventoryLine("SF_ST04", "H_ST04", "ZP", "LA02", "2029/01/01", 20)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 2) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST04", "H_ST04", "ZP", "LA01", "2029/01/01", 15)
    pool(2) = ST_MakeCandidateRow("SF_ST04", "H_ST04", "ZP", "LA02", "2029/01/01", 20)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 10, ledger)

    AssertTrue "TC-ST04 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST04 策略=策略二", "策略二", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST04 选中较小lot=LA01", "LA01", CStr(attempt("LotNo_1"))
    AssertEqualLong "TC-ST04 AllocQty=10", 10, CLng(attempt("AllocQty_1"))

    ' 账本验证：LA01=5（扣减后），LA02=20（未动）
    Dim rows() As InventoryRow
    rows = GetFiveTupleRows(ledger, "SF_ST04", "H_ST04", QC_ZP)
    Dim la01Rem As Long
    Dim la02Rem As Long
    Dim ri As Long
    For ri = LBound(rows) To UBound(rows)
        If rows(ri).LotNo = "LA01" Then la01Rem = rows(ri).CurrentQty
        If rows(ri).LotNo = "LA02" Then la02Rem = rows(ri).CurrentQty
    Next ri
    AssertEqualLong "TC-ST04 LA01 剩余=5", 5, la01Rem
    AssertEqualLong "TC-ST04 LA02 未动=20", 20, la02Rem
End Sub

' TC-ST05：策略三同 QC 跨批号（单 lot 不足，同 QC 合并足够）
' 候选池：ZP/LA01/qty=3, ZP/LA02/qty=5；demand=7
' 策略一失败（无lot=7），策略二失败（无lot>7），策略三：ZP总=8≥7 → 成功
' 第一步按 |qty-demand| 选最接近的 LA02=5，再由 LA01 补足2件
Private Sub TestStrategy_Three_SameQCMultiLot()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST05", "H_ST05", "ZP", "LA01", "2029/01/01", 3)
    inv(2) = MakeInventoryLine("SF_ST05", "H_ST05", "ZP", "LA02", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 2) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST05", "H_ST05", "ZP", "LA01", "2029/01/01", 3)
    pool(2) = ST_MakeCandidateRow("SF_ST05", "H_ST05", "ZP", "LA02", "2029/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 7, ledger)

    AssertTrue "TC-ST05 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST05 策略=策略三", "策略三", CStr(attempt("StrategyUsed"))
    AssertEqualLong "TC-ST05 DetailCount=2", 2, CLng(attempt("DetailCount"))
    AssertEqualString "TC-ST05 Detail_1 最接近需求=LA02", "LA02", CStr(attempt("LotNo_1"))
    AssertEqualLong "TC-ST05 Detail_1 AllocQty=5（LA02全取）", 5, CLng(attempt("AllocQty_1"))
    AssertEqualString "TC-ST05 Detail_2 补足=LA01", "LA01", CStr(attempt("LotNo_2"))
    AssertEqualLong "TC-ST05 Detail_2 AllocQty=2（LA01部分取）", 2, CLng(attempt("AllocQty_2"))
    ' 账本：LA01=0, LA02=1，总剩余=1
    AssertEqualLong "TC-ST05 ZP 总剩余=1", 1, _
        QueryQCTotal(ledger, "SF_ST05", "H_ST05", QC_ZP)
End Sub

' TC-ST06：策略三第一步按数量距离锁定 QC，后续严禁跨 QC
' 候选池：ZP 两行各4（总8），QC 两行各5（总10）
' demand=7（单lot均<7，无精确/sufficient）
' QC 的单行5与需求距离2，小于ZP单行4的距离3，所以先选QC并锁定QC。
Private Sub TestStrategy_Three_NoSpanQC()
    Dim inv(1 To 4) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST06", "H_ST06", "ZP", "LA01", "2029/01/01", 4)
    inv(2) = MakeInventoryLine("SF_ST06", "H_ST06", "ZP", "LA02", "2029/01/01", 4)
    inv(3) = MakeInventoryLine("SF_ST06", "H_ST06", "QC", "LB01", "2029/01/01", 5)
    inv(4) = MakeInventoryLine("SF_ST06", "H_ST06", "QC", "LB02", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 4) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST06", "H_ST06", "ZP", "LA01", "2029/01/01", 4)
    pool(2) = ST_MakeCandidateRow("SF_ST06", "H_ST06", "ZP", "LA02", "2029/01/01", 4)
    pool(3) = ST_MakeCandidateRow("SF_ST06", "H_ST06", "QC", "LB01", "2029/01/01", 5)
    pool(4) = ST_MakeCandidateRow("SF_ST06", "H_ST06", "QC", "LB02", "2029/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 7, ledger)

    AssertTrue "TC-ST06 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST06 策略=策略三", "策略三", CStr(attempt("StrategyUsed"))
    ' 数量距离优先于 QC 平局规则，因此 ZP 完全不动。
    AssertEqualLong "TC-ST06 ZP 未动=8", 8, _
        QueryQCTotal(ledger, "SF_ST06", "H_ST06", QC_ZP)
    AssertEqualLong "TC-ST06 QC 剩余=3", 3, _
        QueryQCTotal(ledger, "SF_ST06", "H_ST06", QC_QC)
    ' 明细均为 QC
    AssertEqualString "TC-ST06 Detail_1 QC=QC", QC_QC, CStr(attempt("QC_1"))
End Sub

' TC-ST07：全策略失败，账本完全不变
' 候选池：ZP/LA01/qty=3；demand=7（3 < 7，策略一二三均失败）
' 期望：Success=False, StrategyUsed="失败"，ZP 账本不变
Private Sub TestStrategy_AllFail_LedgerUnchanged()
    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST07", "H_ST07", "ZP", "LA01", "2029/01/01", 3)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 1) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST07", "H_ST07", "ZP", "LA01", "2029/01/01", 3)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 7, ledger)

    AssertFalse "TC-ST07 Success=False", CBool(attempt("Success"))
    AssertEqualString "TC-ST07 StrategyUsed=失败", "失败", CStr(attempt("StrategyUsed"))
    AssertEqualLong "TC-ST07 DetailCount=0", 0, CLng(attempt("DetailCount"))
    ' 账本不变：ZP 仍为 3
    AssertEqualLong "TC-ST07 ZP 账本不变=3", 3, _
        QueryQCTotal(ledger, "SF_ST07", "H_ST07", QC_ZP)
End Sub

' TC-ST08：CompareByPriority QC 优先级（ZP > QC > NG）
' 验证三种 QC 的两两比较符号（< 0 = a 排前，> 0 = b 排前，= 0 = 相等）
Private Sub TestCompare_QCOrder()
    Dim zpRow As CandidateRow
    Dim qcRow As CandidateRow
    Dim ngRow As CandidateRow
    zpRow = ST_MakeCandidateRow("SF_CM", "H_CM", QC_ZP, "LA01", "2029/01/01", 5)
    qcRow = ST_MakeCandidateRow("SF_CM", "H_CM", QC_QC, "LA01", "2029/01/01", 5)
    ngRow = ST_MakeCandidateRow("SF_CM", "H_CM", QC_NG, "LA01", "2029/01/01", 5)

    AssertTrue "TC-ST08 ZP>QC：Compare(ZP,QC)<0", CompareByPriority(zpRow, qcRow) < 0
    AssertTrue "TC-ST08 QC>NG：Compare(QC,NG)<0", CompareByPriority(qcRow, ngRow) < 0
    AssertTrue "TC-ST08 ZP>NG：Compare(ZP,NG)<0", CompareByPriority(zpRow, ngRow) < 0
    AssertTrue "TC-ST08 反向 QC<ZP：Compare(QC,ZP)>0", CompareByPriority(qcRow, zpRow) > 0
    AssertEqualLong "TC-ST08 相同时=0", 0, CLng(CompareByPriority(zpRow, zpRow))
End Sub

' TC-ST09：CompareByPriority 效期降序（同 QC 时，晚效期优先）
' 相同 QC/批号，效期不同：2030 比 2029 更优先。
Private Sub TestCompare_ExpiryDesc()
    Dim earlyRow As CandidateRow
    Dim lateRow  As CandidateRow
    earlyRow = ST_MakeCandidateRow("SF_CM", "H_CM", QC_ZP, "LA01", "2029/01/01", 5)
    lateRow  = ST_MakeCandidateRow("SF_CM", "H_CM", QC_ZP, "LA01", "2030/01/01", 5)

    AssertTrue "TC-ST09 早效期排后：Compare(2029,2030)>0", _
        CompareByPriority(earlyRow, lateRow) > 0
    AssertTrue "TC-ST09 晚效期排前：Compare(2030,2029)<0", _
        CompareByPriority(lateRow, earlyRow) < 0
End Sub

' TC-ST10：CompareByPriority 批号升序（同 QC 同效期时的平局兜底）
' "A001" < "B001"（字母序），确认 A001 排在 B001 之前
Private Sub TestCompare_LotNoTieBreak()
    Dim rowA As CandidateRow
    Dim rowB As CandidateRow
    rowA = ST_MakeCandidateRow("SF_CM", "H_CM", QC_ZP, "A001", "2029/01/01", 5)
    rowB = ST_MakeCandidateRow("SF_CM", "H_CM", QC_ZP, "B001", "2029/01/01", 5)

    AssertTrue "TC-ST10 批号A<B排前：Compare(A001,B001)<0", _
        CompareByPriority(rowA, rowB) < 0
    AssertTrue "TC-ST10 批号B>A排后：Compare(B001,A001)>0", _
        CompareByPriority(rowB, rowA) > 0
    AssertEqualLong "TC-ST10 相同批号=0", 0, CLng(CompareByPriority(rowA, rowA))
End Sub

' TC-ST11：策略一有多个同 QC 精确匹配时，必须选择晚效期。
Private Sub TestStrategy_One_LateExpiry()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST11", "H_ST11", QC_ZP, "EARLY", "2029/01/01", 5)
    inv(2) = MakeInventoryLine("SF_ST11", "H_ST11", QC_ZP, "LATE", "2030/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 2) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST11", "H_ST11", QC_ZP, "EARLY", "2029/01/01", 5)
    pool(2) = ST_MakeCandidateRow("SF_ST11", "H_ST11", QC_ZP, "LATE", "2030/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 5, ledger)

    AssertTrue "TC-ST11 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST11 策略=策略一", "策略一", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST11 选中晚效期批次", "LATE", CStr(attempt("LotNo_1"))
End Sub

' TC-ST12：策略二数量同样接近时，必须选择晚效期。
Private Sub TestStrategy_Two_LateExpiry()
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST12", "H_ST12", QC_ZP, "EARLY", "2029/01/01", 7)
    inv(2) = MakeInventoryLine("SF_ST12", "H_ST12", QC_ZP, "LATE", "2030/01/01", 7)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 2) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST12", "H_ST12", QC_ZP, "EARLY", "2029/01/01", 7)
    pool(2) = ST_MakeCandidateRow("SF_ST12", "H_ST12", QC_ZP, "LATE", "2030/01/01", 7)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 5, ledger)

    AssertTrue "TC-ST12 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST12 策略=策略二", "策略二", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST12 选中晚效期批次", "LATE", CStr(attempt("LotNo_1"))
End Sub

' TC-ST13：策略三后续步骤中，qty=3 与 qty=5 距 remaining=4 都为1；
' 此时应选择 qty=5，因为它可以一步完成，而不是选择批号更小的 qty=3。
Private Sub TestStrategy_Three_NextClosest()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST13", "H_ST13", QC_ZP, "FIRST", "2029/01/01", 6)
    inv(2) = MakeInventoryLine("SF_ST13", "H_ST13", QC_ZP, "A_SMALL", "2030/01/01", 3)
    inv(3) = MakeInventoryLine("SF_ST13", "H_ST13", QC_ZP, "B_COVER", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 3) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST13", "H_ST13", QC_ZP, "FIRST", "2029/01/01", 6)
    pool(2) = ST_MakeCandidateRow("SF_ST13", "H_ST13", QC_ZP, "A_SMALL", "2030/01/01", 3)
    pool(3) = ST_MakeCandidateRow("SF_ST13", "H_ST13", QC_ZP, "B_COVER", "2029/01/01", 5)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 10, ledger)

    AssertTrue "TC-ST13 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST13 策略=策略三", "策略三", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST13 第一步选qty=6", "FIRST", CStr(attempt("LotNo_1"))
    AssertEqualString "TC-ST13 后续等距选可完成的qty=5", "B_COVER", CStr(attempt("LotNo_2"))
    AssertEqualLong "TC-ST13 第二步只扣remaining=4", 4, CLng(attempt("AllocQty_2"))
End Sub

' TC-ST14：最佳候选后面故意放置距离更远的候选。
' 用于防止循环中的临时布尔变量残留 True，错误覆盖已经找到的最佳行。
Private Sub TestStrategy_Three_FirstKeepsBest()
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine("SF_ST14", "H_ST14", QC_ZP, "A_BEST", "2030/01/01", 6)
    inv(2) = MakeInventoryLine("SF_ST14", "H_ST14", QC_ZP, "B_NEXT", "2030/01/01", 4)
    inv(3) = MakeInventoryLine("SF_ST14", "H_ST14", QC_ZP, "Z_WORSE", "2030/01/01", 2)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim pool(1 To 3) As CandidateRow
    pool(1) = ST_MakeCandidateRow("SF_ST14", "H_ST14", QC_ZP, "A_BEST", "2030/01/01", 6)
    pool(2) = ST_MakeCandidateRow("SF_ST14", "H_ST14", QC_ZP, "B_NEXT", "2030/01/01", 4)
    pool(3) = ST_MakeCandidateRow("SF_ST14", "H_ST14", QC_ZP, "Z_WORSE", "2030/01/01", 2)

    Dim attempt As Object
    Set attempt = TryAllocate(pool, 10, ledger)

    AssertTrue "TC-ST14 Success=True", CBool(attempt("Success"))
    AssertEqualString "TC-ST14 策略=策略三", "策略三", CStr(attempt("StrategyUsed"))
    AssertEqualString "TC-ST14 第一步保持最近qty=6", "A_BEST", CStr(attempt("LotNo_1"))
    AssertEqualLong "TC-ST14 第一步扣6", 6, CLng(attempt("AllocQty_1"))
End Sub

' -----------------------------------------------------------------------------
' M08 测试辅助函数
' -----------------------------------------------------------------------------

' 快速构建 CandidateRow（CurrentQty 与 OriginalQty 相同，模拟分配前初始状态）
Private Function ST_MakeCandidateRow(ByVal shipNo As String, ByVal sku As String, _
                                      ByVal qc As String, ByVal lotNo As String, _
                                      ByVal expiry As String, _
                                      ByVal qty As Long) As CandidateRow
    Dim row As CandidateRow
    row.ShipmentNo  = shipNo
    row.SKU         = sku
    row.QC          = qc
    row.LotNo       = lotNo
    row.Expiry      = expiry
    row.OriginalQty = qty
    row.CurrentQty  = qty
    ST_MakeCandidateRow = row
End Function

' =============================================================================
' M09 测试用例（UT-Backtracking）
' =============================================================================

Public Sub RunBacktrackingTests()
    ' M09 回溯分配引擎模块测试。
    ' 覆盖：
    '   TC-BT01 AllocateShipment 某 SKU 组 E10 → 后续 SKU 组 ErrorCode="连带回滚"
    '   TC-BT02 AllocateSKUGroup（经 AllocateShipment 调用）无需回溯，BacktrackCount=0
    '   TC-BT03 回溯 1 次后成功（贪心选ZP→pos2失败→回溯选QC→成功）
    '   TC-BT03b 详细调试日志包含非最终过程事件
    '   TC-BT04 回溯超过上限（MaxBacktrackCount=2，第3次回溯触发E10）
    BeginSuite "M09 Backtracking Tests"

    On Error GoTo CleanFail

    TestAllocateShipment_E10_CascadeRollback
    TestAllocateSKUGroup_NoBacktrack
    TestAllocateSKUGroup_BacktrackOnce_Success
    TestAllocateSKUGroup_BacktrackExceedsLimit_E10

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M09 Backtracking Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-BT01：AllocateShipment 短路 + 连带回滚
' 场景：物流单号有两个SKU组。
'   SKU-A：ZP=5, QC=5, 3行 D=5；MaxBacktrackCount=2 → 第3次回溯触发 E10
'   SKU-B：ZP=5, 1行 D=5；单独分配必然成功，但因 SKU-A E10 短路 → 连带回滚
' 期望：
'   GroupCount=2
'   Group_1_SKU=SKU-A, Group_1_Success=False, Group_1_ErrorCode="E10"
'   Group_2_SKU=SKU-B, Group_2_Success=False, Group_2_ErrorCode="连带回滚"
Private Sub TestAllocateShipment_E10_CascadeRollback()
    Dim shipNo As String
    shipNo = "SF_BT01"
    Dim skuA As String
    Dim skuB As String
    skuA = "H_BT01_A"
    skuB = "H_BT01_B"

    ' --- 建立库存账本（ZP=5,QC=5 for SKU-A；ZP=5 for SKU-B）---
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, skuA, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, skuA, QC_QC, "LB01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine(shipNo, skuB, QC_ZP, "LC01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    ' --- SKU-A：3行 D=5，贪心无法成功（ZP+QC=10，需要15）→ E10（MaxBT=2）---
    Dim rowsA(1 To 3) As NormalizedReturnLine
    rowsA(1) = BT_MakeReturnLine(shipNo, "TK_BT01", skuA, "00001", 5)
    rowsA(2) = BT_MakeReturnLine(shipNo, "TK_BT01", skuA, "00002", 5)
    rowsA(3) = BT_MakeReturnLine(shipNo, "TK_BT01", skuA, "00003", 5)

    ' --- SKU-B：1行 D=5，应该成功，但被连带回滚 ---
    Dim rowsB(1 To 1) As NormalizedReturnLine
    rowsB(1) = BT_MakeReturnLine(shipNo, "TK_BT01", skuB, "00001", 5)

    ' --- 构建 planMap 和 precheckMap ---
    Dim planMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Dim precheckMap As Object
    Set precheckMap = CreateObject("Scripting.Dictionary")

    Dim planA As Object
    Set planA = BuildStaticPlan(rowsA, ledger)
    planMap.Add skuA, planA
    Dim preA As PrecheckResult
    preA = RunPrecheck(planA, ledger)
    precheckMap.Add skuA, Array(CBool(preA.PrecheckAHit), CBool(preA.PrecheckBHit))

    Dim planB As Object
    Set planB = BuildStaticPlan(rowsB, ledger)
    planMap.Add skuB, planB
    Dim preB As PrecheckResult
    preB = RunPrecheck(planB, ledger)
    precheckMap.Add skuB, Array(CBool(preB.PrecheckAHit), CBool(preB.PrecheckBHit))

    ' --- 配置：MaxBacktrackCount=2（第3次回溯触发E10）---
    Dim cfg As ConfigStruct
    cfg.MaxBacktrackCount = 2
    cfg.DebugLogLevel     = DEBUG_LEVEL_OFF
    cfg.DetailedLogLimit  = DEFAULT_DETAILED_LOG_LIMIT
    cfg.LotCaseSensitive  = False
    cfg.NoExpirySentinel  = DEFAULT_NO_EXPIRY_SENTINEL

    Dim skuList(1 To 2) As String
    skuList(1) = skuA
    skuList(2) = skuB

    ' --- 调用 AllocateShipment ---
    Dim result As Object
    Set result = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    ' --- 验证结果 ---
    AssertEqualString "TC-BT01 ShipmentNo", shipNo, CStr(result("ShipmentNo"))
    AssertEqualLong   "TC-BT01 GroupCount=2", 2, CLng(result("GroupCount"))

    ' SKU-A 应触发 E10
    AssertEqualString "TC-BT01 Group_1_SKU=skuA", skuA, CStr(result("Group_1_SKU"))
    AssertFalse       "TC-BT01 Group_1_Success=False", CBool(result("Group_1_Success"))
    AssertEqualString "TC-BT01 Group_1_ErrorCode=E10", ERR_E10, CStr(result("Group_1_ErrorCode"))

    ' SKU-B 应为连带回滚
    AssertEqualString "TC-BT01 Group_2_SKU=skuB", skuB, CStr(result("Group_2_SKU"))
    AssertFalse       "TC-BT01 Group_2_Success=False", CBool(result("Group_2_Success"))
    AssertEqualString "TC-BT01 Group_2_ErrorCode=连带回滚", "连带回滚", CStr(result("Group_2_ErrorCode"))
End Sub

' TC-BT02：无需回溯（BacktrackCount=0，一次成功）
' 场景：ZP=5，1行 D=5，策略一精确匹配
' 期望：Success=True, BacktrackCount=0, DetailCount=1
Private Sub TestAllocateSKUGroup_NoBacktrack()
    Dim shipNo As String
    shipNo = "SF_BT02"
    Dim sku As String
    sku = "H_BT02"

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 1) As NormalizedReturnLine
    rows(1) = BT_MakeReturnLine(shipNo, "TK_BT02", sku, "00001", 5)

    Dim planMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Dim precheckMap As Object
    Set precheckMap = CreateObject("Scripting.Dictionary")

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)
    planMap.Add sku, plan
    Dim pre As PrecheckResult
    pre = RunPrecheck(plan, ledger)
    precheckMap.Add sku, Array(CBool(pre.PrecheckAHit), CBool(pre.PrecheckBHit))

    Dim cfg As ConfigStruct
    cfg.MaxBacktrackCount = 200
    cfg.DebugLogLevel     = DEBUG_LEVEL_OFF
    cfg.DetailedLogLimit  = DEFAULT_DETAILED_LOG_LIMIT
    cfg.LotCaseSensitive  = False
    cfg.NoExpirySentinel  = DEFAULT_NO_EXPIRY_SENTINEL

    Dim skuList(1 To 1) As String
    skuList(1) = sku

    Dim result As Object
    Set result = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    AssertEqualLong   "TC-BT02 GroupCount=1", 1, CLng(result("GroupCount"))
    AssertTrue        "TC-BT02 Group_1_Success=True", CBool(result("Group_1_Success"))
    AssertEqualLong   "TC-BT02 BacktrackCount=0", 0, CLng(result("Group_1_BacktrackCount"))
    AssertEqualLong   "TC-BT02 DetailCount=1", 1, CLng(result("Group_1_DetailCount"))
    AssertEqualString "TC-BT02 Detail_1_QC=ZP", QC_ZP, CStr(result("Group_1_QC_1"))
    AssertEqualLong   "TC-BT02 Detail_1_AllocQty=5", 5, CLng(result("Group_1_AllocQty_1"))
    ' 账本验证：ZP 已扣减至0
    AssertEqualLong   "TC-BT02 ZP已扣减至0", 0, QueryQCTotal(ledger, shipNo, sku, QC_ZP)
End Sub

' TC-BT03：回溯1次后成功
' 数据设计（关键分析）：
'   ZP=5, QC=13；两行各 D=5
'   BuildStaticPlan 排序后：
'     Row 00002（initQCCount=1，只能用ZP）排pos1（非末行）
'     Row 00001（initQCCount=2，ZP/QC均可用）排pos2（末行）
'   贪心第1次：
'     pos1=Row00002 非末行 nmq=5：
'       ZP(T=5=D→ok), QC(T=13≥5+5=10→ok) → 贪心选ZP → ZP=0
'     pos2=Row00001 末行：
'       ZP=0(不可用), QC(T=13≠5→末行不满足T=D) → 候选池空 → FAIL
'   第1次回溯：撤销pos1(ZP)，将ZP加入pos1已尝试列表，清空后续行的已尝试记录
'   贪心第2次：
'     pos1=Row00002 跳过ZP(已尝试)：QC(13≥10→ok) → 选QC → QC=8
'     pos2=Row00001 末行：ZP(5=D→ok), QC(8≠5→不ok) → 选ZP → SUCCESS
' 期望：BacktrackCount=1, Success=True
Private Sub TestAllocateSKUGroup_BacktrackOnce_Success()
    Dim shipNo As String
    shipNo = "SF_BT03"
    Dim sku As String
    sku = "H_BT03"

    ' ZP=5（精确够一行），QC=13（非精确，但非末行的"足够"条件满足）
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_QC, "LB01", "2029/01/01", 13)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    ' 两行各 D=5
    ' BuildStaticPlan 排序：两行 initQCCount=2 相同，Qty=5 相同，WMSOrderNo 相同
    ' → 行号升序决定顺序：pos1="00001"（非末行），pos2="00002"（末行）
    '
    ' 末行(pos2="00002") FilterCandidatePool 规则 T=D（精确耗尽）：
    '   ZP(T=5=D=5→ok), QC(T=13≠5→NOT ok for last) → 只有ZP可用
    '
    ' 非末行(pos1="00001") FilterCandidatePool nmq=D_pos2=5：
    '   ZP(T=5=D→ok), QC(T=13≥5+5=10→ok) → ZP和QC均可用
    '
    ' 贪心第1次：pos1选ZP(优先)→ZP=0; pos2末行：ZP=0不ok, QC(13≠5)不ok → FAIL
    ' 回溯第1次：撤销pos1(ZP)，ZP=5，triedQCs[pos1]="ZP"
    ' 贪心第2次：pos1跳过ZP(已尝试)→取QC(ok)→QC=8; pos2末行：ZP(5=D→ok) → SUCCESS
    Dim rows(1 To 2) As NormalizedReturnLine
    rows(1) = BT_MakeReturnLine(shipNo, "TK_BT03", sku, "00001", 5)
    rows(2) = BT_MakeReturnLine(shipNo, "TK_BT03", sku, "00002", 5)

    Dim planMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Dim precheckMap As Object
    Set precheckMap = CreateObject("Scripting.Dictionary")

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)
    planMap.Add sku, plan
    Dim pre As PrecheckResult
    pre = RunPrecheck(plan, ledger)
    precheckMap.Add sku, Array(CBool(pre.PrecheckAHit), CBool(pre.PrecheckBHit))

    Dim cfg As ConfigStruct
    cfg.MaxBacktrackCount = 10
    cfg.DebugLogLevel     = DEBUG_LEVEL_DETAIL
    cfg.DetailedLogLimit  = DEFAULT_DETAILED_LOG_LIMIT
    cfg.LotCaseSensitive  = False
    cfg.NoExpirySentinel  = DEFAULT_NO_EXPIRY_SENTINEL

    Dim skuList(1 To 1) As String
    skuList(1) = sku

    Dim result As Object
    Set result = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    AssertTrue        "TC-BT03 Success=True", CBool(result("Group_1_Success"))
    AssertEqualLong   "TC-BT03 BacktrackCount=1", 1, CLng(result("Group_1_BacktrackCount"))
    AssertEqualLong   "TC-BT03 DetailCount=2（两行各一条明细）", 2, CLng(result("Group_1_DetailCount"))
    ' 验证账本：分配后 ZP=0, QC 被扣了5（剩8）
    AssertEqualLong   "TC-BT03 ZP已耗尽", 0, QueryQCTotal(ledger, shipNo, sku, QC_ZP)
    AssertEqualLong   "TC-BT03 QC剩余=8（扣了5）", 8, QueryQCTotal(ledger, shipNo, sku, QC_QC)

    Dim debugEvents() As AllocationEvent
    debugEvents = ExtractDebugEventsFromShipment(result)
    AssertTrue "TC-BT03b 详细日志行数大于最终明细数", BT_TestEventCount(debugEvents) > CLng(result("Group_1_DetailCount"))
    AssertTrue "TC-BT03b 详细日志包含非最终过程事件", BT_TestHasNonFinalEvent(debugEvents)
End Sub

Private Function BT_TestEventCount(ByRef events() As AllocationEvent) As Long
    On Error GoTo EmptyArr
    BT_TestEventCount = UBound(events) - LBound(events) + 1
    Exit Function
EmptyArr:
    BT_TestEventCount = 0
End Function

Private Function BT_TestHasNonFinalEvent(ByRef events() As AllocationEvent) As Boolean
    Dim count As Long
    count = BT_TestEventCount(events)
    If count = 0 Then Exit Function

    Dim i As Long
    For i = LBound(events) To UBound(events)
        If Not events(i).IsFinalResult Then
            BT_TestHasNonFinalEvent = True
            Exit Function
        End If
    Next i
End Function

' TC-BT04：回溯超过上限 → E10
' 场景：ZP=5, QC=5，3行 D=5；总供应10 < 需求15，必然失败
' MaxBacktrackCount=2：
'   回溯1次（undo row2，row2改用ZP失败→再次回溯）
'   回溯2次（undo row1，row1改用QC，row2取ZP，row3再次失败）
'   回溯3次 → 超过MaxBT=2 → E10
' 期望：E10, BacktrackCount=3
Private Sub TestAllocateSKUGroup_BacktrackExceedsLimit_E10()
    Dim shipNo As String
    shipNo = "SF_BT04"
    Dim sku As String
    sku = "H_BT04"

    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_QC, "LB01", "2029/01/01", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim rows(1 To 3) As NormalizedReturnLine
    rows(1) = BT_MakeReturnLine(shipNo, "TK_BT04", sku, "00001", 5)
    rows(2) = BT_MakeReturnLine(shipNo, "TK_BT04", sku, "00002", 5)
    rows(3) = BT_MakeReturnLine(shipNo, "TK_BT04", sku, "00003", 5)

    Dim planMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Dim precheckMap As Object
    Set precheckMap = CreateObject("Scripting.Dictionary")

    Dim plan As Object
    Set plan = BuildStaticPlan(rows, ledger)
    planMap.Add sku, plan
    Dim pre As PrecheckResult
    pre = RunPrecheck(plan, ledger)
    precheckMap.Add sku, Array(CBool(pre.PrecheckAHit), CBool(pre.PrecheckBHit))

    ' MaxBacktrackCount=2，第3次回溯触发E10
    Dim cfg As ConfigStruct
    cfg.MaxBacktrackCount = 2
    cfg.DebugLogLevel     = DEBUG_LEVEL_OFF
    cfg.DetailedLogLimit  = DEFAULT_DETAILED_LOG_LIMIT
    cfg.LotCaseSensitive  = False
    cfg.NoExpirySentinel  = DEFAULT_NO_EXPIRY_SENTINEL

    Dim skuList(1 To 1) As String
    skuList(1) = sku

    Dim result As Object
    Set result = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    AssertFalse       "TC-BT04 Success=False", CBool(result("Group_1_Success"))
    AssertEqualString "TC-BT04 ErrorCode=E10", ERR_E10, CStr(result("Group_1_ErrorCode"))
    ' E10触发时已超过MaxBT(2)，BacktrackCount=3
    AssertEqualLong   "TC-BT04 BacktrackCount=3", 3, CLng(result("Group_1_BacktrackCount"))
    ' E10后账本应完整还原（所有已提交行都被撤销）
    AssertEqualLong   "TC-BT04 ZP账本已还原=5", 5, QueryQCTotal(ledger, shipNo, sku, QC_ZP)
    AssertEqualLong   "TC-BT04 QC账本已还原=5", 5, QueryQCTotal(ledger, shipNo, sku, QC_QC)
End Sub

' -----------------------------------------------------------------------------
' M09 测试辅助函数
' -----------------------------------------------------------------------------

' 快速构建 NormalizedReturnLine（供 M09 测试使用，Valid标志全为 True）
Private Function BT_MakeReturnLine(ByVal shipNo As String, ByVal wmsOrderNo As String, _
                                    ByVal sku As String, ByVal lineNo As String, _
                                    ByVal qty As Long) As NormalizedReturnLine
    Dim line As NormalizedReturnLine
    line.ShipmentNo  = shipNo
    line.WMSOrderNo  = wmsOrderNo
    line.SKU         = sku
    line.LineNo      = lineNo
    line.Qty         = qty
    line.ExcelRowNum = 0
    line.LineNoValid = True
    line.QtyValid    = True
    BT_MakeReturnLine = line
End Function

' =============================================================================
' M10 测试用例（UT-Guards）
' =============================================================================

Public Sub RunGuardsTests()
    ' M10 工程守卫模块测试。
    ' 覆盖：
    '   TC-GD01 正常扣减恢复不触发：AssertConservation 返回 True
    '   TC-GD02 人为漏报一条明细：AssertConservation 返回 False
    '   TC-GD03 RaiseE99 消息格式：包含物流单号/SKU/期望/实际
'   TC-GD04 空快照/空账本不能静默通过
    BeginSuite "M10 Guards Tests"

    On Error GoTo CleanFail

    TestAssertConservation_Pass
    TestAssertConservation_Fail_MissingDetail
    TestRaiseE99_MessageContent
    TestGuards_NullInputsFail

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M10 Guards Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-GD01：正常扣减 + 完整 Detail → AssertConservation 返回 True
' 场景：ZP=10，扣减5件，Detail记录AllocQty=5
'   守恒：快照(10) == 账本(5) + detail(5) → True
Private Sub TestAssertConservation_Pass()
    Dim shipNo As String
    shipNo = "SF_GD01"
    Dim sku As String
    sku = "H_GD01"

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 10)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    ' 拍摄快照（扣减前）
    Dim snapshot As Object
    Set snapshot = TakeSnapshot(ledger, shipNo, sku)

    ' 执行扣减：ZP 扣5件
    Dim key As InventoryKey
    key.ShipmentNo = shipNo
    key.SKU        = sku
    key.QC         = QC_ZP
    key.LotNo      = "LA01"
    key.Expiry     = "2029/01/01"

    Dim undoLog As Object
    Set undoLog = NewUndoLog()
    Dim ok As Boolean
    ok = Deduct(ledger, key, 5, undoLog)
    AssertTrue "TC-GD01 Deduct成功", ok

    ' 构建 AllocationDetail（完整记录）
    Dim details(1 To 1) As AllocationDetail
    details(1).ShipmentNo = shipNo
    details(1).SKU        = sku
    details(1).QC         = QC_ZP
    details(1).LotNo      = "LA01"
    details(1).Expiry     = "2029/01/01"
    details(1).AllocQty   = 5

    ' 断言：守恒应通过（10 == 5 + 5）
    AssertTrue "TC-GD01 AssertConservation 通过", AssertConservation(snapshot, ledger, details)
End Sub

' TC-GD02：人为漏报一条明细 → AssertConservation 返回 False
' 场景：ZP=10，实际扣减了5件，但 Detail 中 AllocQty=0（漏记）
'   守恒：快照(10) ≠ 账本(5) + detail(0) = 5 → False
Private Sub TestAssertConservation_Fail_MissingDetail()
    Dim shipNo As String
    shipNo = "SF_GD02"
    Dim sku As String
    sku = "H_GD02"

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 10)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim snapshot As Object
    Set snapshot = TakeSnapshot(ledger, shipNo, sku)

    ' 执行扣减：ZP 扣5件
    Dim key As InventoryKey
    key.ShipmentNo = shipNo
    key.SKU        = sku
    key.QC         = QC_ZP
    key.LotNo      = "LA01"
    key.Expiry     = "2029/01/01"

    Dim undoLog As Object
    Set undoLog = NewUndoLog()
    Deduct ledger, key, 5, undoLog

    ' 故意漏报：AllocQty=0
    Dim details(1 To 1) As AllocationDetail
    details(1).AllocQty = 0

    ' 断言：守恒应失败（10 ≠ 5 + 0）
    AssertFalse "TC-GD02 AssertConservation 应返回False", AssertConservation(snapshot, ledger, details)
End Sub

' TC-GD03：RaiseE99 消息包含物流单号、SKU、期望值、实际值
' 通过 On Error Resume Next 捕获错误，检查 Err.Description 内容
Private Sub TestRaiseE99_MessageContent()
    Dim shipNo As String
    shipNo = "SF_GD03"
    Dim sku As String
    sku = "H_GD03"
    Dim expected As Long
    expected = 100
    Dim actual As Long
    actual = 90

    On Error Resume Next
    RaiseE99 shipNo, sku, expected, actual, "测试上下文"
    Dim errNum As Long
    errNum = Err.Number
    Dim errDesc As String
    errDesc = Err.Description
    On Error GoTo 0

    ' 验证：抛出了 E99 错误号
    AssertEqualLong "TC-GD03 Err.Number=E99错误号", E99_ERROR_NUMBER, errNum

    ' 验证：消息包含关键信息
    AssertTrue "TC-GD03 描述包含物流单号", InStr(errDesc, shipNo) > 0
    AssertTrue "TC-GD03 描述包含SKU",      InStr(errDesc, sku) > 0
    AssertTrue "TC-GD03 描述包含期望值",   InStr(errDesc, CStr(expected)) > 0
    AssertTrue "TC-GD03 描述包含实际值",   InStr(errDesc, CStr(actual)) > 0
End Sub

' TC-GD04：快照或账本为空时，守卫无法证明账本正确，必须返回 False。
Private Sub TestGuards_NullInputsFail()
    Dim details() As AllocationDetail
    Dim snapshot As Object
    Dim ledger As Object

    AssertFalse "TC-GD04 AssertConservation 空参应失败", AssertConservation(snapshot, ledger, details)
    AssertFalse "TC-GD04 AssertUndoConsistency 空参应失败", AssertUndoConsistency(snapshot, Nothing, ledger)
End Sub

' =============================================================================
' M11 测试用例（UT-Status）
' =============================================================================

Public Sub RunStatusTests()
    ' M11 状态判定模块测试。
    ' 覆盖：
    '   TC-ST01 DetermineLineStatus 1种批号/效期 → 批量导入
    '   TC-ST02 DetermineLineStatus ≥2种批号/效期 → 手工操作
    '   TC-ST03 BuildRollbackReason 连带回滚格式（E10）
    '   TC-ST04 BuildRollbackReason 多错误码升序合并
    '   TC-ST05 ApplyRollback 校验阶段 E08 → 整单无法分配 + 直接原因
    '   TC-ST06 ApplyRollback 分配阶段 E10 → 连带回滚原因格式
    '   TC-ST07 ApplyRollback 成功路径 → 退单号状态聚合（手工操作优先）
    '   TC-ST08b 不同 WMS 相同行号不能互相污染行状态
    '   TC-ST09 AggregateWMSStatus 从 FinalResult 提取汇总
    '   TC-ST10 退单表 WMS 为空时保留 [N/A] 汇总行
    BeginSuite "M11 Status Tests"

    On Error GoTo CleanFail

    TestDetermineLineStatus_OneCombo
    TestDetermineLineStatus_TwoCombos
    TestDetermineLineStatus_NoDetails
    TestBuildRollbackReason_Cascade
    TestBuildRollbackReason_DirectSorted
    TestApplyRollback_ValidationE08
    TestApplyRollback_AllocE10_Cascade
    TestApplyRollback_Success_WmsAggregate
    TestApplyRollback_SameLineNoAcrossWmsBatch
    TestAggregateWMSStatus_Extract
    TestApplyRollback_BlankWmsPlaceholder

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M11 Status Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-ST01：1 种批号/效期 → 批量导入
Private Sub TestDetermineLineStatus_OneCombo()
    Dim details(1 To 1) As AllocationDetail
    details(1).LineNo = "00001"
    details(1).LotNo = "LA01"
    details(1).Expiry = "2029/01/01"
    details(1).AllocQty = 5

    AssertEqualString "TC-ST01 行状态=批量导入", STATUS_BATCH_IMPORT, DetermineLineStatus(details, "00001")
End Sub

' TC-ST02：2 种批号/效期 → 手工操作
Private Sub TestDetermineLineStatus_TwoCombos()
    Dim details(1 To 2) As AllocationDetail
    details(1).LineNo = "00001"
    details(1).LotNo = "LA01"
    details(1).Expiry = "2029/01/01"
    details(1).AllocQty = 1

    details(2).LineNo = "00001"
    details(2).LotNo = "LA02"
    details(2).Expiry = "2030/01/01"
    details(2).AllocQty = 1

    AssertEqualString "TC-ST02 行状态=手工操作", STATUS_MANUAL, DetermineLineStatus(details, "00001")
End Sub

' TC-ST03：无有效明细 → 分配失败
Private Sub TestDetermineLineStatus_NoDetails()
    Dim details() As AllocationDetail
    AssertEqualString "TC-ST03 无明细=分配失败", LINE_STATUS_FAILED, DetermineLineStatus(details, "00001")
End Sub

' TC-ST04：连带回滚原因格式
Private Sub TestBuildRollbackReason_Cascade()
    Dim emptyCodes() As String
    AssertEqualString "TC-ST04 连带回滚格式", _
        "整单回滚（触发原因：E10）", BuildRollbackReason(emptyCodes, ERR_E10)
End Sub

' TC-ST05：多错误码升序去重
Private Sub TestBuildRollbackReason_DirectSorted()
    Dim codes(1 To 3) As String
    codes(1) = ERR_E10
    codes(2) = ERR_E01
    codes(3) = ERR_E01

    AssertEqualString "TC-ST05 多码升序合并", _
        "E01 - 关键字段为空或格式异常; E10 - 回溯超限", BuildRollbackReason(codes, ERR_E10)
End Sub

' TC-ST06：校验阶段 E08 → 两个 WMS 退单号均无法分配（直接原因）
Private Sub TestApplyRollback_ValidationE08()
    Dim shipNo As String
    shipNo = "SF_ST06"

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1) = ST_MakeReturnLine(shipNo, "TK_ST06_A", "H_ST06", "00001", 5)
    orders(2) = ST_MakeReturnLine(shipNo, "TK_ST06_B", "H_ST06", "00001", 3)

    Dim issues(1 To 1) As ValidationIssue
    issues(1).ShipmentNo = shipNo
    issues(1).WMSOrderNo = NA_PLACEHOLDER
    issues(1).SKU = "H_ST06"
    issues(1).ErrorCode = ERR_E08
    issues(1).Reason = "同物流单号+SKU数量不一致"

    Dim validationResult As ValidationResult
    validationResult.HasFailures = True
    validationResult.FailedShipmentCount = 1

    ' ByRef 参数必须传变量，不能直接传 EmptyShipmentResultArray() 函数返回值
    Dim shipmentResults() As Object
    shipmentResults = EmptyShipmentResultArray()

    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipmentResults, validationResult, issues, orders)

    Dim summary() As WMSStatusEntry
    summary = AggregateWMSStatus(finalResult)

    AssertEqualLong "TC-ST06 汇总行数=2", 2, UBound(summary) - LBound(summary) + 1
    AssertEqualString "TC-ST06 WMS-A 状态", STATUS_UNALLOCATED, ST_FindSummaryStatus(summary, "TK_ST06_A")
    AssertEqualString "TC-ST06 WMS-A 原因", "E08 - 同物流单号+SKU数量不一致", ST_FindSummaryReason(summary, "TK_ST06_A")
    AssertEqualString "TC-ST06 WMS-B 原因", "E08 - 同物流单号+SKU数量不一致", ST_FindSummaryReason(summary, "TK_ST06_B")
    AssertEqualLong "TC-ST06 成功明细数=0", 0, CLng(finalResult("DetailCount"))
End Sub

' TC-ST07：分配阶段 E10 触发整单回滚；未直接失败的 WMS 使用连带回滚格式
Private Sub TestApplyRollback_AllocE10_Cascade()
    Dim shipNo As String
    shipNo = "SF_ST07"

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1) = ST_MakeReturnLine(shipNo, "TK_ST07_A", "H_ST07_A", "00001", 5)
    orders(2) = ST_MakeReturnLine(shipNo, "TK_ST07_B", "H_ST07_B", "00001", 5)

    Dim allocMap As Object
    Set allocMap = CreateObject("Scripting.Dictionary")
    allocMap.Add "ShipmentNo", shipNo
    allocMap.Add "GroupCount", CLng(2)
    allocMap.Add "Group_1_SKU", "H_ST07_A"
    allocMap.Add "Group_1_Success", False
    allocMap.Add "Group_1_ErrorCode", ERR_E10
    allocMap.Add "Group_1_DetailCount", CLng(0)
    allocMap.Add "Group_2_SKU", "H_ST07_B"
    allocMap.Add "Group_2_Success", False
    allocMap.Add "Group_2_ErrorCode", ERROR_CASCADE_ROLLBACK
    allocMap.Add "Group_2_DetailCount", CLng(0)

    Dim shipmentResults(1 To 1) As Object
    Set shipmentResults(1) = allocMap

    Dim issues() As ValidationIssue
    issues = EmptyValidationIssueArray()
    Dim validationResult As ValidationResult
    validationResult.HasFailures = False
    validationResult.FailedShipmentCount = 0

    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipmentResults, validationResult, issues, orders)

    Dim summary() As WMSStatusEntry
    summary = AggregateWMSStatus(finalResult)

    AssertEqualLong "TC-ST07 汇总行数=2", 2, UBound(summary) - LBound(summary) + 1
    AssertEqualString "TC-ST07 WMS-A 直接原因", "E10 - 回溯超限", ST_FindSummaryReason(summary, "TK_ST07_A")
    AssertEqualString "TC-ST07 WMS-B 连带回滚", "整单回滚（触发原因：E10）", ST_FindSummaryReason(summary, "TK_ST07_B")
End Sub

' TC-ST08：成功分配 → 同一退单号混合行状态聚合为手工操作
Private Sub TestApplyRollback_Success_WmsAggregate()
    Dim shipNo As String
    shipNo = "SF_ST08"

    Dim allocMap As Object
    Set allocMap = CreateObject("Scripting.Dictionary")
    allocMap.Add "ShipmentNo", shipNo
    allocMap.Add "GroupCount", CLng(1)
    allocMap.Add "Group_1_SKU", "H_ST08"
    allocMap.Add "Group_1_Success", True
    allocMap.Add "Group_1_ErrorCode", vbNullString
    allocMap.Add "Group_1_DetailCount", CLng(3)

    ' 行00001：两种批号/效期 → 手工操作（2条明细）
    allocMap.Add "Group_1_WMSOrderNo_1", "TK_ST08"
    allocMap.Add "Group_1_LineNo_1", "00001"
    allocMap.Add "Group_1_OrderQty_1", CLng(2)
    allocMap.Add "Group_1_QC_1", QC_ZP
    allocMap.Add "Group_1_LotNo_1", "LA03"
    allocMap.Add "Group_1_Expiry_1", "2031/01/01"
    allocMap.Add "Group_1_AllocQty_1", CLng(1)
    allocMap.Add "Group_1_LineStatus_1", STATUS_MANUAL

    allocMap.Add "Group_1_WMSOrderNo_2", "TK_ST08"
    allocMap.Add "Group_1_LineNo_2", "00001"
    allocMap.Add "Group_1_OrderQty_2", CLng(2)
    allocMap.Add "Group_1_QC_2", QC_ZP
    allocMap.Add "Group_1_LotNo_2", "LA02"
    allocMap.Add "Group_1_Expiry_2", "2030/01/01"
    allocMap.Add "Group_1_AllocQty_2", CLng(1)
    allocMap.Add "Group_1_LineStatus_2", STATUS_MANUAL

    ' 行00002：一种批号/效期 → 批量导入
    allocMap.Add "Group_1_WMSOrderNo_3", "TK_ST08"
    allocMap.Add "Group_1_LineNo_3", "00002"
    allocMap.Add "Group_1_OrderQty_3", CLng(1)
    allocMap.Add "Group_1_QC_3", QC_ZP
    allocMap.Add "Group_1_LotNo_3", "LA01"
    allocMap.Add "Group_1_Expiry_3", "2029/01/01"
    allocMap.Add "Group_1_AllocQty_3", CLng(1)
    allocMap.Add "Group_1_LineStatus_3", STATUS_BATCH_IMPORT

    Dim shipmentResults(1 To 1) As Object
    Set shipmentResults(1) = allocMap

    Dim issues() As ValidationIssue
    issues = EmptyValidationIssueArray()
    Dim validationResult As ValidationResult

    Dim emptyOrders() As NormalizedReturnLine
    emptyOrders = EmptyReturnArray()

    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipmentResults, validationResult, issues, emptyOrders)

    AssertEqualLong "TC-ST08 成功明细数=3", 3, CLng(finalResult("DetailCount"))
    AssertEqualString "TC-ST08 退单号状态=手工操作", STATUS_MANUAL, CStr(finalResult("Detail_1_WMSOrderStatus"))
    AssertEqualString "TC-ST08 汇总状态=手工操作", STATUS_MANUAL, CStr(finalResult("Summary_1_Status"))
End Sub

' TC-ST08b：不同 WMS 退单号可同时存在 00001，行状态必须按完整退货行身份判定。
Private Sub TestApplyRollback_SameLineNoAcrossWmsBatch()
    Dim shipNo As String
    shipNo = "SF_ST08B"

    Dim allocMap As Object
    Set allocMap = CreateObject("Scripting.Dictionary")
    allocMap.Add "ShipmentNo", shipNo
    allocMap.Add "GroupCount", CLng(2)

    allocMap.Add "Group_1_SKU", "H_ST08B_A"
    allocMap.Add "Group_1_Success", True
    allocMap.Add "Group_1_ErrorCode", vbNullString
    allocMap.Add "Group_1_DetailCount", CLng(1)
    allocMap.Add "Group_1_WMSOrderNo_1", "TK_ST08B_A"
    allocMap.Add "Group_1_LineNo_1", "00001"
    allocMap.Add "Group_1_OrderQty_1", CLng(12)
    allocMap.Add "Group_1_QC_1", QC_QC
    allocMap.Add "Group_1_LotNo_1", "LA01"
    allocMap.Add "Group_1_Expiry_1", "2029/01/01"
    allocMap.Add "Group_1_AllocQty_1", CLng(12)
    allocMap.Add "Group_1_LineStatus_1", STATUS_MANUAL

    allocMap.Add "Group_2_SKU", "H_ST08B_B"
    allocMap.Add "Group_2_Success", True
    allocMap.Add "Group_2_ErrorCode", vbNullString
    allocMap.Add "Group_2_DetailCount", CLng(1)
    allocMap.Add "Group_2_WMSOrderNo_1", "TK_ST08B_B"
    allocMap.Add "Group_2_LineNo_1", "00001"
    allocMap.Add "Group_2_OrderQty_1", CLng(2)
    allocMap.Add "Group_2_QC_1", QC_NG
    allocMap.Add "Group_2_LotNo_1", "LB01"
    allocMap.Add "Group_2_Expiry_1", "2029/01/01"
    allocMap.Add "Group_2_AllocQty_1", CLng(2)
    allocMap.Add "Group_2_LineStatus_1", STATUS_MANUAL

    Dim shipmentResults(1 To 1) As Object
    Set shipmentResults(1) = allocMap

    Dim issues() As ValidationIssue
    issues = EmptyValidationIssueArray()
    Dim validationResult As ValidationResult

    Dim emptyOrders() As NormalizedReturnLine
    emptyOrders = EmptyReturnArray()

    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipmentResults, validationResult, issues, emptyOrders)

    AssertEqualString "TC-ST08b 明细1行状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Detail_1_LineStatus"))
    AssertEqualString "TC-ST08b 明细2行状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Detail_2_LineStatus"))
    AssertEqualString "TC-ST08b WMS-A 状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Summary_1_Status"))
    AssertEqualString "TC-ST08b WMS-B 状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Summary_2_Status"))
End Sub

' TC-ST09：AggregateWMSStatus 提取 FinalResult 中的汇总
Private Sub TestAggregateWMSStatus_Extract()
    Dim finalResult As Object
    Set finalResult = CreateObject("Scripting.Dictionary")
    finalResult.Add "SummaryCount", CLng(1)
    finalResult.Add "Summary_1_ShipmentNo", "SF_ST09"
    finalResult.Add "Summary_1_WMSOrderNo", "TK_ST09"
    finalResult.Add "Summary_1_Status", STATUS_BATCH_IMPORT
    finalResult.Add "Summary_1_Reason", vbNullString

    Dim summary() As WMSStatusEntry
    summary = AggregateWMSStatus(finalResult)

    AssertEqualLong "TC-ST09 提取行数=1", 1, UBound(summary) - LBound(summary) + 1
    AssertEqualString "TC-ST09 WMS退单号", "TK_ST09", summary(1).WMSOrderNo
    AssertEqualString "TC-ST09 状态", STATUS_BATCH_IMPORT, summary(1).Status
End Sub

' TC-ST10：同一物流单号既有正常 WMS，也有 WMS 为空的退单行。
' 汇总必须同时输出正常 WMS 和 [N/A]，不能因已有正常 WMS 而吞掉空值行。
Private Sub TestApplyRollback_BlankWmsPlaceholder()
    Dim shipNo As String
    shipNo = "SF_ST10"

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1) = ST_MakeReturnLine(shipNo, "TK_ST10", "H_ST10", "00001", 3)
    orders(2) = ST_MakeReturnLine(shipNo, vbNullString, "H_ST10", "00002", 2)
    orders(2).ExcelRowNum = 3

    Dim issues(1 To 1) As ValidationIssue
    issues(1).ShipmentNo = shipNo
    issues(1).WMSOrderNo = NA_PLACEHOLDER
    issues(1).SKU = "H_ST10"
    issues(1).ErrorCode = ERR_E01
    issues(1).SourceTable = SOURCE_RETURN_TABLE
    issues(1).ExcelRowNum = 3
    issues(1).Reason = "字段为空"

    Dim validationResult As ValidationResult
    validationResult.HasFailures = True
    validationResult.FailedShipmentCount = 1

    Dim emptyResults() As Object
    Dim finalResult As Object
    Set finalResult = ApplyRollback(emptyResults, validationResult, issues, orders)

    Dim summary() As WMSStatusEntry
    summary = AggregateWMSStatus(finalResult)

    AssertEqualLong "TC-ST10 汇总行数=2", 2, UBound(summary) - LBound(summary) + 1
    AssertEqualString "TC-ST10 正常WMS状态=无法分配", STATUS_UNALLOCATED, _
                      ST_FindSummaryStatus(summary, "TK_ST10")
    AssertEqualString "TC-ST10 [N/A]状态=无法分配", STATUS_UNALLOCATED, _
                      ST_FindSummaryStatus(summary, NA_PLACEHOLDER)
    AssertEqualString "TC-ST10 [N/A]原因为E01", _
                      "E01 - 关键字段为空或格式异常", _
                      ST_FindSummaryReason(summary, NA_PLACEHOLDER)
End Sub

Private Function ST_MakeReturnLine( _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String, _
    ByVal qty As Long) As NormalizedReturnLine

    Dim line As NormalizedReturnLine
    line.ShipmentNo = shipNo
    line.WMSOrderNo = wmsOrderNo
    line.SKU = sku
    line.LineNo = lineNo
    line.Qty = qty
    line.LineNoValid = True
    line.QtyValid = True
    ST_MakeReturnLine = line
End Function

Private Function ST_FindSummaryReason(ByRef summary() As WMSStatusEntry, ByVal wmsOrderNo As String) As String
    Dim i As Long
    For i = LBound(summary) To UBound(summary)
        If summary(i).WMSOrderNo = wmsOrderNo Then
            ST_FindSummaryReason = summary(i).Reason
            Exit Function
        End If
    Next i
End Function

Private Function ST_FindSummaryStatus(ByRef summary() As WMSStatusEntry, ByVal wmsOrderNo As String) As String
    Dim i As Long
    For i = LBound(summary) To UBound(summary)
        If summary(i).WMSOrderNo = wmsOrderNo Then
            ST_FindSummaryStatus = summary(i).Status
            Exit Function
        End If
    Next i
End Function


' =============================================================================
' M13 测试用例（UT-OutputBuild）
' =============================================================================

Public Sub RunOutputBuilderTests()
    ' M13 输出构建模块测试。
    ' 覆盖：
    '   TC-OB01 干跑汇总仅保留失败项
    '   TC-OB02 E07 汇总 WMS 退单号补 [N/A]
    '   TC-OB03 E06/E08/E11 不进异常明细
    '   TC-OB04 调试日志空数组返回空且不崩溃
    '   TC-OB07 调试日志关闭级别不写行
    '   TC-OB08 简版仅输出 IsFinalResult=True
    '   TC-OB09 调试日志 19 列映射
    '   TC-OB05 明细字段映射
    '   TC-OB06 运行历史字段构建（Dry Run）
    BeginSuite "M13 Output Builder Tests"

    On Error GoTo CleanFail

    TestOutputSummary_DryRunFilter
    TestOutputSummary_E07WmsPlaceholder
    TestAnomalyOutput_FilterSummaryOnlyCodes
    TestDebugLogRows_EmptyEvents
    TestDebugLogRows_OffLevel
    TestDebugLogRows_SimpleFilter
    TestDebugLogRows_NineteenColumns
    TestDetailRows_FieldMapping
    TestRunHistoryRow_DryRunFields

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M13 Output Builder Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-OB01：干跑模式汇总表不包含通过项（只保留无法分配）
Private Sub TestOutputSummary_DryRunFilter()
    Dim statusMap(1 To 3) As WMSStatusEntry

    statusMap(1).ShipmentNo = "SF_OB01"
    statusMap(1).WMSOrderNo = "TK_OB01_A"
    statusMap(1).Status = STATUS_BATCH_IMPORT
    statusMap(1).Reason = vbNullString

    statusMap(2).ShipmentNo = "SF_OB01"
    statusMap(2).WMSOrderNo = "TK_OB01_B"
    statusMap(2).Status = STATUS_MANUAL
    statusMap(2).Reason = vbNullString

    statusMap(3).ShipmentNo = "SF_OB01"
    statusMap(3).WMSOrderNo = "TK_OB01_C"
    statusMap(3).Status = STATUS_UNALLOCATED
    statusMap(3).Reason = "E10 - 回溯超限"

    Dim dryRows() As OutputRow
    dryRows = BuildSummaryRows(statusMap, True)
    AssertEqualLong "TC-OB01 干跑汇总仅1行", 1, OBT_OutputRowCount(dryRows)
    AssertEqualString "TC-OB01 干跑仅保留失败状态", STATUS_UNALLOCATED, OBT_RowColText(dryRows(1), 3)

    Dim fullRows() As OutputRow
    fullRows = BuildSummaryRows(statusMap, False)
    AssertEqualLong "TC-OB01 完整模式保留全部3行", 3, OBT_OutputRowCount(fullRows)
End Sub

' TC-OB02：E07 汇总行 WMS 退单号必须是 [N/A]
Private Sub TestOutputSummary_E07WmsPlaceholder()
    Dim statusMap(1 To 1) As WMSStatusEntry
    statusMap(1).ShipmentNo = "SF_OB02"
    statusMap(1).WMSOrderNo = vbNullString
    statusMap(1).Status = STATUS_UNALLOCATED
    statusMap(1).Reason = "E07 - 物流单号仅存在于质检库存表"

    Dim rows() As OutputRow
    rows = BuildSummaryRows(statusMap, False)

    AssertEqualLong "TC-OB02 汇总行数=1", 1, OBT_OutputRowCount(rows)
    AssertEqualString "TC-OB02 WMS 退单号=[N/A]", NA_PLACEHOLDER, OBT_RowColText(rows(1), 2)
End Sub

' TC-OB03：E08/E11 只进入汇总，不进入异常明细；E06 自 2026-07-20 起按行进入异常明细（与 E07 对称）
Private Sub TestAnomalyOutput_FilterSummaryOnlyCodes()
    Dim anomalies(1 To 4) As AnomalyRow

    anomalies(1).ErrorCode = ERR_E06
    anomalies(1).SourceTable = SOURCE_RETURN_TABLE
    anomalies(1).ExcelRowNum = 2
    anomalies(1).ShipmentNo = "SF_OB03"
    anomalies(1).WMSOrderNo = "TK_OB03"
    anomalies(1).SKU = "H_OB03"
    anomalies(1).FieldName = "物流单号"
    anomalies(1).RawValue = "SF_OB03"
    anomalies(1).Reason = "物流单号仅存在于退单表"
    anomalies(2).ErrorCode = ERR_E08
    anomalies(3).ErrorCode = ERR_E11
    anomalies(4).ErrorCode = ERR_E01
    anomalies(4).SourceTable = SOURCE_RETURN_TABLE
    anomalies(4).ExcelRowNum = 2
    anomalies(4).ShipmentNo = "SF_OB03"
    anomalies(4).WMSOrderNo = "TK_OB03"
    anomalies(4).SKU = "H_OB03"
    anomalies(4).FieldName = "数量"
    anomalies(4).RawValue = "abc"
    anomalies(4).Reason = "数量非法"

    Dim rows() As OutputRow
    rows = BuildAnomalyOutputRows(anomalies)

    AssertEqualLong "TC-OB03 异常明细保留E06+E01两行", 2, OBT_OutputRowCount(rows)
    AssertEqualString "TC-OB03 第1行保留错误码=E06", ERR_E06, OBT_RowColText(rows(1), 8)
    AssertEqualString "TC-OB03 第2行保留错误码=E01", ERR_E01, OBT_RowColText(rows(2), 8)
End Sub

' TC-OB04：调试日志输入空数组时返回空数组（不抛异常）
Private Sub TestDebugLogRows_EmptyEvents()
    Dim events() As AllocationEvent
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim rows() As OutputRow
    rows = BuildDebugLogRows(events, cfg)

    AssertEqualLong "TC-OB04 空日志返回0行", 0, OBT_OutputRowCount(rows)
End Sub

' TC-OB07：调试日志级别=关闭时不输出任何数据行
Private Sub TestDebugLogRows_OffLevel()
    Dim events(1 To 1) As AllocationEvent
    events(1).ShipmentNo = "SF_OB07"
    events(1).IsFinalResult = True

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_OFF

    Dim rows() As OutputRow
    rows = BuildDebugLogRows(events, cfg)

    AssertEqualLong "TC-OB07 关闭级别返回0行", 0, OBT_OutputRowCount(rows)
End Sub

' TC-OB08：简版仅保留 IsFinalResult=True 的事件
Private Sub TestDebugLogRows_SimpleFilter()
    Dim events(1 To 3) As AllocationEvent
    events(1).ShipmentNo = "SF_OB08"
    events(1).IsFinalResult = True
    events(2).ShipmentNo = "SF_OB08"
    events(2).IsFinalResult = False
    events(3).ShipmentNo = "SF_OB08"
    events(3).IsFinalResult = True

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_SIMPLE

    Dim rows() As OutputRow
    rows = BuildDebugLogRows(events, cfg)

    AssertEqualLong "TC-OB08 简版仅2行", 2, OBT_OutputRowCount(rows)
    AssertEqualString "TC-OB08 第1行物流单号", "SF_OB08", OBT_RowColText(rows(1), 1)
    AssertEqualString "TC-OB08 第2行物流单号", "SF_OB08", OBT_RowColText(rows(2), 1)
End Sub

' TC-OB09：调试日志映射为 19 列
Private Sub TestDebugLogRows_NineteenColumns()
    Dim events(1 To 1) As AllocationEvent
    With events(1)
        .ShipmentNo = "SF_OB09"
        .SKU = "H_OB09"
        .WMSOrderNo = "TK_OB09"
        .LineNo = "00001"
        .DemandD = 5
        .ProcessOrder = "1"
        .DynamicNextMinQty = ""
        .CandidateQCCount = "2"
        .ExcludedQCList = ""
        .StrategyUsed = "策略一"
        .UsedQC = QC_ZP
        .QCBefore = "-"
        .QCAfter = "-"
        .LotExpiryComboCount = "1"
        .IsBacktrackRetry = "否"
        .BacktrackNo = 0
        .LineStatus = STATUS_BATCH_IMPORT
        .ErrorCode = ""
        .FailSubType = ""
        .IsFinalResult = True
        .IsRevoked = False
    End With

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_SIMPLE

    Dim rows() As OutputRow
    rows = BuildDebugLogRows(events, cfg)

    AssertEqualLong "TC-OB09 输出1行", 1, OBT_OutputRowCount(rows)
    AssertEqualLong "TC-OB09 列数=19", 19, UBound(rows(1).Values) - LBound(rows(1).Values) + 1
    AssertEqualString "TC-OB09 第5列D", "5", OBT_RowColText(rows(1), 5)
    AssertEqualString "TC-OB09 第10列策略", "策略一", OBT_RowColText(rows(1), 10)
    AssertEqualString "TC-OB09 第11列分配QC", QC_ZP, OBT_RowColText(rows(1), 11)
    AssertEqualString "TC-OB09 第19列失败子类型", "", OBT_RowColText(rows(1), 19)
End Sub

' TC-OB05：FinalResult 明细映射到输出列
Private Sub TestDetailRows_FieldMapping()
    Dim finalResult As Object
    Set finalResult = CreateObject("Scripting.Dictionary")

    finalResult.Add "DetailCount", CLng(1)
    finalResult.Add "Detail_1_ShipmentNo", "SF_OB05"
    finalResult.Add "Detail_1_WMSOrderNo", "TK_OB05"
    finalResult.Add "Detail_1_SKU", "H_OB05"
    finalResult.Add "Detail_1_LineNo", "00001"
    finalResult.Add "Detail_1_OrderQty", CLng(3)
    finalResult.Add "Detail_1_QC", QC_ZP
    finalResult.Add "Detail_1_LotNo", "LA05"
    finalResult.Add "Detail_1_Expiry", "2030/05/01"
    finalResult.Add "Detail_1_AllocQty", CLng(3)
    finalResult.Add "Detail_1_LineStatus", STATUS_BATCH_IMPORT
    finalResult.Add "Detail_1_WMSOrderStatus", STATUS_BATCH_IMPORT

    Dim rows() As OutputRow
    rows = BuildDetailRows(finalResult)

    AssertEqualLong "TC-OB05 明细输出行数=1", 1, OBT_OutputRowCount(rows)
    AssertEqualString "TC-OB05 第1列物流单号", "SF_OB05", OBT_RowColText(rows(1), 1)
    AssertEqualString "TC-OB05 第2列WMS退单号", "TK_OB05", OBT_RowColText(rows(1), 2)
    AssertEqualLong "TC-OB05 第9列分配数量", 3, OBT_RowColLong(rows(1), 9)
    AssertEqualString "TC-OB05 第11列退单号状态", STATUS_BATCH_IMPORT, OBT_RowColText(rows(1), 11)
End Sub

' TC-OB06：运行历史行（Dry Run）字段（20 列布局：17 需求字段 + 3 配置快照）
Private Sub TestRunHistoryRow_DryRunFields()
    Dim stats As RunStats
    stats.InputReturnRows = 5
    stats.InputInventoryRows = 7
    stats.InputShipmentCount = 3
    stats.ValidationFailCount = 2
    stats.AllocSuccessCount = 0
    stats.AllocFailCount = 0
    stats.TotalBacktrackCount = 0
    stats.MaxGroupBacktrack = 0

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_SIMPLE
    cfg.MaxBacktrackCount = 88
    cfg.LotCaseSensitive = False
    cfg.NoExpirySentinel = "2099/01/01"

    Dim row As OutputRow
    row = BuildRunHistoryRow(stats, cfg, True)

    AssertEqualString "TC-OB06 第3列运行类型", "Dry Run", OBT_RowColText(row, 3)
    AssertEqualLong "TC-OB06 第4列退单输入行数", 5, OBT_RowColLong(row, 4)
    AssertEqualLong "TC-OB06 第6列物流单号数", 3, OBT_RowColLong(row, 6)
    AssertEqualLong "TC-OB06 第10列校验失败数", 2, OBT_RowColLong(row, 10)
    AssertEqualString "TC-OB06 第16列日志级别", DEBUG_LEVEL_SIMPLE, OBT_RowColText(row, 16)
    AssertEqualLong "TC-OB06 第18列最大回溯次数", 88, OBT_RowColLong(row, 18)
    AssertEqualString "TC-OB06 第20列哨兵值", "2099/01/01", OBT_RowColText(row, 20)
End Sub

Private Function OBT_OutputRowCount(ByRef rows() As OutputRow) As Long
    On Error GoTo EmptyArr
    OBT_OutputRowCount = UBound(rows) - LBound(rows) + 1
    Exit Function
EmptyArr:
    OBT_OutputRowCount = 0
End Function

Private Function OBT_RowColText(ByRef row As OutputRow, ByVal colIndex As Long) As String
    If Not IsArray(row.Values) Then Exit Function
    OBT_RowColText = CStr(row.Values(colIndex))
End Function

Private Function OBT_RowColLong(ByRef row As OutputRow, ByVal colIndex As Long) As Long
    If Not IsArray(row.Values) Then Exit Function
    OBT_RowColLong = CLng(row.Values(colIndex))
End Function


' =============================================================================
' M14 测试用例（UT-ExcelOutput）
' =============================================================================

Public Sub RunExcelOutputTests()
    ' M14 Excel 写入模块测试。
    ' 覆盖：
    '   TC-EO01 清空输出表：保留表头，不清空输入/配置/运行历史，删除调试日志分表
    '   TC-EO02 WriteSheet 列顺序与数据映射
    '   TC-EO03 WriteSheet 空数组仅保留表头
    '   TC-EO04 WriteDebugLog 超阈值自动分表且编号连续
    '   TC-EO05 调试日志关闭时不写数据且不保留分表
    '   TC-EO06 受保护工作表时中止并提示
    '   TC-EO07 AppendRunHistory 追加写入不覆盖旧记录
    BeginSuite "M14 Excel Output Tests"

    On Error GoTo CleanFail

    TestExcelOutput_ClearOutputSheets
    TestExcelOutput_WriteSheetMapping
    TestExcelOutput_WriteSheetEmptyRows
    TestExcelOutput_WriteDebugLogSplit
    TestExcelOutput_WriteDebugLogOff
    TestExcelOutput_ProtectedSheetAbort
    TestExcelOutput_AppendRunHistory

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M14 Excel Output Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-EO01：清空输出表时保留表头，不影响输入/配置/运行历史；调试分表会被删除。
Private Sub TestExcelOutput_ClearOutputSheets()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet
    Dim wsAnomaly As Worksheet
    Dim wsDebug As Worksheet
    Dim wsHistory As Worksheet
    Dim wsInputReturn As Worksheet
    Dim wsInputInv As Worksheet
    Dim wsInputCfg As Worksheet
    Dim wsDebug2 As Worksheet

    Set wsSummary = wb.Worksheets("分配状态汇总表")
    Set wsDetail = wb.Worksheets("成功分配明细表")
    Set wsAnomaly = wb.Worksheets("数据异常明细表")
    Set wsDebug = wb.Worksheets("调试日志")
    Set wsHistory = wb.Worksheets("运行历史记录表")
    Set wsInputReturn = wb.Worksheets("输入_退单表")
    Set wsInputInv = wb.Worksheets("输入_质检库存表")
    Set wsInputCfg = wb.Worksheets("输入_配置")
    Set wsDebug2 = wb.Worksheets("调试日志_2")

    wsSummary.Cells(1, 1).Value = "汇总头"
    wsSummary.Cells(2, 1).Value = "汇总旧数据"
    wsDetail.Cells(1, 1).Value = "明细头"
    wsDetail.Cells(2, 1).Value = "明细旧数据"
    wsAnomaly.Cells(1, 1).Value = "异常头"
    wsAnomaly.Cells(2, 1).Value = "异常旧数据"
    wsDebug.Cells(1, 1).Value = "日志头"
    wsDebug.Cells(2, 1).Value = "日志旧数据"
    wsDebug2.Cells(1, 1).Value = "日志分表头"
    wsDebug2.Cells(2, 1).Value = "日志分表旧数据"

    wsHistory.Cells(1, 1).Value = "历史头"
    wsHistory.Cells(2, 1).Value = "历史旧数据"
    wsInputReturn.Cells(1, 1).Value = "输入退单头"
    wsInputReturn.Cells(2, 1).Value = "输入退单数据"
    wsInputInv.Cells(1, 1).Value = "输入库存头"
    wsInputInv.Cells(2, 1).Value = "输入库存数据"
    wsInputCfg.Cells(1, 1).Value = "输入配置头"
    wsInputCfg.Cells(2, 1).Value = "输入配置数据"

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    ClearOutputSheets wb, cfg

    AssertEqualString "TC-EO01 汇总表头保留", "汇总头", CStr(wsSummary.Cells(1, 1).Value)
    AssertEqualString "TC-EO01 汇总数据被清空", vbNullString, CStr(wsSummary.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 明细数据被清空", vbNullString, CStr(wsDetail.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 异常数据被清空", vbNullString, CStr(wsAnomaly.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 调试主表数据被清空", vbNullString, CStr(wsDebug.Cells(2, 1).Value)

    AssertFalse "TC-EO01 调试日志分表被删除", EOT_SheetExists(wb, "调试日志_2")
    AssertEqualString "TC-EO01 运行历史不清空", "历史旧数据", CStr(wsHistory.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 输入退单表不清空", "输入退单数据", CStr(wsInputReturn.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 输入库存表不清空", "输入库存数据", CStr(wsInputInv.Cells(2, 1).Value)
    AssertEqualString "TC-EO01 输入配置表不清空", "输入配置数据", CStr(wsInputCfg.Cells(2, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO01 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO02：WriteSheet 按 headers 顺序写入，列映射不偏移。
Private Sub TestExcelOutput_WriteSheetMapping()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = wb.Worksheets("分配状态汇总表")

    Dim headers(1 To 4) As String
    headers(1) = "物流单号"
    headers(2) = "WMS退单号"
    headers(3) = "退单号状态"
    headers(4) = "原因"

    Dim rows(1 To 2) As OutputRow
    rows(1) = EOT_CreateOutputRow("SF_EO02_A", "TK_EO02_A", STATUS_BATCH_IMPORT, vbNullString)
    rows(2) = EOT_CreateOutputRow("SF_EO02_B", "TK_EO02_B", STATUS_UNALLOCATED, "E10 - 回溯超限")

    WriteSheet ws, rows, headers

    AssertEqualString "TC-EO02 表头第1列", "物流单号", CStr(ws.Cells(1, 1).Value)
    AssertEqualString "TC-EO02 表头第4列", "原因", CStr(ws.Cells(1, 4).Value)
    AssertEqualString "TC-EO02 第2行第1列", "SF_EO02_A", CStr(ws.Cells(2, 1).Value)
    AssertEqualString "TC-EO02 第2行第3列", STATUS_BATCH_IMPORT, CStr(ws.Cells(2, 3).Value)
    AssertEqualString "TC-EO02 第3行第2列", "TK_EO02_B", CStr(ws.Cells(3, 2).Value)
    AssertEqualString "TC-EO02 第3行第4列", "E10 - 回溯超限", CStr(ws.Cells(3, 4).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO02 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO03：WriteSheet 输入空数组时仅写表头，旧数据会被清空。
Private Sub TestExcelOutput_WriteSheetEmptyRows()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = wb.Worksheets("分配状态汇总表")
    ws.Cells(2, 1).Value = "旧数据"

    Dim headers(1 To 4) As String
    headers(1) = "物流单号"
    headers(2) = "WMS退单号"
    headers(3) = "退单号状态"
    headers(4) = "原因"

    Dim rows() As OutputRow
    WriteSheet ws, rows, headers

    AssertEqualString "TC-EO03 表头第2列", "WMS退单号", CStr(ws.Cells(1, 2).Value)
    AssertEqualString "TC-EO03 第2行被清空", vbNullString, CStr(ws.Cells(2, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO03 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO04：WriteDebugLog 超过阈值时自动分表，且分表编号连续。
Private Sub TestExcelOutput_WriteDebugLogSplit()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_DETAIL
    cfg.DetailedLogLimit = 2

    Dim rows(1 To 5) As OutputRow
    rows(1) = EOT_CreateOutputRow("SF_EO04_1", "SKU1", "TK1", "00001", "Final", "ZP", "策略一", 0, "False", "True", DEBUG_LEVEL_DETAIL)
    rows(2) = EOT_CreateOutputRow("SF_EO04_2", "SKU2", "TK2", "00002", "Final", "QC", "策略二", 1, "False", "True", DEBUG_LEVEL_DETAIL)
    rows(3) = EOT_CreateOutputRow("SF_EO04_3", "SKU3", "TK3", "00003", "Try", "NG", "策略三", 2, "True", "False", DEBUG_LEVEL_DETAIL)
    rows(4) = EOT_CreateOutputRow("SF_EO04_4", "SKU4", "TK4", "00004", "Final", "ZP", "策略一", 2, "False", "True", DEBUG_LEVEL_DETAIL)
    rows(5) = EOT_CreateOutputRow("SF_EO04_5", "SKU5", "TK5", "00005", "Final", "QC", "策略二", 3, "False", "True", DEBUG_LEVEL_DETAIL)

    WriteDebugLog wb, rows, cfg

    AssertTrue "TC-EO04 主日志表存在", EOT_SheetExists(wb, "调试日志")
    AssertTrue "TC-EO04 分表2存在", EOT_SheetExists(wb, "调试日志_2")
    AssertTrue "TC-EO04 分表3存在", EOT_SheetExists(wb, "调试日志_3")

    AssertEqualString "TC-EO04 主表首行数据", "SF_EO04_1", CStr(wb.Worksheets("调试日志").Cells(2, 1).Value)
    AssertEqualString "TC-EO04 分表2首行数据", "SF_EO04_3", CStr(wb.Worksheets("调试日志_2").Cells(2, 1).Value)
    AssertEqualString "TC-EO04 分表3首行数据", "SF_EO04_5", CStr(wb.Worksheets("调试日志_3").Cells(2, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO04 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO06：工作表被保护时必须中止并给出可定位提示。
Private Sub TestExcelOutput_ProtectedSheetAbort()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = wb.Worksheets("分配状态汇总表")
    ws.Protect Password:="123"

    Dim headers(1 To 2) As String
    headers(1) = "A"
    headers(2) = "B"

    Dim rows(1 To 1) As OutputRow
    rows(1) = EOT_CreateOutputRow("x", "y")

    Dim raised As Boolean
    On Error Resume Next
    WriteSheet ws, rows, headers
    raised = (Err.Number <> 0)
    If raised Then
        AssertContains "TC-EO06 错误信息含保护提示", Err.Description, "受保护"
    End If
    On Error GoTo CleanFail

    AssertTrue "TC-EO06 受保护工作表应抛错", raised
    ws.Unprotect Password:="123"

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO06 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO05：调试日志级别关闭时，主日志表仅保留表头且不保留分表。
Private Sub TestExcelOutput_WriteDebugLogOff()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    wb.Worksheets("调试日志").Cells(1, 1).Value = "旧日志头"
    wb.Worksheets("调试日志").Cells(2, 1).Value = "旧日志数据"
    wb.Worksheets("调试日志_2").Cells(2, 1).Value = "旧分表数据"

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_OFF
    cfg.DetailedLogLimit = 2

    Dim rows(1 To 1) As OutputRow
    rows(1) = EOT_CreateOutputRow("SF_EO05", "SKU", "TK", "00001", "Try", "ZP", "策略一", 0, "False", "True", DEBUG_LEVEL_DETAIL)

    WriteDebugLog wb, rows, cfg

    AssertFalse "TC-EO05 分表被删除", EOT_SheetExists(wb, "调试日志_2")
    AssertEqualString "TC-EO05 主日志首列表头", "物流单号", CStr(wb.Worksheets("调试日志").Cells(1, 1).Value)
    AssertEqualString "TC-EO05 主日志无数据", vbNullString, CStr(wb.Worksheets("调试日志").Cells(2, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO05 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-EO07：AppendRunHistory 追加到末尾，不覆盖既有历史；第 1 列运行编号自增。
Private Sub TestExcelOutput_AppendRunHistory()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = wb.Worksheets("运行历史记录表")
    ws.Cells(1, 1).Value = "运行编号"
    ws.Cells(1, 3).Value = "运行类型"
    ws.Cells(1, 4).Value = "输入：退单表行数"
    ' 旧历史行：编号 1，Full Run
    ws.Cells(2, 1).Value = 1
    ws.Cells(2, 3).Value = "Full Run"
    ws.Cells(2, 4).Value = 9

    Dim row As OutputRow
    row = EOT_CreateOutputRow("", "", "Dry Run", 5)
    AppendRunHistory ws, row

    AssertEqualString "TC-EO07 旧历史保留", "Full Run", CStr(ws.Cells(2, 3).Value)
    AssertEqualString "TC-EO07 新历史追加到第3行", "Dry Run", CStr(ws.Cells(3, 3).Value)
    AssertEqualLong "TC-EO07 新历史退单行数列", 5, CLng(ws.Cells(3, 4).Value)
    AssertEqualLong "TC-EO07 新历史运行编号自增=2", 2, CLng(ws.Cells(3, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-EO07 异常：" & Err.Description
    Resume CleanExit
End Sub

Private Function EOT_CreateWorkbookWithNamedSheets() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)

    wb.Worksheets(1).Name = "分配状态汇总表"
    EOT_EnsureSheet wb, "成功分配明细表"
    EOT_EnsureSheet wb, "数据异常明细表"
    EOT_EnsureSheet wb, "调试日志"
    EOT_EnsureSheet wb, "调试日志_2"
    EOT_EnsureSheet wb, "运行历史记录表"
    EOT_EnsureSheet wb, "输入_退单表"
    EOT_EnsureSheet wb, "输入_质检库存表"
    EOT_EnsureSheet wb, "输入_配置"

    Set EOT_CreateWorkbookWithNamedSheets = wb
End Function

Private Sub EOT_EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String)
    If EOT_SheetExists(wb, sheetName) Then Exit Sub
    wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    wb.Worksheets(wb.Worksheets.Count).Name = sheetName
End Sub

Private Function EOT_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    EOT_SheetExists = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function EOT_CreateOutputRow(ParamArray values() As Variant) As OutputRow
    Dim row As OutputRow
    Dim cols() As Variant
    Dim i As Long

    ReDim cols(1 To UBound(values) - LBound(values) + 1)
    For i = LBound(values) To UBound(values)
        cols(i - LBound(values) + 1) = values(i)
    Next i

    row.Values = cols
    EOT_CreateOutputRow = row
End Function

Private Sub EOT_CloseWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub


' =============================================================================
' M13→M14 集成测试（IT-M13M14）
' =============================================================================
' 目的：验证 M13 输出的列顺序与 M14 写入的表头语义完全对齐。
' 纯单元测试无法覆盖这个风险——M13 的每列位置由参数顺序决定，
' M14 的表头由调用方传入，两者必须由集成测试贯穿才能确认"值写到对的列"。

Public Sub RunM13M14IntegrationTests()
    ' 覆盖：
    '   TC-IT01 汇总表列映射：M13 BuildSummaryRows → WriteSheet，验列名与数据对齐
    '   TC-IT02 明细表列映射：M13 BuildDetailRows → WriteSheet，验 QC/分配数量/退单号状态
    '   TC-IT03 异常明细列映射：M13 BuildAnomalyOutputRows → WriteSheet，验错误码在第8列
    '   TC-IT04 运行历史列映射：M13 BuildRunHistoryRow → AppendRunHistory，验回溯次数/日志级别
    '   TC-IT05 二次运行回归：汇总清空重写 + 历史追加同时正确
    BeginSuite "M13->M14 Integration Tests"

    On Error GoTo CleanFail

    TestIT_SummaryColumnMapping
    TestIT_DetailColumnMapping
    TestIT_AnomalyColumnMapping
    TestIT_HistoryColumnMapping
    TestIT_SecondRunRegression

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M13->M14 Integration Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-IT01：汇总表写入后，每列数据与对应列名语义一致。
Private Sub TestIT_SummaryColumnMapping()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim statusMap(1 To 1) As WMSStatusEntry
    statusMap(1).ShipmentNo = "SF_IT01"
    statusMap(1).WMSOrderNo = "TK_IT01"
    statusMap(1).Status = STATUS_UNALLOCATED
    statusMap(1).Reason = "E10 - 回溯超限"

    Dim rows() As OutputRow
    rows = BuildSummaryRows(statusMap, False)

    Dim headers(1 To 4) As String
    headers(1) = "物流单号"
    headers(2) = "WMS退单号"
    headers(3) = "退单号状态"
    headers(4) = "原因"

    Dim ws As Worksheet
    Set ws = wb.Worksheets("分配状态汇总表")
    WriteSheet ws, rows, headers

    AssertEqualString "TC-IT01 表头第1列=物流单号", "物流单号", CStr(ws.Cells(1, 1).Value)
    AssertEqualString "TC-IT01 表头第3列=退单号状态", "退单号状态", CStr(ws.Cells(1, 3).Value)
    AssertEqualString "TC-IT01 数据第1列=物流单号值", "SF_IT01", CStr(ws.Cells(2, 1).Value)
    AssertEqualString "TC-IT01 数据第2列=WMS退单号值", "TK_IT01", CStr(ws.Cells(2, 2).Value)
    AssertEqualString "TC-IT01 数据第3列=状态值", STATUS_UNALLOCATED, CStr(ws.Cells(2, 3).Value)
    AssertEqualString "TC-IT01 数据第4列=原因值", "E10 - 回溯超限", CStr(ws.Cells(2, 4).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-IT01 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-IT02：明细表写入后，QC/分配数量/退单号状态在正确列。
Private Sub TestIT_DetailColumnMapping()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim finalResult As Object
    Set finalResult = CreateObject("Scripting.Dictionary")
    finalResult.Add "DetailCount", CLng(1)
    finalResult.Add "Detail_1_ShipmentNo", "SF_IT02"
    finalResult.Add "Detail_1_WMSOrderNo", "TK_IT02"
    finalResult.Add "Detail_1_SKU", "H_IT02"
    finalResult.Add "Detail_1_LineNo", "00001"
    finalResult.Add "Detail_1_OrderQty", CLng(3)
    finalResult.Add "Detail_1_QC", QC_ZP
    finalResult.Add "Detail_1_LotNo", "LB02"
    finalResult.Add "Detail_1_Expiry", "2030/06/01"
    finalResult.Add "Detail_1_AllocQty", CLng(3)
    finalResult.Add "Detail_1_LineStatus", STATUS_BATCH_IMPORT
    finalResult.Add "Detail_1_WMSOrderStatus", STATUS_BATCH_IMPORT

    Dim rows() As OutputRow
    rows = BuildDetailRows(finalResult)

    Dim headers(1 To 11) As String
    headers(1) = "物流单号"
    headers(2) = "WMS退单号"
    headers(3) = "SKU"
    headers(4) = "行号"
    headers(5) = "退单数量"
    headers(6) = "QC情况"
    headers(7) = "批号"
    headers(8) = "效期"
    headers(9) = "分配数量"
    headers(10) = "行状态"
    headers(11) = "退单号状态"

    Dim ws As Worksheet
    Set ws = wb.Worksheets("成功分配明细表")
    WriteSheet ws, rows, headers

    AssertEqualString "TC-IT02 表头第6列=QC情况", "QC情况", CStr(ws.Cells(1, 6).Value)
    AssertEqualString "TC-IT02 表头第9列=分配数量", "分配数量", CStr(ws.Cells(1, 9).Value)
    AssertEqualString "TC-IT02 表头第11列=退单号状态", "退单号状态", CStr(ws.Cells(1, 11).Value)
    AssertEqualString "TC-IT02 第4列行号按文本写入", "@", CStr(ws.Columns(4).NumberFormat)
    AssertEqualString "TC-IT02 第8列效期按文本写入", "@", CStr(ws.Columns(8).NumberFormat)
    AssertEqualString "TC-IT02 数据第4列行号保留前导零", "00001", CStr(ws.Cells(2, 4).Value)
    AssertEqualString "TC-IT02 数据第8列效期保留前导零", "2030/06/01", CStr(ws.Cells(2, 8).Value)
    AssertEqualString "TC-IT02 数据第6列=QC值", QC_ZP, CStr(ws.Cells(2, 6).Value)
    AssertEqualLong "TC-IT02 数据第9列=分配数量值", 3, CLng(ws.Cells(2, 9).Value)
    AssertEqualString "TC-IT02 数据第10列=行状态值", STATUS_BATCH_IMPORT, CStr(ws.Cells(2, 10).Value)
    AssertEqualString "TC-IT02 数据第11列=退单号状态值", STATUS_BATCH_IMPORT, CStr(ws.Cells(2, 11).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-IT02 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-IT03：异常明细表写入后，错误码在第8列，原因说明在第9列。
Private Sub TestIT_AnomalyColumnMapping()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim anomalies(1 To 1) As AnomalyRow
    anomalies(1).SourceTable = SOURCE_RETURN_TABLE
    anomalies(1).ExcelRowNum = 3
    anomalies(1).ShipmentNo = "SF_IT03"
    anomalies(1).WMSOrderNo = "TK_IT03"
    anomalies(1).SKU = "H_IT03"
    anomalies(1).FieldName = "行号"
    anomalies(1).RawValue = "1"
    anomalies(1).ErrorCode = ERR_E01
    anomalies(1).Reason = "行号格式不符"

    Dim rows() As OutputRow
    rows = BuildAnomalyOutputRows(anomalies)

    Dim headers(1 To 9) As String
    headers(1) = "来源表"
    headers(2) = "原始行号"
    headers(3) = "物流单号"
    headers(4) = "WMS退单号"
    headers(5) = "SKU"
    headers(6) = "异常字段名"
    headers(7) = "原始值"
    headers(8) = "错误码"
    headers(9) = "原因说明"

    Dim ws As Worksheet
    Set ws = wb.Worksheets("数据异常明细表")
    WriteSheet ws, rows, headers

    AssertEqualString "TC-IT03 表头第8列=错误码", "错误码", CStr(ws.Cells(1, 8).Value)
    AssertEqualString "TC-IT03 数据第1列=来源表值", SOURCE_RETURN_TABLE, CStr(ws.Cells(2, 1).Value)
    AssertEqualLong "TC-IT03 数据第2列=原始行号", 3, CLng(ws.Cells(2, 2).Value)
    AssertEqualString "TC-IT03 数据第8列=错误码值", ERR_E01, CStr(ws.Cells(2, 8).Value)
    AssertEqualString "TC-IT03 数据第9列=原因值", "行号格式不符", CStr(ws.Cells(2, 9).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-IT03 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-IT04：运行历史追加后，总回溯次数在第7列、调试日志级别在第9列。
Private Sub TestIT_HistoryColumnMapping()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim stats As RunStats
    stats.InputReturnRows = 4
    stats.InputInventoryRows = 6
    stats.ValidationFailCount = 1
    stats.AllocSuccessCount = 2
    stats.AllocFailCount = 1
    stats.TotalBacktrackCount = 5
    stats.MaxGroupBacktrack = 3

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()
    cfg.DebugLogLevel = DEBUG_LEVEL_SIMPLE

    Dim row As OutputRow
    row = BuildRunHistoryRow(stats, cfg, False)

    Dim ws As Worksheet
    Set ws = wb.Worksheets("运行历史记录表")
    ws.Cells(1, 1).Value = "运行编号"
    ws.Cells(1, 2).Value = "运行时间"
    ws.Cells(1, 3).Value = "运行类型"
    ws.Cells(1, 4).Value = "输入：退单表行数"
    ws.Cells(1, 5).Value = "输入：质检库存表行数"
    ws.Cells(1, 6).Value = "输入：物流单号数"
    ws.Cells(1, 7).Value = "校验耗时（秒）"
    ws.Cells(1, 8).Value = "分配耗时（秒）"
    ws.Cells(1, 9).Value = "总耗时（秒）"
    ws.Cells(1, 10).Value = "校验失败物流单号数"
    ws.Cells(1, 11).Value = "分配成功物流单号数"
    ws.Cells(1, 12).Value = "分配失败物流单号数"
    ws.Cells(1, 13).Value = "错误码分布"
    ws.Cells(1, 14).Value = "总回溯次数"
    ws.Cells(1, 15).Value = "最大单组回溯次数"
    ws.Cells(1, 16).Value = "调试日志级别"
    ws.Cells(1, 17).Value = "备注"
    ws.Cells(1, 18).Value = "最大回溯次数"
    ws.Cells(1, 19).Value = "批号比较模式"
    ws.Cells(1, 20).Value = "无保质期哨兵值"

    AppendRunHistory ws, row

    AssertEqualLong "TC-IT04 数据第1列=运行编号", 1, CLng(ws.Cells(2, 1).Value)
    AssertEqualString "TC-IT04 数据第3列=Full Run", "Full Run", CStr(ws.Cells(2, 3).Value)
    AssertEqualLong "TC-IT04 数据第4列=退单表行数", 4, CLng(ws.Cells(2, 4).Value)
    AssertEqualLong "TC-IT04 数据第14列=总回溯次数", 5, CLng(ws.Cells(2, 14).Value)
    AssertEqualLong "TC-IT04 数据第15列=最大单组回溯", 3, CLng(ws.Cells(2, 15).Value)
    AssertEqualString "TC-IT04 数据第16列=日志级别", DEBUG_LEVEL_SIMPLE, CStr(ws.Cells(2, 16).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-IT04 异常：" & Err.Description
    Resume CleanExit
End Sub

' TC-IT05：二次运行回归——汇总表清空重写，历史记录追加保留，同时验证。
Private Sub TestIT_SecondRunRegression()
    Dim wb As Workbook
    Set wb = EOT_CreateWorkbookWithNamedSheets()

    On Error GoTo CleanFail

    Dim headers(1 To 4) As String
    headers(1) = "物流单号"
    headers(2) = "WMS退单号"
    headers(3) = "退单号状态"
    headers(4) = "原因"

    Dim wsSummary As Worksheet
    Set wsSummary = wb.Worksheets("分配状态汇总表")

    Dim wsHistory As Worksheet
    Set wsHistory = wb.Worksheets("运行历史记录表")
    wsHistory.Cells(1, 1).Value = "运行编号"
    wsHistory.Cells(1, 3).Value = "运行类型"
    wsHistory.Cells(1, 4).Value = "输入：退单表行数"

    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    ' ====== 第一次运行：写入2条汇总，追加1条历史 ======
    Dim statusMap1(1 To 2) As WMSStatusEntry
    statusMap1(1).ShipmentNo = "SF_IT05_R1_A"
    statusMap1(1).WMSOrderNo = "TK_IT05_R1_A"
    statusMap1(1).Status = STATUS_BATCH_IMPORT
    statusMap1(2).ShipmentNo = "SF_IT05_R1_B"
    statusMap1(2).WMSOrderNo = "TK_IT05_R1_B"
    statusMap1(2).Status = STATUS_UNALLOCATED
    statusMap1(2).Reason = "E09 - 分配路径穷尽"

    Dim rows1() As OutputRow
    rows1 = BuildSummaryRows(statusMap1, False)
    WriteSheet wsSummary, rows1, headers

    Dim stats1 As RunStats
    stats1.InputReturnRows = 2
    AppendRunHistory wsHistory, BuildRunHistoryRow(stats1, cfg, False)

    ' ====== 第二次运行：只写1条汇总，再追加1条历史 ======
    Dim statusMap2(1 To 1) As WMSStatusEntry
    statusMap2(1).ShipmentNo = "SF_IT05_R2"
    statusMap2(1).WMSOrderNo = "TK_IT05_R2"
    statusMap2(1).Status = STATUS_MANUAL

    Dim rows2() As OutputRow
    rows2 = BuildSummaryRows(statusMap2, False)
    WriteSheet wsSummary, rows2, headers

    Dim stats2 As RunStats
    stats2.InputReturnRows = 1
    AppendRunHistory wsHistory, BuildRunHistoryRow(stats2, cfg, False)

    ' ====== 验证汇总：旧数据已清除，只剩第二次的1行 ======
    AssertEqualString "TC-IT05 汇总第2行=R2（R1已清除）", "SF_IT05_R2", CStr(wsSummary.Cells(2, 1).Value)
    AssertEqualString "TC-IT05 汇总第3行为空（R1_B已清除）", vbNullString, CStr(wsSummary.Cells(3, 1).Value)

    ' ====== 验证历史：两次运行的记录均保留 ======
    AssertEqualString "TC-IT05 历史第2行=第一次运行", "Full Run", CStr(wsHistory.Cells(2, 3).Value)
    AssertEqualString "TC-IT05 历史第3行=第二次运行追加", "Full Run", CStr(wsHistory.Cells(3, 3).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-IT05 异常：" & Err.Description
    Resume CleanExit
End Sub


' =============================================================================
' M12 测试用例（UT-PostCheck）
' =============================================================================

Public Sub RunPostValidateTests()
    ' M12 分配后校验模块测试。
    ' 覆盖：
    '   TC-PV01 成功分配 → 通过
    '   TC-PV02 某行分配量不等于退单量 → 失败
    '   TC-PV03 同一行使用两种 QC → 失败
    '   TC-PV04 整单回滚物流单号 → 不参与后校验
'   TC-PV05 成功明细关键字段与输入不一致 → 失败
'   TC-PV06 退单号状态与行状态聚合不一致 → 失败
    BeginSuite "M12 Post Validate Tests"

    On Error GoTo CleanFail

    TestPostValidate_Success
    TestPostValidate_QtyMismatch
    TestPostValidate_TwoQcsSameLine
    TestPostValidate_SkipRollbackShipment
    TestPostValidate_DataMismatch
    TestPostValidate_WMSStatusMismatch

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M12 Post Validate Tests 执行异常：" & Err.Description
    FinishSuite
End Sub

' TC-PV01：分配明细总量等于退单量，且同一行只使用一种 QC → 通过
Private Sub TestPostValidate_Success()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV01", "TK_PV01", "H_PV01", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    PV_AddDetail finalResult, 1, "SF_PV01", "TK_PV01", "H_PV01", "00001", 5, QC_ZP
    PV_AddSummary finalResult, 1, "SF_PV01", "TK_PV01", STATUS_BATCH_IMPORT, vbNullString

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertFalse "TC-PV01 后校验应通过", CBool(result("HasFailures"))
    AssertEqualLong "TC-PV01 问题数=0", 0, CLng(result("IssueCount"))
End Sub

' TC-PV02：同一退单行的分配量合计少于退单量 → 失败
Private Sub TestPostValidate_QtyMismatch()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV02", "TK_PV02", "H_PV02", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    PV_AddDetail finalResult, 1, "SF_PV02", "TK_PV02", "H_PV02", "00001", 4, QC_ZP
    PV_AddSummary finalResult, 1, "SF_PV02", "TK_PV02", STATUS_BATCH_IMPORT, vbNullString

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertTrue "TC-PV02 数量不守恒应失败", CBool(result("HasFailures"))
    AssertEqualString "TC-PV02 错误码", "POST_QTY_MISMATCH", CStr(result("Issue_1_Code"))
End Sub

' TC-PV03：同一退单行被拆成两条明细，但使用了两种 QC → 失败
Private Sub TestPostValidate_TwoQcsSameLine()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV03", "TK_PV03", "H_PV03", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    PV_AddDetail finalResult, 1, "SF_PV03", "TK_PV03", "H_PV03", "00001", 2, QC_ZP
    PV_AddDetail finalResult, 2, "SF_PV03", "TK_PV03", "H_PV03", "00001", 3, QC_QC
    PV_AddSummary finalResult, 1, "SF_PV03", "TK_PV03", STATUS_BATCH_IMPORT, vbNullString

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertTrue "TC-PV03 同行两种QC应失败", CBool(result("HasFailures"))
    AssertEqualString "TC-PV03 错误码", "POST_QC_MISMATCH", CStr(result("Issue_1_Code"))
End Sub

' TC-PV04：整单回滚只有失败汇总、没有成功明细，M12 应跳过该物流单号
Private Sub TestPostValidate_SkipRollbackShipment()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV04", "TK_PV04", "H_PV04", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    finalResult("SummaryCount") = CLng(1)
    finalResult.Add "Summary_1_ShipmentNo", "SF_PV04"
    finalResult.Add "Summary_1_WMSOrderNo", "TK_PV04"
    finalResult.Add "Summary_1_Status", STATUS_UNALLOCATED
    finalResult.Add "Summary_1_Reason", "E10 - 回溯超限"

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertFalse "TC-PV04 回滚单跳过后校验", CBool(result("HasFailures"))
    AssertEqualLong "TC-PV04 问题数=0", 0, CLng(result("IssueCount"))
End Sub

' TC-PV05：成功明细的退单数量与输入退单行不一致 → 数据完整性失败
Private Sub TestPostValidate_DataMismatch()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV05", "TK_PV05", "H_PV05", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    PV_AddDetail finalResult, 1, "SF_PV05", "TK_PV05", "H_PV05", "00001", 4, QC_ZP
    finalResult("Detail_1_OrderQty") = CLng(4)
    PV_AddSummary finalResult, 1, "SF_PV05", "TK_PV05", STATUS_BATCH_IMPORT, vbNullString

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertTrue "TC-PV05 数据完整性不一致应失败", CBool(result("HasFailures"))
    AssertEqualString "TC-PV05 错误码", "POST_QTY_MISMATCH", CStr(result("Issue_1_Code"))
    AssertTrue "TC-PV05 至少包含数据完整性问题", PV_ResultHasIssueCode(result, "POST_DATA_MISMATCH")
End Sub

' TC-PV06：有任意手工行时，WMS 退单号状态必须聚合为“手工操作”。
Private Sub TestPostValidate_WMSStatusMismatch()
    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = PV_MakeReturnLine("SF_PV06", "TK_PV06", "H_PV06", "00001", 5)

    Dim finalResult As Object
    Set finalResult = PV_CreateFinalResult()
    PV_AddDetail finalResult, 1, "SF_PV06", "TK_PV06", "H_PV06", "00001", 5, QC_ZP
    finalResult("Detail_1_LineStatus") = STATUS_MANUAL
    finalResult("Detail_1_WMSOrderStatus") = STATUS_BATCH_IMPORT
    PV_AddSummary finalResult, 1, "SF_PV06", "TK_PV06", STATUS_BATCH_IMPORT, vbNullString

    Dim result As Object
    Set result = ValidatePost(orders, finalResult)

    AssertTrue "TC-PV06 退单号状态聚合不一致应失败", CBool(result("HasFailures"))
    AssertTrue "TC-PV06 包含状态一致性问题", PV_ResultHasIssueCode(result, "POST_STATUS_MISMATCH")
End Sub

Private Function PV_CreateFinalResult() As Object
    Dim finalResult As Object
    Set finalResult = CreateObject("Scripting.Dictionary")
    finalResult.Add "SummaryCount", CLng(0)
    finalResult.Add "DetailCount", CLng(0)
    Set PV_CreateFinalResult = finalResult
End Function

Private Sub PV_AddSummary( _
    ByVal finalResult As Object, _
    ByVal summaryIndex As Long, _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal status As String, _
    ByVal reason As String)

    finalResult("SummaryCount") = summaryIndex
    finalResult.Add "Summary_" & summaryIndex & "_ShipmentNo", shipNo
    finalResult.Add "Summary_" & summaryIndex & "_WMSOrderNo", wmsOrderNo
    finalResult.Add "Summary_" & summaryIndex & "_Status", status
    finalResult.Add "Summary_" & summaryIndex & "_Reason", reason
End Sub

Private Function PV_ResultHasIssueCode(ByVal result As Object, ByVal issueCode As String) As Boolean
    Dim issueCount As Long
    If result Is Nothing Then Exit Function
    If result.Exists("IssueCount") Then issueCount = CLng(result("IssueCount"))

    Dim i As Long
    For i = 1 To issueCount
        If result.Exists("Issue_" & i & "_Code") Then
            If CStr(result("Issue_" & i & "_Code")) = issueCode Then
                PV_ResultHasIssueCode = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub PV_AddDetail( _
    ByVal finalResult As Object, _
    ByVal detailIndex As Long, _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String, _
    ByVal allocQty As Long, _
    ByVal qc As String)

    finalResult("DetailCount") = detailIndex
    finalResult.Add "Detail_" & detailIndex & "_ShipmentNo", shipNo
    finalResult.Add "Detail_" & detailIndex & "_WMSOrderNo", wmsOrderNo
    finalResult.Add "Detail_" & detailIndex & "_SKU", sku
    finalResult.Add "Detail_" & detailIndex & "_LineNo", lineNo
    finalResult.Add "Detail_" & detailIndex & "_OrderQty", allocQty
    finalResult.Add "Detail_" & detailIndex & "_QC", qc
    finalResult.Add "Detail_" & detailIndex & "_LotNo", "LA01"
    finalResult.Add "Detail_" & detailIndex & "_Expiry", "2029/01/01"
    finalResult.Add "Detail_" & detailIndex & "_AllocQty", allocQty
    finalResult.Add "Detail_" & detailIndex & "_LineStatus", STATUS_BATCH_IMPORT
    finalResult.Add "Detail_" & detailIndex & "_WMSOrderStatus", STATUS_BATCH_IMPORT
End Sub

Private Function PV_MakeReturnLine( _
    ByVal shipNo As String, _
    ByVal wmsOrderNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String, _
    ByVal qty As Long) As NormalizedReturnLine

    Dim line As NormalizedReturnLine
    line.ShipmentNo = shipNo
    line.WMSOrderNo = wmsOrderNo
    line.SKU = sku
    line.LineNo = lineNo
    line.Qty = qty
    line.LineNoValid = True
    line.QtyValid = True
    PV_MakeReturnLine = line
End Function


' =============================================================================
' IT-Runner：M15 运行编排模块测试（BuildRunStats 单元测试）
' =============================================================================
' 说明：RunValidationOnly / RunFullAllocation 依赖真实工作表，属于手动集成测试，
'       此处只对可纯函数化的 BuildRunStats 做自动化单元测试。
'
' TC-RUNNER-01：干跑模式 — shipmentResults 为空，分配字段全为 0
' TC-RUNNER-02：完整运行 — 全部成功，回溯次数正确累计
' TC-RUNNER-03：完整运行 — 混合成功/失败，AllocSuccessCount/AllocFailCount 分别计数
' TC-RUNNER-04：全空输入 — 不崩溃，所有字段 = 0
' TC-RUNNER-05：MaxGroupBacktrack — 多组回溯取最大值
' =============================================================================

Public Sub RunRunnerTests()
    ' M15 BuildRunStats 单元测试入口。
    BeginSuite "M15 Runner Tests (BuildRunStats)"

    On Error GoTo CleanFail

    TestRunner_BuildRunStats_DryRun
    TestRunner_BuildRunStats_AllSuccess
    TestRunner_BuildRunStats_MixedResult
    TestRunner_BuildRunStats_EmptyInput
    TestRunner_BuildRunStats_MaxBacktrack
    TestRunner_RerunOverwrite_TC43
    TestRunner_EmptyInputFullRun
    TestRunner_E02RawValueKeepsLeadingZeros

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] M15 Runner Tests 执行异常：" & Err.Description
    FinishSuite
End Sub


' TC-RUNNER-01：干跑模式。
' 传入空的 shipmentResults 数组，验证：
'   - InputReturnRows / InputInventoryRows 来自 orders/inventory 数组长度
'   - ValidationFailCount 来自 validationResult.FailedShipmentCount
'   - 分配相关字段（回溯次数、成功/失败数）全为 0
Private Sub TestRunner_BuildRunStats_DryRun()
    Dim validationResult As ValidationResult
    validationResult.HasFailures = True
    validationResult.FailedShipmentCount = 2

    ' 空的分配结果数组（干跑时不调用 M09）
    Dim emptyResults() As Object

    ' 虚构 3 行退单、5 行库存
    Dim orders(1 To 3) As NormalizedReturnLine
    Dim inventory(1 To 5) As NormalizedInventoryLine

    Dim stats As RunStats
    stats = BuildRunStats(validationResult, emptyResults, orders, inventory)

    AssertEqualLong "TC-RUNNER-01 InputReturnRows=3", 3, stats.InputReturnRows
    AssertEqualLong "TC-RUNNER-01 InputInventoryRows=5", 5, stats.InputInventoryRows
    AssertEqualLong "TC-RUNNER-01 ValidationFailCount=2", 2, stats.ValidationFailCount
    AssertEqualLong "TC-RUNNER-01 干跑TotalBacktrackCount=0", 0, stats.TotalBacktrackCount
    AssertEqualLong "TC-RUNNER-01 干跑MaxGroupBacktrack=0", 0, stats.MaxGroupBacktrack
    AssertEqualLong "TC-RUNNER-01 干跑AllocSuccessCount=0", 0, stats.AllocSuccessCount
    AssertEqualLong "TC-RUNNER-01 干跑AllocFailCount=0", 0, stats.AllocFailCount
End Sub


' TC-RUNNER-02：完整运行 — 1 个物流单号、1 个 SKU 组、全部成功，回溯 3 次。
Private Sub TestRunner_BuildRunStats_AllSuccess()
    Dim validationResult As ValidationResult
    validationResult.FailedShipmentCount = 0

    Dim results(1 To 1) As Object
    Set results(1) = RN_Test_MakeShipResult("SF_R02", 1, True, 3, "")

    Dim orders(1 To 2) As NormalizedReturnLine
    Dim inventory(1 To 4) As NormalizedInventoryLine

    Dim stats As RunStats
    stats = BuildRunStats(validationResult, results, orders, inventory)

    AssertEqualLong "TC-RUNNER-02 AllocSuccessCount=1", 1, stats.AllocSuccessCount
    AssertEqualLong "TC-RUNNER-02 AllocFailCount=0", 0, stats.AllocFailCount
    AssertEqualLong "TC-RUNNER-02 TotalBacktrackCount=3", 3, stats.TotalBacktrackCount
    AssertEqualLong "TC-RUNNER-02 MaxGroupBacktrack=3", 3, stats.MaxGroupBacktrack
End Sub


' TC-RUNNER-03：完整运行 — 2 个物流单号，1 成功（回溯 5 次）+ 1 失败（回溯 2 次）。
' 验证：AllocSuccessCount/AllocFailCount 各为 1，TotalBacktrackCount = 7，MaxGroupBacktrack = 5。
Private Sub TestRunner_BuildRunStats_MixedResult()
    Dim validationResult As ValidationResult
    validationResult.FailedShipmentCount = 1

    Dim results(1 To 2) As Object
    Set results(1) = RN_Test_MakeShipResult("SF_R03_A", 1, True, 5, "")
    Set results(2) = RN_Test_MakeShipResult("SF_R03_B", 1, False, 2, ERR_E10)

    Dim orders(1 To 4) As NormalizedReturnLine
    Dim inventory(1 To 3) As NormalizedInventoryLine

    Dim stats As RunStats
    stats = BuildRunStats(validationResult, results, orders, inventory)

    AssertEqualLong "TC-RUNNER-03 AllocSuccessCount=1", 1, stats.AllocSuccessCount
    AssertEqualLong "TC-RUNNER-03 AllocFailCount=1", 1, stats.AllocFailCount
    AssertEqualLong "TC-RUNNER-03 TotalBacktrackCount=7", 7, stats.TotalBacktrackCount
    AssertEqualLong "TC-RUNNER-03 MaxGroupBacktrack=5", 5, stats.MaxGroupBacktrack
    AssertEqualLong "TC-RUNNER-03 InputReturnRows=4", 4, stats.InputReturnRows
    AssertEqualLong "TC-RUNNER-03 InputInventoryRows=3", 3, stats.InputInventoryRows
End Sub


' TC-RUNNER-04：全空输入 — orders/inventory/shipmentResults 均未初始化。
' 验证：BuildRunStats 不崩溃，所有字段 = 0。
Private Sub TestRunner_BuildRunStats_EmptyInput()
    Dim validationResult As ValidationResult   ' HasFailures=False, FailedShipmentCount=0
    Dim emptyResults() As Object
    Dim emptyOrders() As NormalizedReturnLine
    Dim emptyInventory() As NormalizedInventoryLine

    Dim stats As RunStats
    stats = BuildRunStats(validationResult, emptyResults, emptyOrders, emptyInventory)

    AssertEqualLong "TC-RUNNER-04 InputReturnRows=0", 0, stats.InputReturnRows
    AssertEqualLong "TC-RUNNER-04 InputInventoryRows=0", 0, stats.InputInventoryRows
    AssertEqualLong "TC-RUNNER-04 TotalBacktrackCount=0", 0, stats.TotalBacktrackCount
    AssertEqualLong "TC-RUNNER-04 ValidationFailCount=0", 0, stats.ValidationFailCount
    AssertEqualLong "TC-RUNNER-04 AllocSuccessCount=0", 0, stats.AllocSuccessCount
    AssertEqualLong "TC-RUNNER-04 AllocFailCount=0", 0, stats.AllocFailCount
End Sub


' TC-RUNNER-05：1 个物流单号含 3 个 SKU 组，回溯次数分别为 2、7、4。
' 验证：TotalBacktrackCount = 13，MaxGroupBacktrack = 7（取最大值，不是总和）。
Private Sub TestRunner_BuildRunStats_MaxBacktrack()
    Dim validationResult As ValidationResult

    ' 手工构造 Dictionary，模拟含 3 组的 AllocateShipment 返回值
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ShipmentNo", "SF_R05"
    d.Add "GroupCount", CLng(3)
    d.Add "Group_1_Success", True
    d.Add "Group_1_BacktrackCount", CLng(2)
    d.Add "Group_2_Success", True
    d.Add "Group_2_BacktrackCount", CLng(7)
    d.Add "Group_3_Success", True
    d.Add "Group_3_BacktrackCount", CLng(4)

    Dim results(1 To 1) As Object
    Set results(1) = d

    Dim orders(1 To 1) As NormalizedReturnLine
    Dim inventory(1 To 1) As NormalizedInventoryLine

    Dim stats As RunStats
    stats = BuildRunStats(validationResult, results, orders, inventory)

    AssertEqualLong "TC-RUNNER-05 TotalBacktrackCount=13", 13, stats.TotalBacktrackCount
    AssertEqualLong "TC-RUNNER-05 MaxGroupBacktrack=7", 7, stats.MaxGroupBacktrack
    AssertEqualLong "TC-RUNNER-05 AllocSuccessCount=1(全部组成功)", 1, stats.AllocSuccessCount
    AssertEqualLong "TC-RUNNER-05 AllocFailCount=0", 0, stats.AllocFailCount
End Sub


' TC-43（R127）：连续两次 Full Run —— 输出表清空重写、运行历史累积追加、输入/配置不变。
' 数据：复用批量冒烟工作簿（SF_BT_SMOKE，1 行退单 + 1 行库存，全链路真实运行）。
' 断言对应 TC-43 验收判定：
'   ① 汇总/明细/异常表的脏数据被清空重写，第二次运行不累积重复行
'   ② 运行历史保留人工标记行，每次运行 +1
'   ③ 输入表与配置表内容不变
Private Sub TestRunner_RerunOverwrite_TC43()
    Dim wb As Workbook
    Set wb = BT_CreateSmokeSourceWorkbook()

    On Error GoTo CleanFail

    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet
    Dim wsAnomaly As Worksheet
    Dim wsHistory As Worksheet
    Set wsSummary = wb.Worksheets("分配状态汇总表")
    Set wsDetail = wb.Worksheets("成功分配明细表")
    Set wsAnomaly = wb.Worksheets("数据异常明细表")
    Set wsHistory = wb.Worksheets("运行历史记录表")

    ' 预置脏数据与人工标记，验证清空/保留边界
    wsSummary.Cells(2, 1).Value = "脏数据-汇总"
    wsDetail.Cells(2, 1).Value = "脏数据-明细"
    wsAnomaly.Cells(2, 1).Value = "脏数据-异常"
    wsHistory.Cells(1, 3).Value = "运行类型"
    wsHistory.Cells(2, 1).Value = "人工标记-运行前"

    RunFullAllocation wb, False

    ' 第一次运行后：输出重写为 SF_BT_SMOKE 的 1 行成功结果，历史追加 1 行
    AssertEqualLong "TC-43 第一次后汇总数据行=1", 1, TC43_DataRowCount(wsSummary)
    AssertEqualString "TC-43 汇总脏数据被清空重写", "SF_BT_SMOKE", CStr(wsSummary.Cells(2, 1).Value)
    AssertEqualString "TC-43 汇总状态=批量导入", STATUS_BATCH_IMPORT, CStr(wsSummary.Cells(2, 3).Value)
    AssertEqualLong "TC-43 第一次后明细数据行=1", 1, TC43_DataRowCount(wsDetail)
    AssertEqualLong "TC-43 第一次后异常数据行=0", 0, TC43_DataRowCount(wsAnomaly)
    AssertEqualString "TC-43 历史保留人工标记", "人工标记-运行前", CStr(wsHistory.Cells(2, 1).Value)
    AssertEqualString "TC-43 第一次历史追加FullRun", "Full Run", CStr(wsHistory.Cells(3, 3).Value)
    AssertEqualString "TC-43 第一次后历史无第4行数据", vbNullString, CStr(wsHistory.Cells(4, 1).Value)

    RunFullAllocation wb, False

    ' 第二次运行后：输出行数不累积，历史再 +1
    AssertEqualLong "TC-43 第二次后汇总数据行=1", 1, TC43_DataRowCount(wsSummary)
    AssertEqualLong "TC-43 第二次后明细数据行=1", 1, TC43_DataRowCount(wsDetail)
    AssertEqualString "TC-43 第二次汇总仍是SF_BT_SMOKE", "SF_BT_SMOKE", CStr(wsSummary.Cells(2, 1).Value)
    AssertEqualString "TC-43 历史标记行仍在", "人工标记-运行前", CStr(wsHistory.Cells(2, 1).Value)
    AssertEqualString "TC-43 第二次历史第3行=FullRun", "Full Run", CStr(wsHistory.Cells(3, 3).Value)
    AssertEqualString "TC-43 第二次历史第4行=FullRun", "Full Run", CStr(wsHistory.Cells(4, 3).Value)
    AssertEqualString "TC-43 第二次后历史无第5行数据", vbNullString, CStr(wsHistory.Cells(5, 1).Value)

    ' 输入与配置不被运行改动
    AssertEqualString "TC-43 输入退单表不变", "SF_BT_SMOKE", CStr(wb.Worksheets("输入_退单表").Cells(2, 1).Value)
    AssertEqualString "TC-43 输入库存表不变", "H_BT_SMOKE", CStr(wb.Worksheets("输入_质检库存表").Cells(2, 2).Value)
    AssertEqualString "TC-43 配置表不变", "SF_BT_SMOKE", CStr(wb.Worksheets("输入_配置").Cells(2, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] TC-43 异常：" & Err.Description
    AssertTrue "TC-43 执行不抛错（" & Err.Description & "）", False
    Resume CleanExit
End Sub

' 统计输出表数据行数（第 1 行为表头；空表记 0）。
Private Function TC43_DataRowCount(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then
        TC43_DataRowCount = 0
    Else
        TC43_DataRowCount = lastRow - 1
    End If
End Function

' E02 回归（2026-07-19 用户报告缺陷）：行号重复时，数据异常明细表“原始值”
' 必须保留文本 "00001" 的前导零，不得被 Excel 强转成数值 1 误导用户。
' 根因：M14 WriteSheet 写矩阵时 General 格式列会把数字样式文本转数值；
' 修复：EO_ApplyTextFormats 对“原始值”列按文本写入。
Private Sub TestRunner_E02RawValueKeepsLeadingZeros()
    Dim wb As Workbook
    Set wb = BT_CreateSmokeSourceWorkbook()

    On Error GoTo CleanFail

    ' 退单表：同一 WMS 下两行均为行号 "00001"（文本），数量各 1，合计 2
    Dim wsR As Worksheet
    Set wsR = wb.Worksheets("输入_退单表")
    wsR.Cells(3, 1).Value = "SF_BT_SMOKE"
    wsR.Cells(3, 2).Value = "TK_BT_SMOKE"
    wsR.Cells(3, 3).Value = "H_BT_SMOKE"
    wsR.Cells(3, 4).NumberFormat = "@"
    wsR.Cells(3, 4).Value = "00001"
    wsR.Cells(3, 5).Value = 1

    ' 库存补足到与退单总量一致（E08 通过，确保仅 E02 触发）
    Dim wsI As Worksheet
    Set wsI = wb.Worksheets("输入_质检库存表")
    wsI.Cells(2, 6).Value = 2

    RunFullAllocation wb, False

    ' 原始值在第 7 列（来源表|Excel行号|物流单号|WMS退单号|SKU|字段名|原始值|错误码|原因说明）
    Dim wsA As Worksheet
    Set wsA = wb.Worksheets("数据异常明细表")

    AssertEqualString "E02 第1行错误码", "E02", CStr(wsA.Cells(2, 8).Value)
    AssertEqualString "E02 第1行原始值保留前导零", "00001", wsA.Cells(2, 7).Text
    AssertEqualString "E02 第1行原始值按文本存储", "String", TypeName(wsA.Cells(2, 7).Value)
    AssertEqualString "E02 第2行原始值保留前导零", "00001", wsA.Cells(3, 7).Text
    AssertEqualString "E02 第2行原始值按文本存储", "String", TypeName(wsA.Cells(3, 7).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] E02原始值回归测试异常：" & Err.Description
    AssertTrue "E02原始值 执行不抛错（" & Err.Description & "）", False
    Resume CleanExit
End Sub


' 构造单组 ShipmentResult Dictionary，供 TC-RUNNER-02/03 使用。
' 结构与 modBacktracking.AllocateShipment 的返回格式一致。
Private Function RN_Test_MakeShipResult( _
    ByVal shipNo As String, _
    ByVal groupCount As Long, _
    ByVal success As Boolean, _
    ByVal backtrackCount As Long, _
    ByVal errorCode As String) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ShipmentNo", shipNo
    d.Add "GroupCount", groupCount
    d.Add "Group_1_SKU", "TEST_SKU"
    d.Add "Group_1_Success", success
    d.Add "Group_1_BacktrackCount", backtrackCount
    d.Add "Group_1_ErrorCode", errorCode
    d.Add "Group_1_DetailCount", IIf(success, CLng(1), CLng(0))

    Set RN_Test_MakeShipResult = d
End Function


' =============================================================================
' E2E 验收测试（T01~T14）
' =============================================================================
' 设计意图：
'   以上各 Run*Tests() 套件已对每个模块进行了独立单元测试。
'   本节端到端（E2E）验收测试的目的：
'     1. 验证所有模块在完整分配链路（M05→M06→M07→M09→M11）中协同正确
'     2. 从最终可见的 FinalResult 验证系统行为，不深入中间状态
'     3. 覆盖规格 §6.5.3 要求的 T1~T11+ 验收场景
'   全部测试数据内联构造，不依赖外部 Excel 文件，确保稳定可重复。
' =============================================================================

Public Sub RunAcceptanceTests()
    ' E2E 验收测试入口：T01~T14
    BeginSuite "E2E Acceptance Tests (T01~T14)"

    On Error GoTo CleanFail

    TestAcceptance_T01_StrategyOne
    TestAcceptance_T02_StrategyTwo
    TestAcceptance_T03_StrategyThree
    TestAcceptance_T04_MultiQCSort
    TestAcceptance_T05_BacktrackSuccess
    TestAcceptance_T06_E10_CascadeRollback
    TestAcceptance_T07_E11_FragmentInventory
    TestAcceptance_T08_E08_QtyMismatch
    TestAcceptance_T09_FullRollback
    TestAcceptance_T10_E07_OrphanShipment
    TestAcceptance_T11_LineNoRules
    TestAcceptance_T12_LotCaseSensitive
    TestAcceptance_T13_ExpiryRules
    TestAcceptance_T14_RealBacktrackBatchStatus

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] E2E Acceptance Tests 执行异常：" & Err.Description
    FinishSuite
End Sub


' T01：策略一精确匹配
' 场景：1行 D=5，ZP库存=5 → 精确一次成功，不需要回溯
' 期望：BacktrackCount=0，AllocQty=5，行状态=批量导入，汇总状态=批量导入
Private Sub TestAcceptance_T01_StrategyOne()
    Dim shipNo As String: shipNo = "SF_AT01"
    Dim sku As String: sku = "H_AT01"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT01", sku, "00001", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    ' 调用分配引擎，验证 BacktrackCount=0（无回溯即精确匹配证明）
    Dim allocResult As Object
    Set allocResult = E2E_AllocateSingleSKU(shipNo, sku, orders, ledger, cfg)

    AssertTrue        "T01 分配成功",         CBool(allocResult("Group_1_Success"))
    AssertEqualLong   "T01 BacktrackCount=0", 0, CLng(allocResult("Group_1_BacktrackCount"))
    AssertEqualLong   "T01 分配量=5",         5, CLng(allocResult("Group_1_AllocQty_1"))

    ' 通过 ApplyRollback 验证最终行状态（完整链路校验）
    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualString "T01 行状态=批量导入",   STATUS_BATCH_IMPORT, CStr(finalResult("Detail_1_LineStatus"))
    AssertEqualString "T01 汇总状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Summary_1_Status"))
End Sub


' T02：策略二最接近匹配（剩余库存保留）
'
' 数据设计（关键分析）：
'   2行 D=3 each，ZP-LA01=6（刚好 3+3），QC-LB01=5（不会被使用）
'   BuildStaticPlan（n=2，isOnlyRow=False）：
'     ZP: T=6 ≥ D=3+nmq=3=6 → initQCCount=1（两行均是）
'     QC: T=5 ≥ D=3+nmq=3=6？ 5<6 → 不入候选。T=D=3？5≠3 → 不入候选。
'   排序：两行 initQCCount 相同，lineNo 升序 → pos1="00001"（非末行），pos2="00002"（末行）
'   FilterCandidatePool pos1（非末，D=3,nmq=3）：ZP T=6≥6 → 候选。Strategy 2（T≠D）：分配3，ZP→3
'   FilterCandidatePool pos2（末行，D=3,isLastRow=True）：ZP T=3=D=3 → 候选。Strategy 1：分配3，ZP→0
'   QC=5 全程未使用 → 剩余库存保留 ✓
'
' 期望：两行均成功，ZP剩余=0，QC剩余=5（验证剩余库存保留），行状态=批量导入
Private Sub TestAcceptance_T02_StrategyTwo()
    Dim shipNo As String: shipNo = "SF_AT02"
    Dim sku As String: sku = "H_AT02"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    ' ZP=6（2行各取3，恰好耗尽），QC=5（不被使用，验证"剩余库存保留"）
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 6)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_QC, "LB01", "2029/01/01", 5)

    ' 2行各 D=3，同一 WMS退单号
    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT02", sku, "00001", 3)
    orders(2) = E2E_MakeReturnLine(shipNo, "TK_AT02", sku, "00002", 3)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim allocResult As Object
    Set allocResult = E2E_AllocateSingleSKU(shipNo, sku, orders, ledger, cfg)

    AssertTrue      "T02 分配成功",        CBool(allocResult("Group_1_Success"))
    AssertEqualLong "T02 明细数=2",        2, CLng(allocResult("Group_1_DetailCount"))
    AssertEqualLong "T02 各行分配量=3",    3, CLng(allocResult("Group_1_AllocQty_1"))

    ' 账本验证：ZP 耗尽=0（非末行用 Strategy 2 分配3后剩3，末行 Strategy 1 分配3）
    AssertEqualLong "T02 ZP耗尽=0",        0, QueryQCTotal(ledger, shipNo, sku, QC_ZP)
    ' 关键：QC 从未使用，剩余库存完整保留
    AssertEqualLong "T02 QC剩余库存=5",    5, QueryQCTotal(ledger, shipNo, sku, QC_QC)

    ' FinalResult 行状态验证：单批次（LA01）→ 批量导入
    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualLong   "T02 明细总数=2",       2, CLng(finalResult("DetailCount"))
    AssertEqualString "T02 汇总状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Summary_1_Status"))
End Sub


' T03：策略三跨批号/效期拼凑（多明细）
' 场景：1行 D=5，ZP-LA01=2，ZP-LA02=3 → 拼凑两条明细合计5
' 期望：明细数=2，行状态=手工操作（多批次/效期组合）
Private Sub TestAcceptance_T03_StrategyThree()
    Dim shipNo As String: shipNo = "SF_AT03"
    Dim sku As String: sku = "H_AT03"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 2)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA02", "2030/01/01", 3)

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT03", sku, "00001", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim allocResult As Object
    Set allocResult = E2E_AllocateSingleSKU(shipNo, sku, orders, ledger, cfg)

    AssertTrue      "T03 分配成功",   CBool(allocResult("Group_1_Success"))
    AssertEqualLong "T03 明细数=2",   2, CLng(allocResult("Group_1_DetailCount"))

    ' FinalResult 行状态：跨批号/效期使用了多条明细 → 手工操作
    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualString "T03 行状态=手工操作", STATUS_MANUAL, CStr(finalResult("Detail_1_LineStatus"))
End Sub


' T04：多QC竞争+静态排序（ZP优先于QC优先于NG）
' 场景：1行 D=5，ZP=5，QC=5 → ZP优先被消耗，QC保持不变
' 期望：使用QC=ZP，ZP剩余=0，QC剩余=5
Private Sub TestAcceptance_T04_MultiQCSort()
    Dim shipNo As String: shipNo = "SF_AT04"
    Dim sku As String: sku = "H_AT04"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_QC, "LB01", "2029/01/01", 5)

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT04", sku, "00001", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim allocResult As Object
    Set allocResult = E2E_AllocateSingleSKU(shipNo, sku, orders, ledger, cfg)

    AssertTrue        "T04 分配成功",   CBool(allocResult("Group_1_Success"))
    AssertEqualString "T04 使用QC=ZP", QC_ZP, CStr(allocResult("Group_1_QC_1"))
    ' 账本验证：ZP 被完全消耗，QC 未动
    AssertEqualLong   "T04 ZP剩余=0",  0, QueryQCTotal(ledger, shipNo, sku, QC_ZP)
    AssertEqualLong   "T04 QC剩余=5",  5, QueryQCTotal(ledger, shipNo, sku, QC_QC)
End Sub


' T05：回溯触发并成功
' 场景（与 TC-BT03 相同数据结构）：
'   ZP=5，QC=13；两行各 D=5，同一 WMS退单号
'   贪心第1次：pos1选ZP→ZP=0；pos2末行：ZP=0，QC(T≠D)→失败
'   回溯第1次：撤销pos1(ZP)，改选QC；pos2末行：ZP(T=D=5)→成功
' 期望：BacktrackCount≥1，整单成功，汇总状态=批量导入
Private Sub TestAcceptance_T05_BacktrackSuccess()
    Dim shipNo As String: shipNo = "SF_AT05"
    Dim sku As String: sku = "H_AT05"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, sku, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, sku, QC_QC, "LB01", "2029/01/01", 13)

    Dim orders(1 To 2) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT05", sku, "00001", 5)
    orders(2) = E2E_MakeReturnLine(shipNo, "TK_AT05", sku, "00002", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim allocResult As Object
    Set allocResult = E2E_AllocateSingleSKU(shipNo, sku, orders, ledger, cfg)

    AssertTrue    "T05 回溯后成功",         CBool(allocResult("Group_1_Success"))
    AssertTrue    "T05 BacktrackCount>=1",  CLng(allocResult("Group_1_BacktrackCount")) >= 1
    AssertEqualLong "T05 明细数=2",         2, CLng(allocResult("Group_1_DetailCount"))

    ' FinalResult 验证：回溯后最终成功，汇总不应为"无法分配"
    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualString "T05 汇总状态非无法分配", STATUS_BATCH_IMPORT, CStr(finalResult("Summary_1_Status"))
End Sub


' T06：回溯耗尽 E10 + 连带回滚
' 场景：2个SKU组，SKU-A需求15而库存只有10，回溯耗尽触发E10；SKU-B被连带回滚
' 期望：Group_1 ErrorCode=E10，Group_2 ErrorCode=连带回滚，汇总状态=无法分配
Private Sub TestAcceptance_T06_E10_CascadeRollback()
    Dim shipNo As String: shipNo = "SF_AT06"
    Dim skuA As String: skuA = "H_AT06A"
    Dim skuB As String: skuB = "H_AT06B"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    cfg.MaxBacktrackCount = 2   ' 第3次回溯触发 E10

    ' SKU-A：3行各 D=5，总需15，ZP=5+QC=5=10，数学上不可能满足 → E10
    ' SKU-B：1行 D=5，ZP=5，单独可成功，但因 SKU-A 失败被连带回滚
    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, skuA, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, skuA, QC_QC, "LB01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine(shipNo, skuB, QC_ZP, "LC01", "2029/01/01", 5)

    Dim orders(1 To 4) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT06", skuA, "00001", 5)
    orders(2) = E2E_MakeReturnLine(shipNo, "TK_AT06", skuA, "00002", 5)
    orders(3) = E2E_MakeReturnLine(shipNo, "TK_AT06", skuA, "00003", 5)
    orders(4) = E2E_MakeReturnLine(shipNo, "TK_AT06", skuB, "00001", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim skuList(1 To 2) As String
    skuList(1) = skuA
    skuList(2) = skuB
    Dim planMap As Object, precheckMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Set precheckMap = CreateObject("Scripting.Dictionary")
    E2E_AddSkuToPlanMap orders, shipNo, skuA, ledger, planMap, precheckMap
    E2E_AddSkuToPlanMap orders, shipNo, skuB, ledger, planMap, precheckMap

    Dim allocResult As Object
    Set allocResult = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    AssertFalse       "T06 SKU-A失败",               CBool(allocResult("Group_1_Success"))
    AssertEqualString "T06 SKU-A ErrorCode=E10",     ERR_E10, CStr(allocResult("Group_1_ErrorCode"))
    AssertFalse       "T06 SKU-B被连带回滚",          CBool(allocResult("Group_2_Success"))
    AssertEqualString "T06 SKU-B ErrorCode=连带回滚", ERROR_CASCADE_ROLLBACK, CStr(allocResult("Group_2_ErrorCode"))

    ' FinalResult 验证：整单无法分配，原因含 E10
    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualString "T06 汇总状态=无法分配", STATUS_UNALLOCATED, CStr(finalResult("Summary_1_Status"))
    AssertTrue        "T06 汇总原因含E10",    InStr(1, CStr(finalResult("Summary_1_Reason")), ERR_E10) > 0
End Sub


' T07：E11 碎片库存（每个 QC 单独都不够，拼凑满足条件不成立）
' 场景：1行 D=2，ZP=1+QC=1；每个 QC 只有 1，均 < D_min=2 → E11
' 期望：ValidatePre 返回 E11，校验阶段直接失败
Private Sub TestAcceptance_T07_E11_FragmentInventory()
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    Dim shipNo As String: shipNo = "SF_AT07"

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT07", "H_AT07", "00001", 2)

    ' ZP=1，QC=1，每种 QC 数量均 < D=2，满足 E11 触发条件（0 < T < groupMinQty）
    Dim inv(1 To 2) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, "H_AT07", QC_ZP, "LA01", "2029/01/01", 1)
    inv(2) = MakeInventoryLine(shipNo, "H_AT07", QC_QC, "LB01", "2029/01/01", 1)

    Dim valIssues() As ValidationIssue
    Dim valResult As ValidationResult
    valResult = ValidatePre(orders, inv, EmptyFieldIssueArray(), cfg, valIssues)

    AssertTrue "T07 E11 校验失败",  valResult.HasFailures
    AssertTrue "T07 E11 命中",      ShipmentHasError(valIssues, shipNo, ERR_E11)
End Sub


' T08：E08 退单量与库存量不一致
' 场景：退单 D=5，库存 ZP=3；5≠3 → E08
' 期望：ValidatePre 返回 E08，校验阶段直接失败
Private Sub TestAcceptance_T08_E08_QtyMismatch()
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    Dim shipNo As String: shipNo = "SF_AT08"

    Dim orders(1 To 1) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT08", "H_AT08", "00001", 5)

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, "H_AT08", QC_ZP, "LA01", "2029/01/01", 3)

    Dim valIssues() As ValidationIssue
    Dim valResult As ValidationResult
    valResult = ValidatePre(orders, inv, EmptyFieldIssueArray(), cfg, valIssues)

    AssertTrue "T08 E08 校验失败", valResult.HasFailures
    AssertTrue "T08 E08 命中",     ShipmentHasError(valIssues, shipNo, ERR_E08)
End Sub


' T09：整单回滚（部分 SKU 失败，连带回滚格式正确）
' 场景：SKU-A 需求15超过库存10，MaxBacktrack=1 快速触发 E10；
'       SKU-B 被连带回滚；FinalResult 汇总原因含 E10 触发标记
' 期望：整单状态=无法分配，连带回滚原因格式含"E10"
Private Sub TestAcceptance_T09_FullRollback()
    Dim shipNo As String: shipNo = "SF_AT09"
    Dim skuA As String: skuA = "H_AT09A"
    Dim skuB As String: skuB = "H_AT09B"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    cfg.MaxBacktrackCount = 1   ' 第2次回溯触发 E10，快速触发失败

    Dim inv(1 To 3) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, skuA, QC_ZP, "LA01", "2029/01/01", 5)
    inv(2) = MakeInventoryLine(shipNo, skuA, QC_QC, "LB01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine(shipNo, skuB, QC_ZP, "LC01", "2029/01/01", 5)

    Dim orders(1 To 4) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK_AT09", skuA, "00001", 5)
    orders(2) = E2E_MakeReturnLine(shipNo, "TK_AT09", skuA, "00002", 5)
    orders(3) = E2E_MakeReturnLine(shipNo, "TK_AT09", skuA, "00003", 5)
    orders(4) = E2E_MakeReturnLine(shipNo, "TK_AT09", skuB, "00001", 5)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim skuList(1 To 2) As String
    skuList(1) = skuA
    skuList(2) = skuB
    Dim planMap As Object, precheckMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Set precheckMap = CreateObject("Scripting.Dictionary")
    E2E_AddSkuToPlanMap orders, shipNo, skuA, ledger, planMap, precheckMap
    E2E_AddSkuToPlanMap orders, shipNo, skuB, ledger, planMap, precheckMap

    Dim allocResult As Object
    Set allocResult = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    ' 整单失败
    AssertEqualString "T09 整单状态=无法分配", STATUS_UNALLOCATED, CStr(finalResult("Summary_1_Status"))
    ' 连带回滚原因格式：应含"E10"（触发原因）
    AssertTrue "T09 原因含E10触发标记", InStr(1, CStr(finalResult("Summary_1_Reason")), ERR_E10) > 0
End Sub


' T10：E07 孤立物流单号（库存有该物流单号，退单表没有）
' 场景：inventory 中有 SF_AT10，orders 为空
' 期望：E07 触发，ValidationIssue.WMSOrderNo=[N/A]
Private Sub TestAcceptance_T10_E07_OrphanShipment()
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    Dim shipNo As String: shipNo = "SF_AT10"

    Dim inv(1 To 1) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, "H_AT10", QC_ZP, "LA01", "2029/01/01", 5)

    Dim valIssues() As ValidationIssue
    Dim valResult As ValidationResult
    valResult = ValidatePre(EmptyReturnArray(), inv, EmptyFieldIssueArray(), cfg, valIssues)

    AssertTrue "T10 E07 校验失败", valResult.HasFailures
    AssertTrue "T10 E07 命中",     ShipmentHasError(valIssues, shipNo, ERR_E07)

    ' 验证 WMSOrderNo=[N/A]（孤立物流单号没有对应的 WMS退单号，规格要求填占位符）
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = LBound(valIssues) To UBound(valIssues)
        If valIssues(i).ErrorCode = ERR_E07 And valIssues(i).ShipmentNo = shipNo Then
            AssertEqualString "T10 WMSOrderNo=[N/A]", NA_PLACEHOLDER, valIssues(i).WMSOrderNo
            found = True
            Exit For
        End If
    Next i
    AssertTrue "T10 找到E07问题记录", found
End Sub


' T11：行号五位文本规则（M04 标准化层）
' 场景：LineNo="00001"（文本五位）合法；LineNo=1（数值型）非法
' 期望：合法行 LineNoValid=True；非法行 LineNoValid=False；FieldNormalizeIssue.RawValue="1"
Private Sub TestAcceptance_T11_LineNoRules()
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    ' 内联构造原始退单数据（模拟 M03 输出的 RawReturnRow）
    Dim raw(1 To 2) As RawReturnRow
    raw(1).ExcelRowNum = 2
    raw(1).ShipmentNo  = "SF_AT11"
    raw(1).WMSOrderNo  = "TK_AT11"
    raw(1).SKU         = "H_AT11"
    raw(1).LineNo      = "00001"    ' 合法：文本型五位前导零
    raw(1).Qty         = 3

    raw(2).ExcelRowNum = 3
    raw(2).ShipmentNo  = "SF_AT11"
    raw(2).WMSOrderNo  = "TK_AT11"
    raw(2).SKU         = "H_AT11"
    raw(2).LineNo      = 1          ' 非法：数值型，经 CStr 得 "1"，不满足五位前导零规则
    raw(2).Qty         = 3

    Dim issues() As FieldNormalizeIssue
    Dim normalized() As NormalizedReturnLine
    normalized = NormalizeReturnRows(raw, cfg, issues)

    AssertEqualBool   "T11 文本五位行号合法",   True,  normalized(1).LineNoValid
    AssertEqualBool   "T11 数值型行号非法",     False, normalized(2).LineNoValid
    AssertEqualLong   "T11 产生1条行号问题",    1,     CountFieldIssues(issues)
    AssertEqualString "T11 RawValue=1",        "1",   issues(LBound(issues)).RawValue
End Sub


' T12：批号大小写敏感/不敏感（M04 标准化层）
' 场景：LotNo="a01"；不敏感模式 → 转为大写 "A01"；敏感模式 → 保留原值 "a01"
Private Sub TestAcceptance_T12_LotCaseSensitive()
    ' 内联构造原始库存数据（模拟 M03 输出的 RawInventoryRow）
    Dim inv(1 To 1) As RawInventoryRow
    inv(1).ExcelRowNum    = 2
    inv(1).ShipmentNo     = "SF_AT12"
    inv(1).SKU            = "H_AT12"
    inv(1).QC             = "ZP"
    inv(1).LotNo          = "a01"   ' 小写批号，测试大小写规则
    inv(1).Expiry         = "2029/01/01"
    inv(1).ExpiryCellKind = CELL_KIND_TEXT
    inv(1).Qty            = 5

    ' 不敏感模式（默认）："a01" 应标准化为 "A01"
    Dim cfgI As ConfigStruct: cfgI = E2E_DefaultCfg()
    cfgI.LotCaseSensitive = False
    Dim issuesI() As FieldNormalizeIssue
    Dim normI() As NormalizedInventoryLine
    normI = NormalizeInventoryRows(inv, cfgI, issuesI)
    AssertEqualString "T12 不敏感模式 a01→A01", "A01", normI(1).LotNo

    ' 敏感模式："a01" 应保持原值
    Dim cfgS As ConfigStruct: cfgS = E2E_DefaultCfg()
    cfgS.LotCaseSensitive = True
    Dim issuesS() As FieldNormalizeIssue
    Dim normS() As NormalizedInventoryLine
    normS = NormalizeInventoryRows(inv, cfgS, issuesS)
    AssertEqualString "T12 敏感模式 a01 保持原值", "a01", normS(1).LotNo
End Sub


' T13：效期规则（ExcelDate序列号、文本日期格式、非法日期）
' 覆盖规格 §M04 中效期标准化的四种输入类型
Private Sub TestAcceptance_T13_ExpiryRules()
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()

    Dim raw(1 To 4) As RawInventoryRow

    ' 场景1：ExcelDate 序列号（Excel内部存储的日期） → 正确转 YYYY/MM/DD
    raw(1).ExcelRowNum    = 2
    raw(1).ShipmentNo     = "SF_AT13": raw(1).SKU = "H_AT13": raw(1).QC = "ZP"
    raw(1).LotNo          = "L1"
    raw(1).Expiry         = DateSerial(2029, 6, 15)
    raw(1).ExpiryCellKind = CELL_KIND_EXCEL_DATE
    raw(1).Qty            = 1

    ' 场景2：文本效期 YYYY/MM/DD → 合法，原样输出
    raw(2).ExcelRowNum    = 3
    raw(2).ShipmentNo     = "SF_AT13": raw(2).SKU = "H_AT13": raw(2).QC = "ZP"
    raw(2).LotNo          = "L2"
    raw(2).Expiry         = "2029/01/01"
    raw(2).ExpiryCellKind = CELL_KIND_TEXT
    raw(2).Qty            = 1

    ' 场景3：文本效期 YYYY-MM-DD（连字符分隔） → 合法，统一转 YYYY/MM/DD 格式
    raw(3).ExcelRowNum    = 4
    raw(3).ShipmentNo     = "SF_AT13": raw(3).SKU = "H_AT13": raw(3).QC = "ZP"
    raw(3).LotNo          = "L3"
    raw(3).Expiry         = "2029-12-31"
    raw(3).ExpiryCellKind = CELL_KIND_TEXT
    raw(3).Qty            = 1

    ' 场景4：非闰年的 2月29日 → 非法（字符串级校验，2029不是闰年）
    raw(4).ExcelRowNum    = 5
    raw(4).ShipmentNo     = "SF_AT13": raw(4).SKU = "H_AT13": raw(4).QC = "ZP"
    raw(4).LotNo          = "L4"
    raw(4).Expiry         = "2029/02/29"
    raw(4).ExpiryCellKind = CELL_KIND_TEXT
    raw(4).Qty            = 1

    Dim issues() As FieldNormalizeIssue
    Dim normalized() As NormalizedInventoryLine
    normalized = NormalizeInventoryRows(raw, cfg, issues)

    AssertEqualString "T13 ExcelDate转换正确",       "2029/06/15", normalized(1).Expiry
    AssertEqualBool   "T13 ExcelDate合法",           True,         normalized(1).ExpiryValid
    AssertEqualString "T13 斜杠文本效期原样输出",     "2029/01/01", normalized(2).Expiry
    AssertEqualBool   "T13 斜杠文本效期合法",         True,         normalized(2).ExpiryValid
    AssertEqualString "T13 连字符转斜杠格式",         "2029/12/31", normalized(3).Expiry
    AssertEqualBool   "T13 连字符文本效期合法",       True,         normalized(3).ExpiryValid
    AssertEqualBool   "T13 非闰年2月29日非法",       False,        normalized(4).ExpiryValid
    AssertEqualString "T13 非法效期RawValue保留",    "2029/02/29", issues(LBound(issues)).RawValue
End Sub


' T14：真实回归场景——回溯成功后仍应批量导入
' 覆盖曾经暴露的问题：
'   1) 同一物流单号下多个 WMS 退单号均有 00001，行状态不能互相污染
'   2) 策略二/回溯成功不等于手工操作
'   3) 行号与效期在最终结果中保持标准字符串
Private Sub TestAcceptance_T14_RealBacktrackBatchStatus()
    Dim shipNo As String: shipNo = "SF3190000000016"
    Dim cfg As ConfigStruct: cfg = E2E_DefaultCfg()
    cfg.DebugLogLevel = DEBUG_LEVEL_DETAIL

    Dim inv(1 To 7) As NormalizedInventoryLine
    inv(1) = MakeInventoryLine(shipNo, "H000000001", QC_QC, "LA01", "2029/01/01", 12)
    inv(2) = MakeInventoryLine(shipNo, "H000000001", QC_QC, "LA01", "2029/01/01", 5)
    inv(3) = MakeInventoryLine(shipNo, "H000000001", QC_QC, "LA01", "2029/01/01", 5)
    inv(4) = MakeInventoryLine(shipNo, "H000000001", QC_ZP, "LA01", "2029/01/01", 8)
    inv(5) = MakeInventoryLine(shipNo, "H000000001", QC_ZP, "LA01", "2029/01/01", 12)
    inv(6) = MakeInventoryLine(shipNo, "H000000002", QC_NG, "LB01", "2029/01/01", 5)
    inv(7) = MakeInventoryLine(shipNo, "H000000003", QC_NG, "LB01", "2029/01/01", 1)

    Dim orders(1 To 10) As NormalizedReturnLine
    orders(1) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00001", 12)
    orders(2) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00002", 5)
    orders(3) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00003", 5)
    orders(4) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00004", 5)
    orders(5) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00005", 5)
    orders(6) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00006", 5)
    orders(7) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000001", "00007", 5)
    orders(8) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000002", "00008", 3)
    orders(9) = E2E_MakeReturnLine(shipNo, "TK10000162", "H000000002", "00001", 2)
    orders(10) = E2E_MakeReturnLine(shipNo, "TK10000161", "H000000003", "00009", 1)

    Dim ledger As Object
    Set ledger = BuildLedger(inv)

    Dim skuList(1 To 3) As String
    skuList(1) = "H000000001"
    skuList(2) = "H000000002"
    skuList(3) = "H000000003"

    Dim planMap As Object, precheckMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Set precheckMap = CreateObject("Scripting.Dictionary")
    E2E_AddSkuToPlanMap orders, shipNo, skuList(1), ledger, planMap, precheckMap
    E2E_AddSkuToPlanMap orders, shipNo, skuList(2), ledger, planMap, precheckMap
    E2E_AddSkuToPlanMap orders, shipNo, skuList(3), ledger, planMap, precheckMap

    Dim allocResult As Object
    Set allocResult = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)

    AssertTrue "T14 H000000001 回溯后成功", CBool(allocResult("Group_1_Success"))
    AssertTrue "T14 H000000001 发生回溯", CLng(allocResult("Group_1_BacktrackCount")) > 0
    AssertTrue "T14 H000000002 成功", CBool(allocResult("Group_2_Success"))
    AssertTrue "T14 H000000003 成功", CBool(allocResult("Group_3_Success"))

    Dim shipResults(1 To 1) As Object
    Set shipResults(1) = allocResult
    Dim emptyIssues() As ValidationIssue
    Dim valResult As ValidationResult
    Dim finalResult As Object
    Set finalResult = ApplyRollback(shipResults, valResult, emptyIssues, orders)

    AssertEqualLong "T14 成功明细数=10", 10, CLng(finalResult("DetailCount"))
    AssertEqualString "T14 TK10000161 汇总=批量导入", STATUS_BATCH_IMPORT, E2E_FindSummaryStatus(finalResult, "TK10000161")
    AssertEqualString "T14 TK10000162 汇总=批量导入", STATUS_BATCH_IMPORT, E2E_FindSummaryStatus(finalResult, "TK10000162")
    AssertEqualString "T14 首行行号保留00001", "00001", CStr(finalResult("Detail_1_LineNo"))
    AssertEqualString "T14 首行效期保留yyyy/mm/dd", "2029/01/01", CStr(finalResult("Detail_1_Expiry"))
    AssertEqualString "T14 首行状态=批量导入", STATUS_BATCH_IMPORT, CStr(finalResult("Detail_1_LineStatus"))
End Sub


' =============================================================================
' E2E 验收测试私有辅助函数
' =============================================================================

' 构造测试用 NormalizedReturnLine（全字段有效）。
' 设计意图：与 BT_/ST_/PV_ 前缀的同类函数职责一致，以 E2E_ 前缀与其他模块的
' 测试辅助函数区分，明确它只服务于端到端验收测试。
Private Function E2E_MakeReturnLine( _
    ByVal shipNo As String, _
    ByVal wmsNo As String, _
    ByVal sku As String, _
    ByVal lineNo As String, _
    ByVal qty As Long) As NormalizedReturnLine

    Dim r As NormalizedReturnLine
    r.ShipmentNo  = shipNo
    r.WMSOrderNo  = wmsNo
    r.SKU         = sku
    r.LineNo      = lineNo
    r.Qty         = qty
    r.LineNoValid = True
    r.QtyValid    = True
    E2E_MakeReturnLine = r
End Function


' 构造默认测试配置（MaxBacktrack=200，日志关闭，批号不敏感，默认哨兵值）。
' 设计意图：避免每个测试函数重复初始化 ConfigStruct，需要覆盖特定字段时在调用方修改。
Private Function E2E_DefaultCfg() As ConfigStruct
    Dim cfg As ConfigStruct
    cfg.MaxBacktrackCount = DEFAULT_MAX_BACKTRACK_COUNT
    cfg.DebugLogLevel     = DEBUG_LEVEL_OFF
    cfg.DetailedLogLimit  = DEFAULT_DETAILED_LOG_LIMIT
    cfg.LotCaseSensitive  = False
    cfg.NoExpirySentinel  = DEFAULT_NO_EXPIRY_SENTINEL
    E2E_DefaultCfg = cfg
End Function


Private Function E2E_FindSummaryStatus(ByVal finalResult As Object, ByVal wmsOrderNo As String) As String
    If finalResult Is Nothing Then Exit Function
    If Not finalResult.Exists("SummaryCount") Then Exit Function

    Dim i As Long
    For i = 1 To CLng(finalResult("SummaryCount"))
        If CStr(finalResult("Summary_" & i & "_WMSOrderNo")) = wmsOrderNo Then
            E2E_FindSummaryStatus = CStr(finalResult("Summary_" & i & "_Status"))
            Exit Function
        End If
    Next i
End Function


' 针对单一物流单号 + 单一 SKU 的简化分配助手。
' 封装了 BuildStaticPlan → RunPrecheck → AllocateShipment 三步，
' 减少 T01~T05 等简单场景的重复代码。
' 返回：AllocateShipment 的 Object（Dictionary），可直接检查 BacktrackCount 等字段，
' 也可传入 ApplyRollback 进一步得到 FinalResult。
Private Function E2E_AllocateSingleSKU( _
    ByVal shipNo As String, _
    ByVal sku As String, _
    ByRef orders() As NormalizedReturnLine, _
    ByVal ledger As Object, _
    ByRef cfg As ConfigStruct) As Object

    Dim planMap As Object
    Dim precheckMap As Object
    Set planMap = CreateObject("Scripting.Dictionary")
    Set precheckMap = CreateObject("Scripting.Dictionary")

    ' 筛选该物流单号 + SKU 的退单行
    Dim groupRows() As NormalizedReturnLine
    groupRows = E2E_FilterRows(orders, shipNo, sku)

    Dim plan As Object
    Set plan = BuildStaticPlan(groupRows, ledger)
    planMap.Add sku, plan

    Dim pr As PrecheckResult
    pr = RunPrecheck(plan, ledger)
    precheckMap.Add sku, Array(CBool(pr.PrecheckAHit), CBool(pr.PrecheckBHit))

    Dim skuList(1 To 1) As String
    skuList(1) = sku

    Set E2E_AllocateSingleSKU = AllocateShipment(shipNo, skuList, planMap, precheckMap, ledger, cfg)
End Function


' 向 planMap/precheckMap 添加指定物流单号+SKU 的静态计划和预检测结论。
' 设计意图：多 SKU 场景（T06、T09）需要为每个 SKU 分别构建 plan，此函数复用逻辑
' 避免重复代码；等价于 modRunner.bas 中的 RN_BuildSkuGroupsForShipment 单SKU版本。
Private Sub E2E_AddSkuToPlanMap( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipNo As String, _
    ByVal sku As String, _
    ByVal ledger As Object, _
    ByRef planMap As Object, _
    ByRef precheckMap As Object)

    Dim groupRows() As NormalizedReturnLine
    groupRows = E2E_FilterRows(orders, shipNo, sku)

    Dim plan As Object
    Set plan = BuildStaticPlan(groupRows, ledger)
    planMap.Add sku, plan

    Dim pr As PrecheckResult
    pr = RunPrecheck(plan, ledger)
    precheckMap.Add sku, Array(CBool(pr.PrecheckAHit), CBool(pr.PrecheckBHit))
End Sub


' 从 orders 数组中筛选指定物流单号和 SKU 的退单行，返回子数组。
' 两遍扫描（先计数后填充），避免动态增长数组的内存开销。
' 设计意图：等价于 modRunner.bas 中的 RN_FilterOrdersByShipmentSKU，
' 在测试侧重复定义以保持 modTestRunner 的独立性（不依赖 modRunner 私有函数）。
Private Function E2E_FilterRows( _
    ByRef orders() As NormalizedReturnLine, _
    ByVal shipNo As String, _
    ByVal sku As String) As NormalizedReturnLine()

    Dim total As Long
    On Error Resume Next
    total = UBound(orders) - LBound(orders) + 1
    On Error GoTo 0
    If total <= 0 Then Exit Function

    ' 第一遍：统计匹配行数
    Dim matchCount As Long
    Dim i As Long
    For i = LBound(orders) To LBound(orders) + total - 1
        If orders(i).ShipmentNo = shipNo And orders(i).SKU = sku Then
            matchCount = matchCount + 1
        End If
    Next i
    If matchCount = 0 Then Exit Function

    ' 第二遍：填充结果数组
    Dim result() As NormalizedReturnLine
    ReDim result(1 To matchCount)
    Dim idx As Long
    For i = LBound(orders) To LBound(orders) + total - 1
        If orders(i).ShipmentNo = shipNo And orders(i).SKU = sku Then
            idx = idx + 1
            result(idx) = orders(i)
        End If
    Next i

    E2E_FilterRows = result
End Function


' =============================================================================
' M16 批量测试运行器冒烟测试
' =============================================================================

Public Sub RunBatchRunnerSmokeTest()
    ' 覆盖：
    '   BT-01 三张批量表表头自动生成
    '   BT-02 启用批次读取与 DryRun 子批次执行
    '   BT-03 批量结果表写入
    BeginSuite "Batch Test Runner Smoke Tests"

    On Error GoTo CleanFail

    TestBT_EnsureBatchSheetHeaders
    TestBT_MinimalDryRunPipeline

    FinishSuite
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] Batch Test Runner Smoke Tests 异常：" & Err.Description
    FinishSuite
End Sub

' BT-01：EnsureBatchTestSheets 应写入标准表头。
Private Sub TestBT_EnsureBatchSheetHeaders()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)

    On Error GoTo CleanExit
    EnsureBatchTestSheets wb

    AssertEqualString "BT-01 计划表首列=批次ID", "批次ID", CStr(wb.Worksheets("批量测试计划").Cells(1, 1).Value)
    AssertEqualString "BT-01 结果表首列=批次ID", "批次ID", CStr(wb.Worksheets("批量测试结果").Cells(1, 1).Value)
    AssertEqualString "BT-01 断言表第3列=断言类型", "断言类型", CStr(wb.Worksheets("批量断言结果").Cells(1, 3).Value)

CleanExit:
    EOT_CloseWorkbook wb
End Sub

' 空输入（两张输入表只有表头）：全链路应优雅完成，不抛错；输出全空、历史追加 1 行 Full Run。
' 回归：2026-07-19 前 RN_RunAllAllocations 对未初始化数组直接求 LBound，空输入触发错误 9（下标越界）。
Private Sub TestRunner_EmptyInputFullRun()
    Dim wb As Workbook
    Set wb = BT_CreateSmokeSourceWorkbook()
    wb.Worksheets("输入_退单表").Rows("2:" & wb.Worksheets("输入_退单表").Rows.Count).ClearContents
    wb.Worksheets("输入_质检库存表").Rows("2:" & wb.Worksheets("输入_质检库存表").Rows.Count).ClearContents

    On Error GoTo CleanFail

    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet
    Dim wsAnomaly As Worksheet
    Dim wsHistory As Worksheet
    Set wsSummary = wb.Worksheets("分配状态汇总表")
    Set wsDetail = wb.Worksheets("成功分配明细表")
    Set wsAnomaly = wb.Worksheets("数据异常明细表")
    Set wsHistory = wb.Worksheets("运行历史记录表")
    wsHistory.Cells(1, 3).Value = "运行类型"

    RunFullAllocation wb, False

    AssertEqualLong "空输入 汇总数据行=0", 0, TC43_DataRowCount(wsSummary)
    AssertEqualLong "空输入 明细数据行=0", 0, TC43_DataRowCount(wsDetail)
    AssertEqualLong "空输入 异常数据行=0", 0, TC43_DataRowCount(wsAnomaly)
    AssertEqualString "空输入 历史追加FullRun", "Full Run", CStr(wsHistory.Cells(2, 3).Value)
    AssertEqualString "空输入 历史无第3行数据", vbNullString, CStr(wsHistory.Cells(3, 1).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub

CleanFail:
    Debug.Print "[ERROR] 空输入回归测试异常：" & Err.Description
    AssertTrue "空输入 全链路不抛错（" & Err.Description & "）", False
    Resume CleanExit
End Sub

' BT-02/03：最小输入 + 启用 DryRun 批次，应写入一条成功结果。
Private Sub TestBT_MinimalDryRunPipeline()
    Dim wb As Workbook
    Set wb = BT_CreateSmokeSourceWorkbook()

    On Error GoTo CleanExit

    EnsureBatchTestSheets wb
    BT_WriteSmokePlan wb

    ' 静默执行，避免自动化验收被批量结果弹窗阻塞。
    RunBatchTestPlan wb, False

    Dim wsResult As Worksheet
    Set wsResult = wb.Worksheets("批量测试结果")
    AssertEqualString "BT-02 结果批次ID", "BATCH-SMOKE-01", CStr(wsResult.Cells(2, 1).Value)
    AssertEqualString "BT-02 运行模式=DryRun", "DryRun", CStr(wsResult.Cells(2, 3).Value)
    AssertEqualString "BT-03 运行状态=成功", "成功", CStr(wsResult.Cells(2, 6).Value)

CleanExit:
    EOT_CloseWorkbook wb
    Exit Sub
End Sub

' 构造批量冒烟用的最小源工作簿（含输入表、配置表、输出表、运行历史表）。
Private Function BT_CreateSmokeSourceWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)

    wb.Worksheets(1).Name = "输入_退单表"
    wb.Worksheets(1).Cells(1, 1).Value = "物流单号"
    wb.Worksheets(1).Cells(1, 2).Value = "WMS退单号"
    wb.Worksheets(1).Cells(1, 3).Value = "SKU"
    wb.Worksheets(1).Cells(1, 4).Value = "行号"
    wb.Worksheets(1).Cells(1, 5).Value = "数量"
    wb.Worksheets(1).Cells(2, 1).Value = "SF_BT_SMOKE"
    wb.Worksheets(1).Cells(2, 2).Value = "TK_BT_SMOKE"
    wb.Worksheets(1).Cells(2, 3).Value = "H_BT_SMOKE"
    ' 行号必须是文本：直接 .Value 赋 "00001" 会被 Excel 强转成数值 1，
    ' 触发 E01（CStr(1) 不是五位文本）。先设文本格式再赋值。
    wb.Worksheets(1).Cells(2, 4).NumberFormat = "@"
    wb.Worksheets(1).Cells(2, 4).Value = "00001"
    wb.Worksheets(1).Cells(2, 5).Value = 1

    Dim wsInv As Worksheet
    Set wsInv = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsInv.Name = "输入_质检库存表"
    wsInv.Cells(1, 1).Value = "物流单号"
    wsInv.Cells(1, 2).Value = "SKU"
    wsInv.Cells(1, 3).Value = "QC情况"
    wsInv.Cells(1, 4).Value = "批号"
    wsInv.Cells(1, 5).Value = "效期"
    wsInv.Cells(1, 6).Value = "数量"
    wsInv.Cells(2, 1).Value = "SF_BT_SMOKE"
    wsInv.Cells(2, 2).Value = "H_BT_SMOKE"
    wsInv.Cells(2, 3).Value = "ZP"
    wsInv.Cells(2, 4).Value = "LA01"
    ' 效期同样先设文本格式，避免 Excel 按本机区域设置把字符串转成日期序列号，
    ' 保证测试在不同区域设置下行为一致（M04 文本分支确定性处理）。
    wsInv.Cells(2, 5).NumberFormat = "@"
    wsInv.Cells(2, 5).Value = "2029/01/01"
    wsInv.Cells(2, 6).Value = 1

    Dim wsCfg As Worksheet
    Set wsCfg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsCfg.Name = "输入_配置"
    wsCfg.Cells(1, 1).Value = "物流单号"
    wsCfg.Cells(1, 2).Value = "TC编号"
    wsCfg.Cells(1, 3).Value = "最大回溯次数"
    wsCfg.Cells(1, 4).Value = "调试日志级别"
    wsCfg.Cells(1, 5).Value = "批号比较模式"
    wsCfg.Cells(1, 6).Value = "无保质期哨兵值"
    wsCfg.Cells(2, 1).Value = "SF_BT_SMOKE"
    wsCfg.Cells(2, 2).Value = "BT-SMOKE"
    wsCfg.Cells(2, 3).Value = 10
    wsCfg.Cells(2, 4).Value = DEBUG_LEVEL_OFF
    wsCfg.Cells(2, 5).Value = LOT_MODE_INSENSITIVE
    wsCfg.Cells(2, 6).Value = DEFAULT_NO_EXPIRY_SENTINEL

    EOT_EnsureSheet wb, "分配状态汇总表"
    EOT_EnsureSheet wb, "成功分配明细表"
    EOT_EnsureSheet wb, "数据异常明细表"
    EOT_EnsureSheet wb, "调试日志"
    EOT_EnsureSheet wb, "运行历史记录表"

    Set BT_CreateSmokeSourceWorkbook = wb
End Function

' 向批量测试计划写入一条启用的 DryRun 批次。
Private Sub BT_WriteSmokePlan(ByVal wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("批量测试计划")

    ws.Cells(2, 1).Value = "BATCH-SMOKE-01"
    ws.Cells(2, 2).Value = "是"
    ws.Cells(2, 3).Value = "DryRun"
    ws.Cells(2, 4).Value = "SF_BT_SMOKE"
    ws.Cells(2, 5).Value = "按物流单号读取"
    ws.Cells(2, 6).Value = vbNullString
    ws.Cells(2, 7).Value = vbNullString
    ws.Cells(2, 8).Value = "批量运行器冒烟"
End Sub
