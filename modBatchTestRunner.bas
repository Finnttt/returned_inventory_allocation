    Option Explicit

    ' =============================================================================
    ' M17_批量回归扩展（modBatchTestRunner）
    ' =============================================================================
    ' 职责：从“批量测试计划”读取启用批次，按物流单号筛选输入、按配置拆分子批次，
    '       在临时工作簿中复用 RunValidationOnly / RunFullAllocation，再与预期表比对。
    ' 编号说明：2026-07-19 拍板为独立模块 M17（不并入 M16）；M16 保持“内存测试入口”定位。
    '
    ' 给新手的解释：
    '   单点按钮只能跑整本工作簿；批量运行器让你一次定义多组“跑哪些单号、用什么模式”，
    '   自动写回“批量测试结果”和“批量断言结果”，便于回归测试。
    ' =============================================================================

    Private Const SHEET_BATCH_PLAN   As String = "批量测试计划"
    Private Const SHEET_BATCH_RESULT As String = "批量测试结果"
    Private Const SHEET_BATCH_ASSERT As String = "批量断言结果"

    Private Const SHEET_RETURN_INPUT    As String = "输入_退单表"
    Private Const SHEET_INVENTORY_INPUT As String = "输入_质检库存表"
    Private Const SHEET_CONFIG          As String = "输入_配置"
    Private Const SHEET_RUN_HISTORY     As String = "运行历史记录表"

    Private Const SHEET_SUMMARY_ACTUAL As String = "分配状态汇总表"
    Private Const SHEET_DETAIL_ACTUAL  As String = "成功分配明细表"
    Private Const SHEET_ANOMALY_ACTUAL As String = "数据异常明细表"

    Private Const SHEET_SUMMARY_EXPECTED As String = "预期_汇总表"
    Private Const SHEET_DETAIL_EXPECTED  As String = "预期_成功分配明细"
    Private Const SHEET_ANOMALY_EXPECTED As String = "预期_数据异常明细"
    Private Const SHEET_ASSERT_EXPECTED  As String = "预期_断言"

    Private Const RUN_MODE_DRY  As String = "DryRun"
    Private Const RUN_MODE_FULL As String = "FullRun"

    Private Const CFG_BY_SHIPMENT As String = "按物流单号读取"
    Private Const CFG_FIXED_ROW   As String = "固定配置行"

    Private Const COL_BATCH_ID As String = "批次ID"
    Private Const COL_ENABLED  As String = "启用"
    Private Const COL_RUN_MODE As String = "运行模式"
    Private Const COL_SHIP_LIST As String = "物流单号列表"
    Private Const COL_CFG_STRATEGY As String = "配置策略"
    Private Const COL_FIXED_CFG_SHIP As String = "固定配置物流单号"
    Private Const COL_DEBUG_OVERRIDE As String = "调试日志级别覆盖"
    Private Const COL_REMARK As String = "备注"

    Private Const HEADER_BATCH_COL As String = "运行批次"

    ' 批量测试类型分层：
    ' 1) 正常成功类：期望 FullRun/DryRun 正常完成，并比对汇总/明细/历史。
    ' 2) 字段异常类：表结构正确，字段值非法，期望输出异常明细和失败汇总。
    ' 3) 结构异常类：表头错位、缺列等，期望读取阶段中止并断言错误信息。
    ' 4) 配置异常类：配置值非法或缺失，期望配置读取阶段中止并断言错误信息。
    ' 结构/配置异常不要强行写入“数据异常明细表”，否则会把输入读取层和业务输出层耦合在一起。
    Private Const EXPECTED_ERROR_ITEM As String = "期望错误"
    Private Const EXPECTED_ERROR_SNIPPET_PREFIX As String = "关键报错片段"

    Private Type BatchPlanEntry
        BatchId As String
        Enabled As Boolean
        RunMode As String
        ShipmentListText As String
        ConfigStrategy As String
        FixedConfigShipment As String
        DebugLevelOverride As String
        Remark As String
        RowIndex As Long
    End Type

    Private Type BatchSubRunSpec
        BatchId As String
        SubIndex As Long
        Shipments() As String
        ConfigShipment As String
    End Type

    Private Type BatchAssertRecord
        BatchId As String
        SubIndex As Long
        AssertType As String
        Locator As String
        FieldName As String
        ExpectedValue As String
        ActualValue As String
        Passed As Boolean
        Remark As String
    End Type


    ' =============================================================================
    ' 一、公开入口
    ' =============================================================================

    ' 批量测试主入口。默认对当前活动工作簿执行；自动化调用可关闭结果弹窗。
    Public Sub RunBatchTestPlan( _
        Optional ByVal wb As Workbook = Nothing, _
        Optional ByVal showMessages As Boolean = True)

        If wb Is Nothing Then Set wb = ActiveWorkbook
        If wb Is Nothing Then
            If showMessages Then MsgBox "未找到可运行的工作簿。", vbCritical
            Exit Sub
        End If

        On Error GoTo RunFail

        EnsureBatchTestSheets wb

        Dim plans() As BatchPlanEntry
        plans = BT_ReadEnabledPlans(wb)
        If BT_PlanCount(plans) = 0 Then
            If showMessages Then MsgBox "批量测试计划中没有启用=是的批次。", vbExclamation
            Exit Sub
        End If

        Dim resultWs As Worksheet
        Dim assertWs As Worksheet
        Set resultWs = wb.Worksheets(SHEET_BATCH_RESULT)
        Set assertWs = wb.Worksheets(SHEET_BATCH_ASSERT)

        BT_ClearResultData resultWs
        BT_ClearResultData assertWs

        Dim totalPass As Long
        Dim totalFail As Long
        Dim batchFailCount As Long

        Dim i As Long
        For i = LBound(plans) To UBound(plans)
            Dim plan As BatchPlanEntry
            plan = plans(i)
            Dim sourceStateNotice As String
            sourceStateNotice = BT_SourceFilterNotice(wb)

            Dim validateErr As String
            validateErr = BT_ValidatePlanEntry(wb, plan)
            If validateErr <> vbNullString Then
                BT_AppendBatchResult resultWs, plan.BatchId, 0, plan.RunMode, plan.ShipmentListText, _
                                    vbNullString, "配置错误", Now, Now, 0, 1, BT_CombineMessages(validateErr, sourceStateNotice)
                batchFailCount = batchFailCount + 1
                GoTo NextPlan
            End If

            Dim shipments() As String
            On Error GoTo PlanFail
            shipments = BT_ResolveShipmentsForPlan(wb, plan)

            Dim subRuns() As BatchSubRunSpec
            subRuns = BT_BuildSubRuns(wb, plan, shipments)
            On Error GoTo RunFail
            If BT_SubRunCount(subRuns) = 0 Then
                BT_AppendBatchResult resultWs, plan.BatchId, 0, plan.RunMode, plan.ShipmentListText, _
                                    vbNullString, "跳过", Now, Now, 0, 0, BT_CombineMessages("未生成任何子批次", sourceStateNotice)
                GoTo NextPlan
            End If

            Dim s As Long
            For s = LBound(subRuns) To UBound(subRuns)
                Dim subRun As BatchSubRunSpec
                subRun = subRuns(s)

                Dim startedAt As Date
                Dim endedAt As Date
                Dim runStatus As String
                Dim runError As String
                Dim passCount As Long
                Dim failCount As Long

                startedAt = Now
                runStatus = "成功"
                runError = vbNullString
                passCount = 0
                failCount = 0

                Dim tempWb As Workbook
                Set tempWb = Nothing

                On Error GoTo SubRunFail
                Set tempWb = BT_CreateTempWorkbook(wb, subRun, plan.DebugLevelOverride)
                BT_AssertRequiredTempSheets tempWb
                BT_RunSubBatch tempWb, plan.RunMode

                Dim asserts() As BatchAssertRecord
                If BT_HasExpectedErrorAssertions(wb, plan, subRun) Then
                    asserts = BT_RunExpectedErrorNotRaisedAssertions(wb, plan, subRun)
                Else
                    asserts = BT_RunAssertions(wb, tempWb, plan, subRun)
                End If
                BT_CountAsserts asserts, passCount, failCount
                BT_WriteAssertRecords assertWs, asserts

                If failCount > 0 Then
                    runStatus = "断言失败"
                    batchFailCount = batchFailCount + 1
                End If

                endedAt = Now
                BT_CloseTempWorkbook tempWb
                GoTo SubRunDone

    SubRunFail:
                endedAt = Now
                Dim failedErrNumber As Long
                Dim failedErrSource As String
                Dim failedErrDescription As String
                failedErrNumber = Err.Number
                failedErrSource = Err.Source
                failedErrDescription = Err.Description
                runError = BT_BuildSubRunErrorMessage(failedErrNumber, failedErrSource, failedErrDescription, tempWb)

                Dim errorAsserts() As BatchAssertRecord
                errorAsserts = BT_RunExpectedErrorAssertions(wb, plan, subRun, runError)
                BT_CountAsserts errorAsserts, passCount, failCount
                BT_WriteAssertRecords assertWs, errorAsserts

                If BT_AssertRecordCount(errorAsserts) > 0 Then
                    If failCount = 0 Then
                        runStatus = "期望错误通过"
                    Else
                        runStatus = "期望错误失败"
                        batchFailCount = batchFailCount + 1
                    End If
                Else
                    runStatus = "运行失败"
                    failCount = failCount + 1
                    batchFailCount = batchFailCount + 1
                End If
                On Error Resume Next
                BT_CloseTempWorkbook tempWb
                On Error GoTo RunFail

    SubRunDone:
                totalPass = totalPass + passCount
                totalFail = totalFail + failCount

                BT_AppendBatchResult resultWs, subRun.BatchId, subRun.SubIndex, plan.RunMode, _
                                    BT_JoinShipments(subRun.Shipments), subRun.ConfigShipment, _
                                    runStatus, startedAt, endedAt, passCount, failCount, BT_CombineMessages(runError, sourceStateNotice)
            Next s

            GoTo NextPlan

    PlanFail:
            Dim planErrText As String
            planErrText = BT_BuildPlanErrorMessage(Err.Number, Err.Source, Err.Description)

            Dim planSubRun As BatchSubRunSpec
            planSubRun.BatchId = plan.BatchId
            planSubRun.SubIndex = 0
            If BT_StringArrayCount(shipments) > 0 Then
                planSubRun.Shipments = shipments
            Else
                planSubRun.Shipments = BT_ParseShipmentList(plan.ShipmentListText)
            End If

            Dim planErrPass As Long
            Dim planErrFail As Long
            Dim planErrAsserts() As BatchAssertRecord
            planErrAsserts = BT_RunExpectedErrorAssertions(wb, plan, planSubRun, planErrText)
            BT_CountAsserts planErrAsserts, planErrPass, planErrFail
            BT_WriteAssertRecords assertWs, planErrAsserts

            If BT_AssertRecordCount(planErrAsserts) > 0 And planErrFail = 0 Then
                BT_AppendBatchResult resultWs, plan.BatchId, 0, plan.RunMode, plan.ShipmentListText, _
                                    vbNullString, "期望错误通过", Now, Now, planErrPass, 0, BT_CombineMessages(planErrText, sourceStateNotice)
                totalPass = totalPass + planErrPass
            Else
                If BT_AssertRecordCount(planErrAsserts) = 0 Then planErrFail = 1
                BT_AppendBatchResult resultWs, plan.BatchId, 0, plan.RunMode, plan.ShipmentListText, _
                                    vbNullString, "计划失败", Now, Now, planErrPass, planErrFail, BT_CombineMessages(planErrText, sourceStateNotice)
                totalPass = totalPass + planErrPass
                totalFail = totalFail + planErrFail
                batchFailCount = batchFailCount + 1
            End If
            Err.Clear
            On Error GoTo RunFail

    NextPlan:
        Next i

        If showMessages Then
            MsgBox "批量测试完成。" & vbNewLine & _
                "断言通过：" & totalPass & " 条" & vbNewLine & _
                "断言失败：" & totalFail & " 条" & vbNewLine & _
                "异常批次/子批次：" & batchFailCount & " 个", _
                IIf(totalFail > 0 Or batchFailCount > 0, vbExclamation, vbInformation)
        End If
        Exit Sub

    RunFail:
        If showMessages Then MsgBox "批量测试运行器异常：" & Err.Description, vbCritical
    End Sub


    ' 若三张批量表不存在或只有空表头，则自动创建标准表头，避免手工建表出错。
    Public Sub EnsureBatchTestSheets(ByVal wb As Workbook)
        If wb Is Nothing Then Err.Raise vbObjectError + 1600, "EnsureBatchTestSheets", "工作簿为空。"

        Dim wsPlan As Worksheet
        Set wsPlan = BT_EnsureSheet(wb, SHEET_BATCH_PLAN)
        If BT_SheetNeedsPlanHeader(wsPlan) Then BT_WritePlanHeaders wsPlan

        Dim wsResult As Worksheet
        Set wsResult = BT_EnsureSheet(wb, SHEET_BATCH_RESULT)
        If BT_SheetNeedsResultHeader(wsResult) Then BT_WriteResultHeaders wsResult

        Dim wsAssert As Worksheet
        Set wsAssert = BT_EnsureSheet(wb, SHEET_BATCH_ASSERT)
        If BT_SheetNeedsAssertHeader(wsAssert) Then BT_WriteAssertHeaders wsAssert
    End Sub

    ' 追加/更新推荐批次计划。默认全部写为“否”，由使用者按阶段手动启用。
    Public Sub SeedTestSystemBatchPlan( _
        Optional ByVal wb As Workbook = Nothing, _
        Optional ByVal showMessages As Boolean = True)

        If wb Is Nothing Then Set wb = ActiveWorkbook
        If wb Is Nothing Then
            If showMessages Then MsgBox "未找到可写入批量计划的工作簿。", vbCritical
            Exit Sub
        End If

        EnsureBatchTestSheets wb

        Dim ws As Worksheet
        Set ws = wb.Worksheets(SHEET_BATCH_PLAN)

        BT_UpsertPlanRow ws, "BATCH-ERR-STRUCTURE", "否", RUN_MODE_DRY, _
            "SF3190000000052;SF3190000000053;SF3190000000055", CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_OFF, "独立坏表头文件专用；汇总工作簿勿启用，改运行 RunSingleTest 17"

        BT_UpsertPlanRow ws, "BATCH-ERR-CONFIG", "否", RUN_MODE_DRY, _
            "SF3190000000054A;SF3190000000054B", CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_OFF, "A/B 是文件后缀而非物流单号；汇总工作簿勿启用，改运行 RunSingleTest 17"

        BT_UpsertPlanRow ws, "BATCH-05-DEBUG-SAMPLE", "否", RUN_MODE_FULL, _
            "SF3190000000016", CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_DETAIL, "详细日志抽样：只跑代表性单号，避免日志爆炸"

        ' 以下两个批次 ID 为兼容既有工作簿而保留 PENDING 字样；对应 TC 已完成自动化。
        BT_UpsertPlanRow ws, "BATCH-PENDING-TC-DRY", "否", RUN_MODE_DRY, _
            "SF3190000000063;SF3190000000064", CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_OFF, "回归TC：TC-41异常明细格式、TC-28多错误码原因格式"

        BT_UpsertPlanRow ws, "BATCH-PENDING-TC-FULL", "否", RUN_MODE_FULL, _
            "SF3190000000062", CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_SIMPLE, "回归TC：TC-40预检测B，核对E09与调试日志子类型"

        BT_UpsertPlanRow ws, "BATCH-99-ALL-DRY", "否", RUN_MODE_DRY, _
            vbNullString, CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_OFF, "全量DryRun：物流单号列表留空表示全部输入行"

        BT_UpsertPlanRow ws, "BATCH-100-ALL-FULL", "否", RUN_MODE_FULL, _
            vbNullString, CFG_BY_SHIPMENT, _
            vbNullString, DEBUG_LEVEL_OFF, "全量FullRun：最终回归，建议最后启用"

        If showMessages Then
            MsgBox "推荐批量测试计划已写入/更新。默认均为未启用，请按阶段手动改为“是”。", vbInformation
        End If
    End Sub


    ' =============================================================================
    ' 二、表头生成
    ' =============================================================================

    Private Sub BT_WritePlanHeaders(ByVal ws As Worksheet)
        ws.Cells(1, 1).Value = COL_BATCH_ID
        ws.Cells(1, 2).Value = COL_ENABLED
        ws.Cells(1, 3).Value = COL_RUN_MODE
        ws.Cells(1, 4).Value = COL_SHIP_LIST
        ws.Cells(1, 5).Value = COL_CFG_STRATEGY
        ws.Cells(1, 6).Value = COL_FIXED_CFG_SHIP
        ws.Cells(1, 7).Value = COL_DEBUG_OVERRIDE
        ws.Cells(1, 8).Value = COL_REMARK
    End Sub

    Private Sub BT_WriteResultHeaders(ByVal ws As Worksheet)
        ws.Cells(1, 1).Value = "批次ID"
        ws.Cells(1, 2).Value = "子批次序号"
        ws.Cells(1, 3).Value = "运行模式"
        ws.Cells(1, 4).Value = "物流单号列表"
        ws.Cells(1, 5).Value = "配置物流单号"
        ws.Cells(1, 6).Value = "运行状态"
        ws.Cells(1, 7).Value = "开始时间"
        ws.Cells(1, 8).Value = "结束时间"
        ws.Cells(1, 9).Value = "断言通过数"
        ws.Cells(1, 10).Value = "断言失败数"
        ws.Cells(1, 11).Value = "错误信息"
    End Sub

    Private Sub BT_WriteAssertHeaders(ByVal ws As Worksheet)
        ws.Cells(1, 1).Value = "批次ID"
        ws.Cells(1, 2).Value = "子批次序号"
        ws.Cells(1, 3).Value = "断言类型"
        ws.Cells(1, 4).Value = "定位键"
        ws.Cells(1, 5).Value = "字段名"
        ws.Cells(1, 6).Value = "预期值"
        ws.Cells(1, 7).Value = "实际值"
        ws.Cells(1, 8).Value = "是否通过"
        ws.Cells(1, 9).Value = "备注"
    End Sub

    Private Sub BT_UpsertPlanRow( _
        ByVal ws As Worksheet, _
        ByVal batchId As String, _
        ByVal enabledText As String, _
        ByVal runMode As String, _
        ByVal shipmentList As String, _
        ByVal configStrategy As String, _
        ByVal fixedConfigShipment As String, _
        ByVal debugOverride As String, _
        ByVal remark As String)

        Dim rowIndex As Long
        rowIndex = BT_FindPlanRow(ws, batchId)
        If rowIndex = 0 Then
            rowIndex = BT_GetLastUsedRow(ws, 1) + 1
            If rowIndex < 2 Then rowIndex = 2
        End If

        ws.Cells(rowIndex, 1).Value = batchId
        ws.Cells(rowIndex, 2).Value = enabledText
        ws.Cells(rowIndex, 3).Value = runMode
        ws.Cells(rowIndex, 4).Value = shipmentList
        ws.Cells(rowIndex, 5).Value = configStrategy
        ws.Cells(rowIndex, 6).Value = fixedConfigShipment
        ws.Cells(rowIndex, 7).Value = debugOverride
        ws.Cells(rowIndex, 8).Value = remark
    End Sub

    Private Function BT_FindPlanRow(ByVal ws As Worksheet, ByVal batchId As String) As Long
        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, 1)

        Dim r As Long
        For r = 2 To lastRow
            If Trim$(CStr(ws.Cells(r, 1).Value)) = batchId Then
                BT_FindPlanRow = r
                Exit Function
            End If
        Next r
    End Function


    ' =============================================================================
    ' 三、读取与校验批次计划
    ' =============================================================================

    Private Function BT_ReadEnabledPlans(ByVal wb As Workbook) As BatchPlanEntry()
        Dim ws As Worksheet
        Set ws = wb.Worksheets(SHEET_BATCH_PLAN)

        Dim colBatch As Long
        Dim colEnabled As Long
        Dim colMode As Long
        Dim colShipList As Long
        Dim colStrategy As Long
        Dim colFixedShip As Long
        Dim colDebugOverride As Long
        Dim colRemark As Long

        colBatch = BT_FindHeaderColumn(ws, COL_BATCH_ID, True)
        colEnabled = BT_FindHeaderColumn(ws, COL_ENABLED, True)
        colMode = BT_FindHeaderColumn(ws, COL_RUN_MODE, True)
        colShipList = BT_FindHeaderColumn(ws, COL_SHIP_LIST, False)
        colStrategy = BT_FindHeaderColumn(ws, COL_CFG_STRATEGY, False)
        colFixedShip = BT_FindHeaderColumn(ws, COL_FIXED_CFG_SHIP, False)
        colDebugOverride = BT_FindHeaderColumn(ws, COL_DEBUG_OVERRIDE, False)
        colRemark = BT_FindHeaderColumn(ws, COL_REMARK, False)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, colBatch)
        If lastRow < 2 Then Exit Function

        Dim temp() As BatchPlanEntry
        Dim count As Long

        Dim r As Long
        For r = 2 To lastRow
            Dim batchId As String
            batchId = Trim$(CStr(ws.Cells(r, colBatch).Value))
            If batchId = vbNullString Then GoTo NextRow
            If Not BT_IsEnabled(ws.Cells(r, colEnabled).Value) Then GoTo NextRow

            count = count + 1
            ReDim Preserve temp(1 To count)

            With temp(count)
                .BatchId = batchId
                .Enabled = True
                .RunMode = Trim$(CStr(ws.Cells(r, colMode).Value))
                If colShipList > 0 Then .ShipmentListText = Trim$(CStr(ws.Cells(r, colShipList).Value))
                If colStrategy > 0 Then .ConfigStrategy = Trim$(CStr(ws.Cells(r, colStrategy).Value))
                If colFixedShip > 0 Then .FixedConfigShipment = Trim$(CStr(ws.Cells(r, colFixedShip).Value))
                If colDebugOverride > 0 Then .DebugLevelOverride = Trim$(CStr(ws.Cells(r, colDebugOverride).Value))
                If colRemark > 0 Then .Remark = Trim$(CStr(ws.Cells(r, colRemark).Value))
                .RowIndex = r
            End With
    NextRow:
        Next r

        If count > 0 Then BT_ReadEnabledPlans = temp
    End Function

    Private Function BT_ValidatePlanEntry(ByVal wb As Workbook, ByRef plan As BatchPlanEntry) As String
        If plan.BatchId = vbNullString Then
            BT_ValidatePlanEntry = "批次ID 不能为空。"
            Exit Function
        End If

        If plan.BatchId = "BATCH-ERR-STRUCTURE" Or plan.BatchId = "BATCH-ERR-CONFIG" Then
            BT_ValidatePlanEntry = "该计划依赖独立异常工作簿，不能在标准汇总工作簿中执行；请运行 RunSingleTest 17。"
            Exit Function
        End If

        Dim modeUpper As String
        modeUpper = UCase$(plan.RunMode)
        If modeUpper <> UCase$(RUN_MODE_DRY) And modeUpper <> UCase$(RUN_MODE_FULL) Then
            BT_ValidatePlanEntry = "运行模式必须是 DryRun 或 FullRun，当前=[" & plan.RunMode & "]。"
            Exit Function
        End If
        plan.RunMode = IIf(modeUpper = UCase$(RUN_MODE_DRY), RUN_MODE_DRY, RUN_MODE_FULL)

        If Len(plan.ConfigStrategy) = 0 Then plan.ConfigStrategy = CFG_BY_SHIPMENT

        If plan.ConfigStrategy = CFG_FIXED_ROW Then
            If Len(plan.FixedConfigShipment) = 0 Then
                BT_ValidatePlanEntry = "配置策略为固定配置行时，固定配置物流单号不能为空。"
                Exit Function
            End If
        End If

        If Len(plan.DebugLevelOverride) > 0 Then
            If Not BT_IsValidDebugLevel(plan.DebugLevelOverride) Then
                BT_ValidatePlanEntry = "调试日志级别覆盖非法，允许：关闭/简版/详细。"
                Exit Function
            End If
        End If

        Dim tmpWs As Worksheet
        If Not BT_TryGetWorksheet(wb, SHEET_RETURN_INPUT, tmpWs) Then
            BT_ValidatePlanEntry = "缺少工作表 [" & SHEET_RETURN_INPUT & "]。"
            Exit Function
        End If

        Set tmpWs = Nothing
        If Not BT_TryGetWorksheet(wb, SHEET_INVENTORY_INPUT, tmpWs) Then
            BT_ValidatePlanEntry = "缺少工作表 [" & SHEET_INVENTORY_INPUT & "]。"
            Exit Function
        End If

        Set tmpWs = Nothing
        If Not BT_TryGetWorksheet(wb, SHEET_CONFIG, tmpWs) Then
            BT_ValidatePlanEntry = "缺少工作表 [" & SHEET_CONFIG & "]。"
            Exit Function
        End If
    End Function


    ' =============================================================================
    ' 四、子批次拆分与临时工作簿
    ' =============================================================================

    Private Function BT_ResolveShipmentsForPlan(ByVal wb As Workbook, ByRef plan As BatchPlanEntry) As String()
        Dim listed() As String
        listed = BT_ParseShipmentList(plan.ShipmentListText)

        If BT_StringArrayCount(listed) > 0 Then
            BT_ResolveShipmentsForPlan = listed
            Exit Function
        End If

        BT_ResolveShipmentsForPlan = BT_CollectShipmentsFromReturnSheet(wb.Worksheets(SHEET_RETURN_INPUT))
    End Function

    Private Function BT_BuildSubRuns( _
        ByVal wb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef shipments() As String) As BatchSubRunSpec()

        If BT_StringArrayCount(shipments) = 0 Then Exit Function

        If plan.ConfigStrategy = CFG_FIXED_ROW Then
            Dim one(1 To 1) As BatchSubRunSpec
            one(1).BatchId = plan.BatchId
            one(1).SubIndex = 1
            one(1).Shipments = shipments
            one(1).ConfigShipment = plan.FixedConfigShipment
            BT_BuildSubRuns = one
            Exit Function
        End If

        ' 按物流单号读取配置：相同配置指纹的物流单号合并为一个子批次。
        Dim cfgWs As Worksheet
        Set cfgWs = wb.Worksheets(SHEET_CONFIG)

        Dim groups As Object
        Set groups = CreateObject("Scripting.Dictionary")
        groups.CompareMode = vbTextCompare

        Dim i As Long
        For i = LBound(shipments) To UBound(shipments)
            Dim shipNo As String
            shipNo = Trim$(shipments(i))
            If shipNo = vbNullString Then GoTo NextShip

            Dim cfg As ConfigStruct
            cfg = LoadConfigForShipment(cfgWs, shipNo)
            Dim fp As String
            fp = BT_ConfigFingerprint(cfg)

            If Not groups.Exists(fp) Then
                Dim pack(0 To 1) As Variant
                pack(0) = shipNo
                pack(1) = shipNo
                groups.Add fp, pack
            Else
                Dim existing As Variant
                existing = groups(fp)
                existing(0) = existing(0) & "|" & shipNo
                If shipNo < existing(1) Then existing(1) = shipNo
                groups(fp) = existing
            End If
    NextShip:
        Next i

        If groups.Count = 0 Then Exit Function

        Dim result() As BatchSubRunSpec
        ReDim result(1 To groups.Count)

        Dim idx As Long
        Dim key As Variant
        For Each key In groups.Keys
            idx = idx + 1
            Dim item As Variant
            item = groups(key)

            result(idx).BatchId = plan.BatchId
            result(idx).SubIndex = idx
            result(idx).Shipments = BT_ParseShipmentList(CStr(item(0)), "|")
            result(idx).ConfigShipment = CStr(item(1))
        Next key

        BT_BuildSubRuns = result
    End Function

    Private Function BT_CreateTempWorkbook( _
        ByVal sourceWb As Workbook, _
        ByRef subRun As BatchSubRunSpec, _
        ByVal debugLevelOverride As String) As Workbook

        Dim tempWb As Workbook
        Set tempWb = Workbooks.Add(xlWBATWorksheet)

        BT_EnsureOutputSheets tempWb
        BT_CopyFilteredReturnSheet sourceWb.Worksheets(SHEET_RETURN_INPUT), tempWb, subRun.Shipments
        BT_CopyFilteredInventorySheet sourceWb.Worksheets(SHEET_INVENTORY_INPUT), tempWb, subRun.Shipments
        BT_CopyConfigForShipment sourceWb.Worksheets(SHEET_CONFIG), tempWb, subRun.ConfigShipment, debugLevelOverride

        Set BT_CreateTempWorkbook = tempWb
    End Function

    Private Sub BT_EnsureOutputSheets(ByVal wb As Workbook)
        wb.Worksheets(1).Name = SHEET_SUMMARY_ACTUAL
        BT_EnsureSheet wb, SHEET_DETAIL_ACTUAL
        BT_EnsureSheet wb, SHEET_ANOMALY_ACTUAL
        BT_EnsureSheet wb, SHEET_CONFIG
        BT_EnsureSheet wb, SHEET_RETURN_INPUT
        BT_EnsureSheet wb, SHEET_INVENTORY_INPUT
        BT_EnsureSheet wb, SHEET_RUN_HISTORY
        BT_EnsureSheet wb, "调试日志"

        Dim wsHist As Worksheet
        Set wsHist = wb.Worksheets(SHEET_RUN_HISTORY)
        If Trim$(CStr(wsHist.Cells(1, 1).Value)) = vbNullString Then
            ' 20 列：需求 §5.6 的 17 字段 + 3 个配置快照字段（2026-07-19 拍板口径）
            wsHist.Cells(1, 1).Value = "运行编号"
            wsHist.Cells(1, 2).Value = "运行时间"
            wsHist.Cells(1, 3).Value = "运行类型"
            wsHist.Cells(1, 4).Value = "输入：退单表行数"
            wsHist.Cells(1, 5).Value = "输入：质检库存表行数"
            wsHist.Cells(1, 6).Value = "输入：物流单号数"
            wsHist.Cells(1, 7).Value = "校验耗时（秒）"
            wsHist.Cells(1, 8).Value = "分配耗时（秒）"
            wsHist.Cells(1, 9).Value = "总耗时（秒）"
            wsHist.Cells(1, 10).Value = "校验失败物流单号数"
            wsHist.Cells(1, 11).Value = "分配成功物流单号数"
            wsHist.Cells(1, 12).Value = "分配失败物流单号数"
            wsHist.Cells(1, 13).Value = "错误码分布"
            wsHist.Cells(1, 14).Value = "总回溯次数"
            wsHist.Cells(1, 15).Value = "最大单组回溯次数"
            wsHist.Cells(1, 16).Value = "调试日志级别"
            wsHist.Cells(1, 17).Value = "备注"
            wsHist.Cells(1, 18).Value = "最大回溯次数"
            wsHist.Cells(1, 19).Value = "批号比较模式"
            wsHist.Cells(1, 20).Value = "无保质期哨兵值"
        End If
    End Sub

    Private Sub BT_CopyFilteredReturnSheet( _
        ByVal sourceWs As Worksheet, _
        ByVal tempWb As Workbook, _
        ByRef shipments() As String)

        Dim targetWs As Worksheet
        Set targetWs = tempWb.Worksheets(SHEET_RETURN_INPUT)

        BT_CopyHeaderRow sourceWs, targetWs

        Dim shipSet As Object
        Set shipSet = BT_ToShipmentSet(shipments)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(sourceWs, 1)
        If lastRow < 2 Then Exit Sub

        Dim outRow As Long
        outRow = 1
        Dim r As Long
        For r = 2 To lastRow
            Dim shipNo As String
            shipNo = Trim$(CStr(sourceWs.Cells(r, 1).Value))
            If shipSet.Exists(shipNo) Then
                outRow = outRow + 1
                BT_CopyUsedRowCells sourceWs, r, targetWs, outRow
            End If
        Next r
    End Sub

    Private Sub BT_CopyFilteredInventorySheet( _
        ByVal sourceWs As Worksheet, _
        ByVal tempWb As Workbook, _
        ByRef shipments() As String)

        Dim targetWs As Worksheet
        Set targetWs = tempWb.Worksheets(SHEET_INVENTORY_INPUT)

        BT_CopyHeaderRow sourceWs, targetWs

        Dim shipSet As Object
        Set shipSet = BT_ToShipmentSet(shipments)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(sourceWs, 1)
        If lastRow < 2 Then Exit Sub

        Dim outRow As Long
        outRow = 1
        Dim r As Long
        For r = 2 To lastRow
            Dim shipNo As String
            shipNo = Trim$(CStr(sourceWs.Cells(r, 1).Value))
            If shipSet.Exists(shipNo) Then
                outRow = outRow + 1
                BT_CopyUsedRowCells sourceWs, r, targetWs, outRow
            End If
        Next r
    End Sub

    Private Sub BT_CopyConfigForShipment( _
        ByVal sourceWs As Worksheet, _
        ByVal tempWb As Workbook, _
        ByVal configShipment As String, _
        ByVal debugLevelOverride As String)

        Dim targetWs As Worksheet
        Set targetWs = tempWb.Worksheets(SHEET_CONFIG)

        Dim cfgRow As Long
        cfgRow = BT_FindConfigRow(sourceWs, configShipment)

        BT_CopyUsedRowCells sourceWs, 1, targetWs, 1
        BT_CopyUsedRowCells sourceWs, cfgRow, targetWs, 2

        If Len(debugLevelOverride) > 0 Then
            Dim debugCol As Long
            debugCol = BT_FindHeaderColumn(targetWs, "调试日志级别", True)
            targetWs.Cells(2, debugCol).Value = debugLevelOverride
        End If
    End Sub

    Private Sub BT_RunSubBatch(ByVal tempWb As Workbook, ByVal runMode As String)
        If runMode = RUN_MODE_DRY Then
            RunValidationOnlySilent tempWb
        Else
            RunFullAllocationSilent tempWb
        End If
    End Sub

    Private Sub BT_AssertRequiredTempSheets(ByVal tempWb As Workbook)
        Dim missing As String
        missing = BT_MissingRequiredTempSheets(tempWb)
        If missing <> vbNullString Then
            Err.Raise vbObjectError + 1610, "BT_AssertRequiredTempSheets", _
                    "临时工作簿缺少必要工作表：" & missing & "；当前工作表=" & BT_WorksheetNameList(tempWb)
        End If
    End Sub

    Private Function BT_MissingRequiredTempSheets(ByVal tempWb As Workbook) As String
        Dim requiredNames As Variant
        requiredNames = Array( _
            SHEET_RETURN_INPUT, _
            SHEET_INVENTORY_INPUT, _
            SHEET_CONFIG, _
            SHEET_SUMMARY_ACTUAL, _
            SHEET_DETAIL_ACTUAL, _
            SHEET_ANOMALY_ACTUAL, _
            "调试日志", _
            SHEET_RUN_HISTORY)

        Dim parts As String
        Dim i As Long
        For i = LBound(requiredNames) To UBound(requiredNames)
            Dim ws As Worksheet
            If Not BT_TryGetWorksheet(tempWb, CStr(requiredNames(i)), ws) Then
                If parts <> vbNullString Then parts = parts & ","
                parts = parts & CStr(requiredNames(i))
            End If
            Set ws = Nothing
        Next i

        BT_MissingRequiredTempSheets = parts
    End Function

    Private Function BT_BuildSubRunErrorMessage( _
        ByVal errNumber As Long, _
        ByVal errSource As String, _
        ByVal errDescription As String, _
        ByVal tempWb As Workbook) As String

        Dim msg As String
        msg = "错误号=" & CStr(errNumber) & _
            "；来源=" & errSource & _
            "；说明=" & errDescription

        If Not tempWb Is Nothing Then
            msg = msg & "；临时工作簿Sheets=" & BT_WorksheetNameList(tempWb)
        Else
            msg = msg & "；临时工作簿=未创建"
        End If

        BT_BuildSubRunErrorMessage = msg
    End Function

    Private Function BT_BuildPlanErrorMessage( _
        ByVal errNumber As Long, _
        ByVal errSource As String, _
        ByVal errDescription As String) As String

        BT_BuildPlanErrorMessage = "错误号=" & CStr(errNumber) & _
                                "；来源=" & errSource & _
                                "；说明=" & errDescription
    End Function

    Private Function BT_WorksheetNameList(ByVal wb As Workbook) As String
        If wb Is Nothing Then Exit Function

        Dim names As String
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If names <> vbNullString Then names = names & ","
            names = names & ws.Name
        Next ws

        BT_WorksheetNameList = names
    End Function

    Private Function BT_SourceFilterNotice(ByVal wb As Workbook) As String
        If wb Is Nothing Then Exit Function

        Dim sheetNames As Variant
        sheetNames = Array(SHEET_RETURN_INPUT, SHEET_INVENTORY_INPUT, SHEET_CONFIG)

        Dim notice As String
        Dim i As Long
        For i = LBound(sheetNames) To UBound(sheetNames)
            Dim ws As Worksheet
            If BT_TryGetWorksheet(wb, CStr(sheetNames(i)), ws) Then
                If ws.AutoFilterMode Or ws.FilterMode Then
                    If notice <> vbNullString Then notice = notice & ","
                    notice = notice & ws.Name
                End If
            End If
            Set ws = Nothing
        Next i

        If notice <> vbNullString Then
            BT_SourceFilterNotice = "源表存在筛选状态（已按单元格读取，不依赖可见行）：" & notice
        End If
    End Function

    Private Function BT_CombineMessages(ByVal firstText As String, ByVal secondText As String) As String
        If Trim$(firstText) = vbNullString Then
            BT_CombineMessages = secondText
        ElseIf Trim$(secondText) = vbNullString Then
            BT_CombineMessages = firstText
        Else
            BT_CombineMessages = firstText & "；" & secondText
        End If
    End Function


    ' =============================================================================
    ' 五、断言（第一版：汇总/明细/异常/运行历史 + 预期_断言）
    ' =============================================================================

    Private Function BT_RunAssertions( _
        ByVal sourceWb As Workbook, _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec) As BatchAssertRecord()

        Dim records() As BatchAssertRecord
        Dim count As Long

        ' DryRun 只做输入校验，不产生完整分配明细；避免拿 FullRun 预期误判。
        If plan.RunMode = RUN_MODE_FULL Then
            BT_AssertSummary sourceWb, tempWb, plan, subRun, records, count
            BT_AssertDetail sourceWb, tempWb, plan, subRun, records, count
        End If
        BT_AssertAnomaly sourceWb, tempWb, plan, subRun, records, count
        BT_AssertRunHistory tempWb, plan, subRun, records, count
        BT_AssertExpectedSheet sourceWb, tempWb, plan, subRun, records, count

        If count > 0 Then
            ReDim Preserve records(1 To count)
            BT_RunAssertions = records
        End If
    End Function

    Private Function BT_HasExpectedErrorAssertions( _
        ByVal sourceWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec) As Boolean

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_ASSERT_EXPECTED, wsExp) Then Exit Function

        Dim colBatch As Long
        Dim colShip As Long
        Dim colItem As Long
        Dim colExpected As Long
        BT_GetExpectedAssertColumns wsExp, colBatch, colShip, colItem, colExpected

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsExp, colItem)
        If lastRow < 2 Then Exit Function

        Dim r As Long
        For r = 2 To lastRow
            If BT_ExpectedAssertRowApplies(wsExp, r, colBatch, colShip, plan, subRun) Then
                If BT_IsExpectedErrorAssertItem(CStr(wsExp.Cells(r, colItem).Value)) Then
                    BT_HasExpectedErrorAssertions = True
                    Exit Function
                End If
            End If
        Next r
    End Function

    Private Function BT_RunExpectedErrorAssertions( _
        ByVal sourceWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByVal errorText As String) As BatchAssertRecord()

        Dim records() As BatchAssertRecord
        Dim count As Long

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_ASSERT_EXPECTED, wsExp) Then Exit Function

        Dim colBatch As Long
        Dim colShip As Long
        Dim colItem As Long
        Dim colExpected As Long
        BT_GetExpectedAssertColumns wsExp, colBatch, colShip, colItem, colExpected

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsExp, colItem)
        If lastRow < 2 Then Exit Function

        Dim r As Long
        For r = 2 To lastRow
            If Not BT_ExpectedAssertRowApplies(wsExp, r, colBatch, colShip, plan, subRun) Then GoTo NextRow

            Dim itemName As String
            Dim expectedValue As String
            itemName = Trim$(CStr(wsExp.Cells(r, colItem).Value))
            expectedValue = Trim$(CStr(wsExp.Cells(r, colExpected).Value))
            If Not BT_IsExpectedErrorAssertItem(itemName) Then GoTo NextRow
            If expectedValue = vbNullString Then GoTo NextRow

            Dim passed As Boolean
            If StrComp(itemName, EXPECTED_ERROR_ITEM, vbTextCompare) = 0 Then
                passed = (Len(Trim$(errorText)) > 0)
            Else
                passed = (InStr(1, errorText, expectedValue, vbTextCompare) > 0)
            End If

            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "期望错误", itemName, _
                            itemName, expectedValue, errorText, _
                            passed, vbNullString
    NextRow:
        Next r

        If count > 0 Then BT_RunExpectedErrorAssertions = records
    End Function

    Private Function BT_RunExpectedErrorNotRaisedAssertions( _
        ByVal sourceWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec) As BatchAssertRecord()

        Dim records() As BatchAssertRecord
        Dim count As Long

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_ASSERT_EXPECTED, wsExp) Then Exit Function

        Dim colBatch As Long
        Dim colShip As Long
        Dim colItem As Long
        Dim colExpected As Long
        BT_GetExpectedAssertColumns wsExp, colBatch, colShip, colItem, colExpected

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsExp, colItem)
        Dim r As Long
        For r = 2 To lastRow
            If Not BT_ExpectedAssertRowApplies(wsExp, r, colBatch, colShip, plan, subRun) Then GoTo NextRow

            Dim itemName As String
            Dim expectedValue As String
            itemName = Trim$(CStr(wsExp.Cells(r, colItem).Value))
            expectedValue = Trim$(CStr(wsExp.Cells(r, colExpected).Value))
            If Not BT_IsExpectedErrorAssertItem(itemName) Then GoTo NextRow

            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "期望错误", itemName, _
                            itemName, expectedValue, "运行成功，未抛出预期错误", False, vbNullString
    NextRow:
        Next r

        If count > 0 Then BT_RunExpectedErrorNotRaisedAssertions = records
    End Function

    Private Sub BT_GetExpectedAssertColumns( _
        ByVal wsExp As Worksheet, _
        ByRef colBatch As Long, _
        ByRef colShip As Long, _
        ByRef colItem As Long, _
        ByRef colExpected As Long)

        colBatch = BT_FindHeaderColumn(wsExp, HEADER_BATCH_COL, False)
        colShip = BT_FindHeaderColumn(wsExp, "物流单号", False)
        colItem = BT_FindHeaderColumn(wsExp, "断言项", False)
        colExpected = BT_FindHeaderColumn(wsExp, "预期", False)
        If colExpected <= 0 Then colExpected = BT_FindHeaderColumn(wsExp, "期望值", False)

        If colItem <= 0 Then colItem = 1
        If colExpected <= 0 Then colExpected = 2
    End Sub

    Private Function BT_ExpectedAssertRowApplies( _
        ByVal wsExp As Worksheet, _
        ByVal rowIndex As Long, _
        ByVal colBatch As Long, _
        ByVal colShip As Long, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec) As Boolean

        If colBatch > 0 Then
            If Trim$(CStr(wsExp.Cells(rowIndex, colBatch).Value)) = plan.BatchId Then
                BT_ExpectedAssertRowApplies = True
                Exit Function
            End If
        End If

        If colShip > 0 Then
            Dim shipNo As String
            shipNo = Trim$(CStr(wsExp.Cells(rowIndex, colShip).Value))
            If shipNo <> vbNullString And BT_ShipmentInList(shipNo, subRun.Shipments) Then
                BT_ExpectedAssertRowApplies = True
                Exit Function
            End If
        End If

        BT_ExpectedAssertRowApplies = (colBatch <= 0 And colShip <= 0)
    End Function

    Private Function BT_IsExpectedErrorAssertItem(ByVal itemName As String) As Boolean
        Dim textValue As String
        textValue = Trim$(itemName)
        BT_IsExpectedErrorAssertItem = _
            (StrComp(textValue, EXPECTED_ERROR_ITEM, vbTextCompare) = 0) Or _
            (InStr(1, textValue, EXPECTED_ERROR_SNIPPET_PREFIX, vbTextCompare) = 1)
    End Function

    Private Sub BT_AssertSummary( _
        ByVal sourceWb As Workbook, _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_SUMMARY_EXPECTED, wsExp) Then Exit Sub

        Dim wsAct As Worksheet
        Set wsAct = tempWb.Worksheets(SHEET_SUMMARY_ACTUAL)

        Dim expMap As Object
        Set expMap = BT_BuildExpectedRowMap(wsExp, plan.BatchId, subRun.Shipments, _
                                            Array("物流单号", "WMS退单号"), _
                                            Array("退单号状态", "原因"))

        Dim actMap As Object
        Set actMap = BT_BuildActualRowMap(wsAct, Array("物流单号", "WMS退单号"), _
                                        Array("退单号状态", "原因"))

        BT_CompareRowMaps "汇总表", plan.BatchId, subRun.SubIndex, expMap, actMap, records, count
    End Sub

    Private Sub BT_AssertDetail( _
        ByVal sourceWb As Workbook, _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_DETAIL_EXPECTED, wsExp) Then Exit Sub

        Dim wsAct As Worksheet
        Set wsAct = tempWb.Worksheets(SHEET_DETAIL_ACTUAL)

        Dim expMap As Object
        Set expMap = BT_BuildExpectedRowMap(wsExp, plan.BatchId, subRun.Shipments, _
                                            Array("物流单号", "WMS退单号", "SKU", "行号", _
                                                  "QC情况", "批号", "效期"), _
                                            Array("退单数量", "分配数量", "行状态", "退单号状态"))

        Dim actMap As Object
        Set actMap = BT_BuildActualRowMap(wsAct, _
                                        Array("物流单号", "WMS退单号", "SKU", "行号", _
                                              "QC情况", "批号", "效期"), _
                                        Array("退单数量", "分配数量", "行状态", "退单号状态"))

        BT_CompareRowMaps "明细表", plan.BatchId, subRun.SubIndex, expMap, actMap, records, count
    End Sub

    Private Sub BT_AssertAnomaly( _
        ByVal sourceWb As Workbook, _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim wsExp As Worksheet
        If Not BT_TryGetExpectedAnomalySheet(sourceWb, wsExp) Then Exit Sub

        Dim wsAct As Worksheet
        Set wsAct = tempWb.Worksheets(SHEET_ANOMALY_ACTUAL)

        Dim expCodes As Object
        Set expCodes = BT_CollectExpectedCodes(wsExp, plan.BatchId, subRun.Shipments)
        Dim actCodes As Object
        Set actCodes = BT_CollectActualCodes(wsAct)

        Dim key As Variant
        For Each key In expCodes.Keys
            Dim codeKey As String
            codeKey = CStr(key)
            Dim passed As Boolean
            passed = actCodes.Exists(codeKey)
            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "异常表", codeKey, _
                            "错误码出现", codeKey, BT_CodePresentText(actCodes, codeKey), passed, vbNullString
        Next key

        For Each key In actCodes.Keys
            codeKey = CStr(key)
            If Not expCodes.Exists(codeKey) Then
                BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "异常表", codeKey, _
                                "错误码出现", "(不应出现)", codeKey, False, _
                                "实际异常表出现预期未定义的物流单号+错误码组合"
            End If
        Next key
    End Sub

    Private Sub BT_AssertRunHistory( _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim wsHist As Worksheet
        If Not BT_TryGetRunHistorySheet(tempWb, wsHist) Then Exit Sub

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsHist, 1)
        If lastRow < 2 Then Exit Sub

        Dim expectedRunType As String
        expectedRunType = IIf(plan.RunMode = RUN_MODE_DRY, "Dry Run", "Full Run")
        Dim actualRunType As String
        actualRunType = Trim$(CStr(wsHist.Cells(lastRow, 3).Value))

        BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "运行历史", "最新行", _
                        "运行类型", expectedRunType, actualRunType, (expectedRunType = actualRunType), vbNullString

        If plan.RunMode = RUN_MODE_FULL Then
            Dim successCount As Long
            Dim failCount As Long
            Dim backtrackCount As Long
            successCount = BT_SafeCLng(wsHist.Cells(lastRow, 11).Value)
            failCount = BT_SafeCLng(wsHist.Cells(lastRow, 12).Value)
            backtrackCount = BT_SafeCLng(wsHist.Cells(lastRow, 14).Value)

            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "运行历史", "最新行", _
                            "分配成功数>=0", ">=0", CStr(successCount), (successCount >= 0), vbNullString
            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "运行历史", "最新行", _
                            "分配失败数>=0", ">=0", CStr(failCount), (failCount >= 0), vbNullString
            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "运行历史", "最新行", _
                            "总回溯次数>=0", ">=0", CStr(backtrackCount), (backtrackCount >= 0), vbNullString
        End If
    End Sub

    Private Sub BT_AssertExpectedSheet( _
        ByVal sourceWb As Workbook, _
        ByVal tempWb As Workbook, _
        ByRef plan As BatchPlanEntry, _
        ByRef subRun As BatchSubRunSpec, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim wsExp As Worksheet
        If Not BT_TryGetWorksheet(sourceWb, SHEET_ASSERT_EXPECTED, wsExp) Then Exit Sub

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsExp, 1)
        If lastRow < 2 Then Exit Sub

        Dim colBatch As Long
        Dim colShip As Long
        Dim colItem As Long
        Dim colExpected As Long

        colBatch = BT_FindHeaderColumn(wsExp, HEADER_BATCH_COL, False)
        colShip = BT_FindHeaderColumn(wsExp, "物流单号", False)
        colItem = BT_FindHeaderColumn(wsExp, "断言项", False)
        colExpected = BT_FindHeaderColumn(wsExp, "预期", False)
        If colExpected <= 0 Then colExpected = BT_FindHeaderColumn(wsExp, "期望值", False)

        ' 兼容旧的两列表：断言项 | 预期。没有物流单号时只能用于单工作簿结构异常测试。
        If colItem <= 0 Then colItem = 1
        If colExpected <= 0 Then colExpected = 2

        Dim r As Long
        For r = 2 To lastRow
            If colBatch > 0 Then
                If Trim$(CStr(wsExp.Cells(r, colBatch).Value)) <> plan.BatchId Then GoTo NextAssertRow
            End If

            If colShip > 0 Then
                Dim shipNo As String
                shipNo = Trim$(CStr(wsExp.Cells(r, colShip).Value))
                If Len(shipNo) > 0 Then
                    If Not BT_ShipmentInList(shipNo, subRun.Shipments) Then GoTo NextAssertRow
                End If
            End If

            Dim itemName As String
            Dim expectedValue As String
            itemName = Trim$(CStr(wsExp.Cells(r, colItem).Value))
            expectedValue = Trim$(CStr(wsExp.Cells(r, colExpected).Value))
            If itemName = vbNullString Then GoTo NextAssertRow

            Dim actualValue As String
            Dim passed As Boolean
            actualValue = BT_ResolveStructuredAssert(tempWb, plan.RunMode, itemName)
            passed = BT_CompareAssertValue(expectedValue, actualValue, itemName)
            BT_AddAssertRecord records, count, plan.BatchId, subRun.SubIndex, "预期断言", itemName, _
                            itemName, expectedValue, actualValue, passed, vbNullString
    NextAssertRow:
        Next r
    End Sub


    ' =============================================================================
    ' 六、结果写入
    ' =============================================================================

    Private Sub BT_WriteAssertRecords(ByVal ws As Worksheet, ByRef records() As BatchAssertRecord)
        Dim n As Long
        n = BT_AssertRecordCount(records)
        If n <= 0 Then Exit Sub

        Dim startRow As Long
        startRow = BT_GetLastUsedRow(ws, 1) + 1
        If startRow < 2 Then startRow = 2

        ' 预期值/实际值按文本写入，避免 Excel 把 2029/01/01 显示成 2029/1/1。
        ws.Columns(6).NumberFormat = "@"
        ws.Columns(7).NumberFormat = "@"

        Dim i As Long
        For i = 1 To n
            ws.Cells(startRow + i - 1, 1).Value = records(i).BatchId
            ws.Cells(startRow + i - 1, 2).Value = records(i).SubIndex
            ws.Cells(startRow + i - 1, 3).Value = records(i).AssertType
            ws.Cells(startRow + i - 1, 4).Value = records(i).Locator
            ws.Cells(startRow + i - 1, 5).Value = records(i).FieldName
            ws.Cells(startRow + i - 1, 6).Value = CStr(records(i).ExpectedValue)
            ws.Cells(startRow + i - 1, 7).Value = CStr(records(i).ActualValue)
            ws.Cells(startRow + i - 1, 8).Value = IIf(records(i).Passed, "是", "否")
            ws.Cells(startRow + i - 1, 9).Value = records(i).Remark
        Next i
    End Sub

    Private Sub BT_AppendBatchResult( _
        ByVal ws As Worksheet, _
        ByVal batchId As String, _
        ByVal subIndex As Long, _
        ByVal runMode As String, _
        ByVal shipList As String, _
        ByVal configShipment As String, _
        ByVal runStatus As String, _
        ByVal startedAt As Date, _
        ByVal endedAt As Date, _
        ByVal passCount As Long, _
        ByVal failCount As Long, _
        ByVal errText As String)

        Dim rowIndex As Long
        rowIndex = BT_GetLastUsedRow(ws, 1) + 1
        If rowIndex < 2 Then rowIndex = 2

        ws.Cells(rowIndex, 1).Value = batchId
        ws.Cells(rowIndex, 2).Value = subIndex
        ws.Cells(rowIndex, 3).Value = runMode
        ws.Cells(rowIndex, 4).Value = shipList
        ws.Cells(rowIndex, 5).Value = configShipment
        ws.Cells(rowIndex, 6).Value = runStatus
        ws.Cells(rowIndex, 7).Value = startedAt
        ws.Cells(rowIndex, 8).Value = endedAt
        ws.Cells(rowIndex, 9).Value = passCount
        ws.Cells(rowIndex, 10).Value = failCount
        ws.Cells(rowIndex, 11).Value = errText
    End Sub


    ' =============================================================================
    ' 七、断言辅助
    ' =============================================================================

    Private Sub BT_CompareRowMaps( _
        ByVal assertType As String, _
        ByVal batchId As String, _
        ByVal subIndex As Long, _
        ByVal expMap As Object, _
        ByVal actMap As Object, _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long)

        Dim key As Variant
        For Each key In expMap.Keys
            Dim locator As String
            locator = CStr(key)
            Dim expFields As Object
            Set expFields = expMap(key)

            Dim actFields As Object
            If actMap.Exists(locator) Then
                Set actFields = actMap(locator)
            Else
                Set actFields = CreateObject("Scripting.Dictionary")
            End If

            Dim fieldKey As Variant
            For Each fieldKey In expFields.Keys
                Dim fieldName As String
                fieldName = CStr(fieldKey)
                Dim expectedValue As String
                expectedValue = CStr(expFields(fieldKey))

                Dim actualValue As String
                If actFields.Exists(fieldName) Then
                    actualValue = CStr(actFields(fieldName))
                Else
                    actualValue = vbNullString
                End If

                Dim passed As Boolean
                passed = BT_CompareAssertValue(expectedValue, actualValue, fieldName)
                BT_AddAssertRecord records, count, batchId, subIndex, assertType, locator, _
                                fieldName, expectedValue, actualValue, passed, vbNullString
            Next fieldKey
        Next key

        ' 反向检查实际结果：只遍历预期键会漏掉“系统多输出了一行”的错误。
        ' 对汇总表和明细表而言，预期表代表完整结果，因此任何额外定位键都必须失败。
        For Each key In actMap.Keys
            locator = CStr(key)
            If Not expMap.Exists(locator) Then
                BT_AddAssertRecord records, count, batchId, subIndex, assertType, locator, _
                                "记录存在性", "(不应存在)", "(实际存在)", False, _
                                "实际结果存在预期表未定义的额外记录"
            End If
        Next key
    End Sub

    Private Function BT_CompareAssertValue( _
        ByVal expectedValue As String, _
        ByVal actualValue As String, _
        ByVal fieldName As String) As Boolean

        If InStr(1, fieldName, "原因", vbTextCompare) > 0 And Right$(expectedValue, 1) = "*" Then
            Dim prefix As String
            prefix = Left$(expectedValue, Len(expectedValue) - 1)
            BT_CompareAssertValue = (InStr(1, actualValue, prefix, vbTextCompare) > 0)
            Exit Function
        End If

        If InStr(1, fieldName, "原因", vbTextCompare) > 0 Then
            Dim expectedCodes As String
            Dim actualCodes As String
            expectedCodes = BT_ExtractErrorCodes(expectedValue)
            actualCodes = BT_ExtractErrorCodes(actualValue)
            If expectedCodes <> vbNullString Then
                ' 比较完整错误码序列，避免“预期 E01，实际 E01+E04”被误判为通过。
                BT_CompareAssertValue = (expectedCodes = actualCodes)
                Exit Function
            End If
        End If

        If StrComp(Trim$(fieldName), "效期", vbTextCompare) = 0 Then
            BT_CompareAssertValue = (BT_NormalizeAssertDateText(expectedValue) = BT_NormalizeAssertDateText(actualValue))
            Exit Function
        End If

        BT_CompareAssertValue = (Trim$(expectedValue) = Trim$(actualValue))
    End Function

    ' 仅用于批量断言比较：把 Excel 日期显示差异统一成 yyyy/mm/dd。
    ' 注意：这里不参与输入校验，不会放宽 M04/M05 对原始效期格式的严格要求。
    Private Function BT_NormalizeAssertDateText(ByVal valueText As String) As String
        Dim textValue As String
        textValue = Trim$(valueText)
        If textValue = vbNullString Then Exit Function

        On Error GoTo Fallback
        If IsDate(textValue) Then
            BT_NormalizeAssertDateText = Format$(CDate(textValue), "yyyy/mm/dd")
            Exit Function
        End If

    Fallback:
        BT_NormalizeAssertDateText = textValue
    End Function

    ' 从原因文本提取全部错误码，并保持出现顺序，例如：
    ' "E01 - ...; E04 - ..." → "E01;E04"。
    Private Function BT_ExtractErrorCodes(ByVal valueText As String) As String
        Dim textValue As String
        textValue = UCase$(Trim$(valueText))
        If Len(textValue) < 3 Then Exit Function

        Dim result As String
        Dim i As Long
        For i = 1 To Len(textValue) - 2
            Dim candidate As String
            candidate = Mid$(textValue, i, 3)

            If Left$(candidate, 1) = "E" _
                And IsNumeric(Mid$(candidate, 2, 2)) Then
                If InStr(1, ";" & result & ";", ";" & candidate & ";", vbTextCompare) = 0 Then
                    If result <> vbNullString Then result = result & ";"
                    result = result & candidate
                End If
            End If
        Next i

        BT_ExtractErrorCodes = result
    End Function

    Private Function BT_BuildExpectedRowMap( _
        ByVal ws As Worksheet, _
        ByVal batchId As String, _
        ByRef shipments() As String, _
        ByRef keyHeaders As Variant, _
        ByRef valueHeaders As Variant) As Object

        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare

        Dim colInfo As Object
        Set colInfo = BT_ParseSheetColumns(ws, keyHeaders, valueHeaders, batchId)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, colInfo("FirstDataCol"))
        If lastRow < 2 Then
            Set BT_BuildExpectedRowMap = result
            Exit Function
        End If

        ' 推荐计划使用 BATCH-* 描述性 ID，而历史预期表使用“批次1/批次9”等标签。
        ' 只有预期表中真实存在同名批次时才按批次列过滤，否则按物流单号匹配。
        Dim filterByBatch As Boolean
        If colInfo("HasBatchCol") Then
            filterByBatch = BT_BatchValueExists(ws, colInfo("BatchCol"), batchId)
        End If

        Dim r As Long
        For r = 2 To lastRow
            If filterByBatch Then
                If Trim$(CStr(ws.Cells(r, colInfo("BatchCol")).Value)) <> batchId Then GoTo NextExpRow
            End If

            Dim shipNo As String
            shipNo = Trim$(CStr(ws.Cells(r, colInfo("ShipCol")).Value))
            If Len(shipNo) > 0 Then
                If Not BT_ShipmentInList(shipNo, shipments) Then GoTo NextExpRow
            End If

            Dim locator As String
            locator = BT_BuildLocator(ws, r, colInfo, keyHeaders)

            Dim fields As Object
            If result.Exists(locator) Then
                Set fields = result(locator)
            Else
                Set fields = CreateObject("Scripting.Dictionary")
                fields.CompareMode = vbTextCompare
                result.Add locator, fields
            End If

            Dim i As Long
            For i = LBound(valueHeaders) To UBound(valueHeaders)
                Dim headerName As String
                headerName = CStr(valueHeaders(i))
                fields(headerName) = Trim$(CStr(ws.Cells(r, colInfo("ValueCol_" & headerName)).Value))
            Next i
    NextExpRow:
        Next r

        Set BT_BuildExpectedRowMap = result
    End Function

    Private Function BT_BatchValueExists( _
        ByVal ws As Worksheet, _
        ByVal batchCol As Long, _
        ByVal batchId As String) As Boolean

        If batchCol <= 0 Or Len(Trim$(batchId)) = 0 Then Exit Function

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, batchCol)

        Dim r As Long
        For r = 2 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(r, batchCol).Value)), batchId, vbTextCompare) = 0 Then
                BT_BatchValueExists = True
                Exit Function
            End If
        Next r
    End Function

    Private Function BT_BuildActualRowMap( _
        ByVal ws As Worksheet, _
        ByRef keyHeaders As Variant, _
        ByRef valueHeaders As Variant) As Object

        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare

        Dim colInfo As Object
        Set colInfo = BT_ParseSheetColumns(ws, keyHeaders, valueHeaders, vbNullString)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, colInfo("FirstDataCol"))
        If lastRow < 2 Then
            Set BT_BuildActualRowMap = result
            Exit Function
        End If

        Dim r As Long
        For r = 2 To lastRow
            Dim locator As String
            locator = BT_BuildLocator(ws, r, colInfo, keyHeaders)
            If locator = vbNullString Then GoTo NextActRow

            Dim fields As Object
            Set fields = CreateObject("Scripting.Dictionary")
            fields.CompareMode = vbTextCompare

            Dim i As Long
            For i = LBound(valueHeaders) To UBound(valueHeaders)
                Dim headerName As String
                headerName = CStr(valueHeaders(i))
                fields(headerName) = Trim$(CStr(ws.Cells(r, colInfo("ValueCol_" & headerName)).Value))
            Next i

            result.Add locator, fields
    NextActRow:
        Next r

        Set BT_BuildActualRowMap = result
    End Function

    Private Function BT_ParseSheetColumns( _
        ByVal ws As Worksheet, _
        ByRef keyHeaders As Variant, _
        ByRef valueHeaders As Variant, _
        ByVal batchId As String) As Object

        Dim info As Object
        Set info = CreateObject("Scripting.Dictionary")

        info("HasBatchCol") = False
        If Trim$(CStr(ws.Cells(1, 1).Value)) = HEADER_BATCH_COL Then
            info("HasBatchCol") = True
            info("BatchCol") = 1
        End If

        Dim i As Long
        For i = LBound(keyHeaders) To UBound(keyHeaders)
            Dim keyHeader As String
            keyHeader = CStr(keyHeaders(i))
            info("KeyCol_" & keyHeader) = BT_FindHeaderColumn(ws, keyHeader, True)
            If i = LBound(keyHeaders) Then
                info("ShipCol") = info("KeyCol_" & keyHeader)
                info("FirstDataCol") = info("KeyCol_" & keyHeader)
            End If
        Next i

        For i = LBound(valueHeaders) To UBound(valueHeaders)
            Dim valueHeader As String
            valueHeader = CStr(valueHeaders(i))
            info("ValueCol_" & valueHeader) = BT_FindHeaderColumn(ws, valueHeader, True)
        Next i

        Set BT_ParseSheetColumns = info
    End Function

    Private Function BT_BuildLocator( _
        ByVal ws As Worksheet, _
        ByVal rowIndex As Long, _
        ByVal colInfo As Object, _
        ByRef keyHeaders As Variant) As String

        Dim parts() As String
        ReDim parts(LBound(keyHeaders) To UBound(keyHeaders))

        Dim i As Long
        For i = LBound(keyHeaders) To UBound(keyHeaders)
            Dim headerName As String
            headerName = CStr(keyHeaders(i))
            parts(i) = BT_NormalizeLocatorPart(headerName, ws.Cells(rowIndex, colInfo("KeyCol_" & headerName)).Value)
        Next i

        BT_BuildLocator = Join(parts, "|")
    End Function

    ' 仅用于批量断言定位键：统一 WMS 占位值和效期显示格式，
    ' 避免预期表与输出表只有 Excel 显示类型不同却被误判为不同明细。
    Private Function BT_NormalizeLocatorPart(ByVal headerName As String, ByVal cellValue As Variant) As String
        Dim textValue As String

        If StrComp(Trim$(headerName), "WMS退单号", vbTextCompare) = 0 Then
            If IsError(cellValue) Then
                BT_NormalizeLocatorPart = "[N/A]"
                Exit Function
            End If

            textValue = Trim$(CStr(cellValue))
            If textValue = vbNullString _
            Or StrComp(textValue, "[N/A]", vbTextCompare) = 0 _
            Or StrComp(textValue, "NA", vbTextCompare) = 0 Then
                BT_NormalizeLocatorPart = "[N/A]"
                Exit Function
            End If
        Else
            If IsError(cellValue) Then
                BT_NormalizeLocatorPart = vbNullString
                Exit Function
            End If

            textValue = Trim$(CStr(cellValue))
        End If

        If StrComp(Trim$(headerName), "效期", vbTextCompare) = 0 Then
            textValue = BT_NormalizeAssertDateText(textValue)
        End If

        BT_NormalizeLocatorPart = textValue
    End Function

    Private Function BT_CollectExpectedCodes( _
        ByVal ws As Worksheet, _
        ByVal batchId As String, _
        ByRef shipments() As String) As Object

        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare

        Dim codeCol As Long
        codeCol = BT_FindHeaderColumn(ws, "错误码", False)
        If codeCol <= 0 Then codeCol = 8

        Dim shipCol As Long
        shipCol = BT_FindHeaderColumn(ws, "物流单号", False)
        If shipCol <= 0 Then shipCol = IIf(Trim$(CStr(ws.Cells(1, 1).Value)) = HEADER_BATCH_COL, 4, 3)

        Dim batchCol As Long
        batchCol = 0
        If Trim$(CStr(ws.Cells(1, 1).Value)) = HEADER_BATCH_COL Then batchCol = 1
        Dim filterByBatch As Boolean
        If batchCol > 0 Then filterByBatch = BT_BatchValueExists(ws, batchCol, batchId)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, codeCol)
        Dim r As Long
        For r = 2 To lastRow
            If filterByBatch Then
                If Trim$(CStr(ws.Cells(r, batchCol).Value)) <> batchId Then GoTo NextCodeRow
            End If

            If shipCol > 0 Then
                Dim shipNo As String
                shipNo = Trim$(CStr(ws.Cells(r, shipCol).Value))
                If Len(shipNo) > 0 Then
                    If Not BT_ShipmentInList(shipNo, shipments) Then GoTo NextCodeRow
                End If
            End If

            Dim codeValue As String
            codeValue = Trim$(CStr(ws.Cells(r, codeCol).Value))
            If Len(codeValue) > 0 Then
                Dim codeKey As String
                codeKey = shipNo & "|" & codeValue
                If Not result.Exists(codeKey) Then result.Add codeKey, True
            End If
    NextCodeRow:
        Next r

        Set BT_CollectExpectedCodes = result
    End Function

    Private Function BT_CollectActualCodes(ByVal ws As Worksheet) As Object
        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare

        Dim codeCol As Long
        codeCol = BT_FindHeaderColumn(ws, "错误码", False)
        If codeCol <= 0 Then codeCol = 8

        Dim shipCol As Long
        shipCol = BT_FindHeaderColumn(ws, "物流单号", False)
        If shipCol <= 0 Then shipCol = 3

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, codeCol)
        Dim r As Long
        For r = 2 To lastRow
            Dim codeValue As String
            codeValue = Trim$(CStr(ws.Cells(r, codeCol).Value))
            If Len(codeValue) > 0 Then
                Dim shipNo As String
                shipNo = Trim$(CStr(ws.Cells(r, shipCol).Value))
                Dim codeKey As String
                codeKey = shipNo & "|" & codeValue
                If Not result.Exists(codeKey) Then result.Add codeKey, True
            End If
        Next r

        Set BT_CollectActualCodes = result
    End Function

    Private Function BT_ResolveStructuredAssert( _
        ByVal tempWb As Workbook, _
        ByVal runMode As String, _
        ByVal itemName As String) As String

        Select Case LCase$(Trim$(itemName))
            Case "运行历史.运行类型", "运行类型"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 3, runMode)
            Case "运行历史.退单表行数", "退单表行数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 4, runMode)
            Case "运行历史.库存表行数", "库存表行数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 5, runMode)
            Case "运行历史.校验失败数", "校验失败数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 10, runMode)
            Case "运行历史.分配成功数", "分配成功数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 11, runMode)
            Case "运行历史.分配失败数", "分配失败数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 12, runMode)
            Case "运行历史.总回溯次数", "总回溯次数"
                BT_ResolveStructuredAssert = BT_ReadHistoryField(tempWb, 14, runMode)
            Case "汇总表行数"
                BT_ResolveStructuredAssert = CStr(BT_SheetDataRowCount(tempWb, SHEET_SUMMARY_ACTUAL))
            Case "明细表行数"
                BT_ResolveStructuredAssert = CStr(BT_SheetDataRowCount(tempWb, SHEET_DETAIL_ACTUAL))
            Case "异常表行数"
                BT_ResolveStructuredAssert = CStr(BT_SheetDataRowCount(tempWb, SHEET_ANOMALY_ACTUAL))
            Case "历史行数"
                BT_ResolveStructuredAssert = CStr(BT_HistoryDataRowCount(tempWb))
            Case "汇总错误码"
                BT_ResolveStructuredAssert = BT_CollectSummaryErrorCodes(tempWb)
            Case Else
                BT_ResolveStructuredAssert = vbNullString
        End Select
    End Function

    ' 统计实际输出表的数据行数（第 1 行表头；空表或无表记 0）。
    Private Function BT_SheetDataRowCount( _
        ByVal wb As Workbook, _
        ByVal sheetName As String) As Long

        Dim ws As Worksheet
        If Not BT_TryGetWorksheet(wb, sheetName, ws) Then Exit Function

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, 1)
        If lastRow <= 1 Then Exit Function
        BT_SheetDataRowCount = lastRow - 1
    End Function

    ' 运行历史表的数据行数（追加式日志，行数本身是 R127/R128 的验证点）。
    Private Function BT_HistoryDataRowCount(ByVal wb As Workbook) As Long
        Dim ws As Worksheet
        If Not BT_TryGetRunHistorySheet(wb, ws) Then Exit Function

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, 1)
        If lastRow <= 1 Then Exit Function
        BT_HistoryDataRowCount = lastRow - 1
    End Function

    ' 汇总表“原因”列中出现的全部错误码，按首次出现顺序去重，分号连接（如 E01;E04）。
    Private Function BT_CollectSummaryErrorCodes(ByVal wb As Workbook) As String
        Dim ws As Worksheet
        If Not BT_TryGetWorksheet(wb, SHEET_SUMMARY_ACTUAL, ws) Then Exit Function

        Dim reasonCol As Long
        reasonCol = BT_FindHeaderColumn(ws, "原因", False)
        If reasonCol <= 0 Then reasonCol = 4

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, reasonCol)

        Dim result As String
        Dim r As Long
        For r = 2 To lastRow
            Dim codes As String
            codes = BT_ExtractErrorCodes(CStr(ws.Cells(r, reasonCol).Value))
            Dim parts() As String
            parts = Split(codes, ";")
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                If Len(parts(i)) > 0 Then
                    If InStr(1, ";" & result & ";", ";" & parts(i) & ";", vbTextCompare) = 0 Then
                        If result <> vbNullString Then result = result & ";"
                        result = result & parts(i)
                    End If
                End If
            Next i
        Next r

        BT_CollectSummaryErrorCodes = result
    End Function

    Private Function BT_ReadHistoryField( _
        ByVal tempWb As Workbook, _
        ByVal colIndex As Long, _
        ByVal runMode As String) As String

        Dim wsHist As Worksheet
        If Not BT_TryGetRunHistorySheet(tempWb, wsHist) Then Exit Function

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(wsHist, 1)
        If lastRow < 2 Then Exit Function

        BT_ReadHistoryField = Trim$(CStr(wsHist.Cells(lastRow, colIndex).Value))
    End Function


    ' =============================================================================
    ' 八、通用工具
    ' =============================================================================

    Private Sub BT_AddAssertRecord( _
        ByRef records() As BatchAssertRecord, _
        ByRef count As Long, _
        ByVal batchId As String, _
        ByVal subIndex As Long, _
        ByVal assertType As String, _
        ByVal locator As String, _
        ByVal fieldName As String, _
        ByVal expectedValue As String, _
        ByVal actualValue As String, _
        ByVal passed As Boolean, _
        ByVal remark As String)

        count = count + 1
        ReDim Preserve records(1 To count)
        With records(count)
            .BatchId = batchId
            .SubIndex = subIndex
            .AssertType = assertType
            .Locator = locator
            .FieldName = fieldName
            .ExpectedValue = expectedValue
            .ActualValue = actualValue
            .Passed = passed
            .Remark = remark
        End With
    End Sub

    Private Sub BT_CountAsserts(ByRef records() As BatchAssertRecord, ByRef passCount As Long, ByRef failCount As Long)
        Dim n As Long
        n = BT_AssertRecordCount(records)
        If n <= 0 Then Exit Sub

        Dim i As Long
        For i = 1 To n
            If records(i).Passed Then
                passCount = passCount + 1
            Else
                failCount = failCount + 1
            End If
        Next i
    End Sub

    Private Function BT_AssertRecordCount(ByRef records() As BatchAssertRecord) As Long
        On Error GoTo EmptyArr
        BT_AssertRecordCount = UBound(records) - LBound(records) + 1
        Exit Function
    EmptyArr:
        BT_AssertRecordCount = 0
    End Function

    Private Function BT_ConfigFingerprint(ByRef cfg As ConfigStruct) As String
        BT_ConfigFingerprint = CStr(cfg.MaxBacktrackCount) & "|" & _
                            cfg.DebugLogLevel & "|" & _
                            CStr(cfg.LotCaseSensitive) & "|" & _
                            cfg.NoExpirySentinel & "|" & _
                            CStr(cfg.DetailedLogLimit)
    End Function

    Private Function BT_ParseShipmentList(ByVal text As String, Optional ByVal delimiter As String = ";") As String()
        Dim raw As String
        raw = Replace(text, "；", delimiter)
        raw = Replace(raw, ",", delimiter)
        raw = Replace(raw, "，", delimiter)
        raw = Replace(raw, vbCrLf, delimiter)
        raw = Replace(raw, vbLf, delimiter)

        Dim parts() As String
        parts = Split(raw, delimiter)

        Dim result() As String
        Dim count As Long
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim token As String
            token = Trim$(parts(i))
            If token <> vbNullString Then
                count = count + 1
                ReDim Preserve result(1 To count)
                result(count) = token
            End If
        Next i

        If count > 0 Then BT_ParseShipmentList = result
    End Function

    Private Function BT_CollectShipmentsFromReturnSheet(ByVal ws As Worksheet) As String()
        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, 1)
        If lastRow < 2 Then Exit Function

        Dim seen As Object
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = vbTextCompare

        Dim r As Long
        For r = 2 To lastRow
            Dim shipNo As String
            shipNo = Trim$(CStr(ws.Cells(r, 1).Value))
            If shipNo <> vbNullString Then
                If Not seen.Exists(shipNo) Then seen.Add shipNo, True
            End If
        Next r

        If seen.Count = 0 Then Exit Function

        Dim result() As String
        ReDim result(1 To seen.Count)
        Dim idx As Long
        Dim key As Variant
        For Each key In seen.Keys
            idx = idx + 1
            result(idx) = CStr(key)
        Next key

        BT_CollectShipmentsFromReturnSheet = result
    End Function

    Private Function BT_FindConfigRow(ByVal ws As Worksheet, ByVal shipmentNo As String) As Long
        Dim shipCol As Long
        shipCol = BT_FindHeaderColumn(ws, "物流单号", True)

        Dim lastRow As Long
        lastRow = BT_GetLastUsedRow(ws, shipCol)

        Dim r As Long
        For r = 2 To lastRow
            If Trim$(CStr(ws.Cells(r, shipCol).Value)) = Trim$(shipmentNo) Then
                BT_FindConfigRow = r
                Exit Function
            End If
        Next r

        Err.Raise vbObjectError + 1601, "BT_FindConfigRow", _
              "输入_配置 中找不到物流单号 [" & shipmentNo & "]。"
    End Function

    Private Function BT_JoinShipments(ByRef shipments() As String) As String
        If BT_StringArrayCount(shipments) = 0 Then Exit Function
        BT_JoinShipments = Join(shipments, ";")
    End Function

    Private Function BT_ToShipmentSet(ByRef shipments() As String) As Object
        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare

        Dim i As Long
        For i = LBound(shipments) To UBound(shipments)
            Dim shipNo As String
            shipNo = Trim$(shipments(i))
            If shipNo <> vbNullString Then
                If Not result.Exists(shipNo) Then result.Add shipNo, True
            End If
        Next i

        Set BT_ToShipmentSet = result
    End Function

    Private Function BT_ShipmentInList(ByVal shipNo As String, ByRef shipments() As String) As Boolean
        Dim i As Long
        For i = LBound(shipments) To UBound(shipments)
            If StrComp(Trim$(shipments(i)), Trim$(shipNo), vbTextCompare) = 0 Then
                BT_ShipmentInList = True
                Exit Function
            End If
        Next i
    End Function

    Private Function BT_IsEnabled(ByVal cellValue As Variant) As Boolean
        Dim textValue As String
        textValue = Trim$(CStr(cellValue))
        BT_IsEnabled = (textValue = "是") Or (UCase$(textValue) = "Y") Or (UCase$(textValue) = "YES") Or (textValue = "1") Or (UCase$(textValue) = "TRUE")
    End Function

    Private Function BT_IsValidDebugLevel(ByVal levelText As String) As Boolean
        Select Case Trim$(levelText)
            Case DEBUG_LEVEL_OFF, DEBUG_LEVEL_SIMPLE, DEBUG_LEVEL_DETAIL, "关闭", "简版", "详细"
                BT_IsValidDebugLevel = True
            Case Else
                BT_IsValidDebugLevel = False
        End Select
    End Function

    Private Function BT_CodePresentText(ByVal actCodes As Object, ByVal codeKey As String) As String
        If actCodes.Exists(codeKey) Then
            BT_CodePresentText = codeKey
        Else
            BT_CodePresentText = "(未出现)"
        End If
    End Function

    Private Sub BT_CopyHeaderRow(ByVal sourceWs As Worksheet, ByVal targetWs As Worksheet)
        BT_CopyUsedRowCells sourceWs, 1, targetWs, 1
    End Sub

    ' 按单元格复制一行已使用区域的值和类型。
    ' 文本值必须先把目标格设为文本格式再写入，否则 Excel 会把 "00001"
    ' 自动转换为数值 1，导致临时工作簿与源测试数据语义不同。
    ' 这样也不会受 AutoFilter 隐藏行影响或改变用户当前筛选状态。
    Private Sub BT_CopyUsedRowCells( _
        ByVal sourceWs As Worksheet, _
        ByVal sourceRow As Long, _
        ByVal targetWs As Worksheet, _
        ByVal targetRow As Long)

        Dim lastCol As Long
        lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
        If lastCol <= 0 Then Exit Sub

        Dim c As Long
        For c = 1 To lastCol
            Dim sourceValue As Variant
            sourceValue = sourceWs.Cells(sourceRow, c).Value2

            With targetWs.Cells(targetRow, c)
                If VarType(sourceValue) = vbString Then
                    .NumberFormat = "@"
                    .Value2 = CStr(sourceValue)
                Else
                    .NumberFormat = sourceWs.Cells(sourceRow, c).NumberFormat
                    .Value2 = sourceValue
                End If
            End With
        Next c
    End Sub

    Private Sub BT_ClearResultData(ByVal ws As Worksheet)
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow > 1 Then ws.Rows("2:" & lastRow).ClearContents
    End Sub

    Private Sub BT_CloseTempWorkbook(ByVal wb As Workbook)
        On Error Resume Next

        If wb Is Nothing Then Exit Sub

        ' 关闭临时工作簿属于清理动作；即使 Excel 对象状态异常，也不应打断批量结果写入。
        Dim oldDisplayAlerts As Boolean
        oldDisplayAlerts = Application.DisplayAlerts

        Application.CutCopyMode = False
        Application.DisplayAlerts = False

        wb.Saved = True
        wb.Close SaveChanges:=False

        Application.DisplayAlerts = oldDisplayAlerts
        On Error GoTo 0
    End Sub

    Private Function BT_EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
        On Error Resume Next
        Set BT_EnsureSheet = wb.Worksheets(sheetName)
        On Error GoTo 0

        If BT_EnsureSheet Is Nothing Then
            wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
            wb.Worksheets(wb.Worksheets.Count).Name = sheetName
            Set BT_EnsureSheet = wb.Worksheets(sheetName)
        End If
    End Function

    Private Function BT_TryGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        BT_TryGetWorksheet = Not ws Is Nothing
        On Error GoTo 0
    End Function

    Private Function BT_TryGetExpectedAnomalySheet(ByVal wb As Workbook, ByRef ws As Worksheet) As Boolean
        If BT_TryGetWorksheet(wb, SHEET_ANOMALY_EXPECTED, ws) Then
            BT_TryGetExpectedAnomalySheet = True
            Exit Function
        End If
        BT_TryGetExpectedAnomalySheet = BT_TryGetWorksheet(wb, "预期_数据异常明细表", ws)
    End Function

    Private Function BT_TryGetRunHistorySheet(ByVal wb As Workbook, ByRef ws As Worksheet) As Boolean
        If BT_TryGetWorksheet(wb, SHEET_RUN_HISTORY, ws) Then
            BT_TryGetRunHistorySheet = True
            Exit Function
        End If
        ' 兼容尚未经过迁移脚本处理的旧测试工作簿。
        BT_TryGetRunHistorySheet = BT_TryGetWorksheet(wb, "运行历史记录", ws)
    End Function

    Private Function BT_SheetNeedsPlanHeader(ByVal ws As Worksheet) As Boolean
        BT_SheetNeedsPlanHeader = (Trim$(CStr(ws.Cells(1, 1).Value)) <> COL_BATCH_ID)
    End Function

    Private Function BT_SheetNeedsResultHeader(ByVal ws As Worksheet) As Boolean
        BT_SheetNeedsResultHeader = (Trim$(CStr(ws.Cells(1, 1).Value)) <> "批次ID")
    End Function

    Private Function BT_SheetNeedsAssertHeader(ByVal ws As Worksheet) As Boolean
        BT_SheetNeedsAssertHeader = (Trim$(CStr(ws.Cells(1, 1).Value)) <> "批次ID")
    End Function

    Private Function BT_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String, ByVal required As Boolean) As Long
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        Dim c As Long
        For c = 1 To lastCol
            If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
                BT_FindHeaderColumn = c
                Exit Function
            End If
        Next c

        If required Then
            Err.Raise vbObjectError + 1602, "BT_FindHeaderColumn", _
                    "工作表 [" & ws.Name & "] 缺少表头 [" & headerText & "]。"
        End If
    End Function

    Private Function BT_GetLastUsedRow(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
        Dim lastCell As Range
        Set lastCell = ws.Cells.Find(What:="*", _
                                    After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

        If lastCell Is Nothing Then
            BT_GetLastUsedRow = 1
        Else
            BT_GetLastUsedRow = lastCell.Row
        End If
    End Function

    Private Function BT_PlanCount(ByRef plans() As BatchPlanEntry) As Long
        On Error GoTo EmptyArr
        BT_PlanCount = UBound(plans) - LBound(plans) + 1
        Exit Function
    EmptyArr:
        BT_PlanCount = 0
    End Function

    Private Function BT_SubRunCount(ByRef subRuns() As BatchSubRunSpec) As Long
        On Error GoTo EmptyArr
        BT_SubRunCount = UBound(subRuns) - LBound(subRuns) + 1
        Exit Function
    EmptyArr:
        BT_SubRunCount = 0
    End Function

    Private Function BT_StringArrayCount(ByRef arr() As String) As Long
        On Error GoTo EmptyArr
        BT_StringArrayCount = UBound(arr) - LBound(arr) + 1
        Exit Function
    EmptyArr:
        BT_StringArrayCount = 0
    End Function

    Private Function BT_SafeCLng(ByVal value As Variant) As Long
        On Error Resume Next
        BT_SafeCLng = CLng(value)
        On Error GoTo 0
    End Function
