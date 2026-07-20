Option Explicit

' =============================================================================
' M02_配置管理（modConfig）
' =============================================================================
' 职责：
' 1. 从“输入_配置”工作表读取配置。
' 2. 对缺失的可选配置使用默认值。
' 3. 对非法配置尽早报错，避免后续分配算法在错误参数下运行。
'
' 当前支持两种配置表结构：
' 1. 生产键值结构：参数名 | 值 | 说明
' 2. 批量测试结构：物流单号 | TC编号 | 最大回溯次数 | 调试日志级别 | ...
'
' 说明：
' - LoadConfig：自动识别生产键值结构或批量测试结构。
' - LoadConfigForShipment：按物流单号读取指定行，适用于“测试用例部分汇总.xlsm”这类多用例汇总表。
' - “详细日志单表上限”如缺失，默认使用 DEFAULT_DETAILED_LOG_LIMIT。
' =============================================================================

Private Const ERR_BASE As Long = vbObjectError + 2000

Private Const HEADER_SHIPMENT_NO As String = "物流单号"
Private Const HEADER_PARAM_NAME As String = "参数名"
Private Const HEADER_PARAM_VALUE As String = "值"
Private Const HEADER_MAX_BACKTRACK As String = "最大回溯次数"
Private Const HEADER_DEBUG_LEVEL As String = "调试日志级别"
Private Const HEADER_LOT_MODE As String = "批号比较模式"
Private Const HEADER_NO_EXPIRY_SENTINEL As String = "无保质期哨兵值"
Private Const HEADER_DETAILED_LOG_LIMIT As String = "详细日志单表上限"

' -----------------------------------------------------------------------------
' 公开函数
' -----------------------------------------------------------------------------

Public Function LoadConfig(ByVal ws As Worksheet) As ConfigStruct
    ' 生产工作簿推荐用“参数名/值/说明”；测试汇总工作簿继续用横向多行结构。
    ValidateWorksheet ws
    If IsKeyValueConfigSheet(ws) Then
        LoadConfig = BuildConfigFromKeyValue(ws)
    Else
        LoadConfig = BuildConfigFromRow(ws, 2)
    End If
End Function

Public Function LoadConfigForShipment(ByVal ws As Worksheet, ByVal shipmentNo As String) As ConfigStruct
    ' 按物流单号读取配置。主要给后续 M16 测试入口使用。
    ValidateWorksheet ws

    If IsKeyValueConfigSheet(ws) Then
        LoadConfigForShipment = BuildConfigFromKeyValue(ws)
        Exit Function
    End If

    Dim targetShipmentNo As String
    targetShipmentNo = Trim$(shipmentNo)
    If targetShipmentNo = vbNullString Then
        RaiseConfigError "读取配置失败：物流单号不能为空。"
    End If

    Dim shipmentCol As Long
    shipmentCol = FindHeaderColumn(ws, HEADER_SHIPMENT_NO, True)

    Dim lastRow As Long
    lastRow = GetLastUsedRow(ws)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If Trim$(CStr(ws.Cells(rowIndex, shipmentCol).Value)) = targetShipmentNo Then
            LoadConfigForShipment = BuildConfigFromRow(ws, rowIndex)
            Exit Function
        End If
    Next rowIndex

    RaiseConfigError "读取配置失败：输入_配置 中找不到物流单号 [" & targetShipmentNo & "]。"
End Function

Public Function BuildDefaultConfig() As ConfigStruct
    ' 默认配置集中在这里，后续测试也可以直接调用。
    With BuildDefaultConfig
        .MaxBacktrackCount = DEFAULT_MAX_BACKTRACK_COUNT
        .DebugLogLevel = DEFAULT_DEBUG_LOG_LEVEL
        .DetailedLogLimit = DEFAULT_DETAILED_LOG_LIMIT
        .LotCaseSensitive = False
        .NoExpirySentinel = DEFAULT_NO_EXPIRY_SENTINEL
    End With
End Function

Private Function BuildConfigFromKeyValue(ByVal ws As Worksheet) As ConfigStruct
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    Dim nameCol As Long
    Dim valueCol As Long
    nameCol = FindHeaderColumn(ws, HEADER_PARAM_NAME, True)
    valueCol = FindHeaderColumn(ws, HEADER_PARAM_VALUE, True)

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = GetLastUsedRow(ws)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim paramName As String
        paramName = Trim$(CStr(ws.Cells(rowIndex, nameCol).Value))
        If Len(paramName) > 0 Then
            values(paramName) = ws.Cells(rowIndex, valueCol).Value
        End If
    Next rowIndex

    cfg.MaxBacktrackCount = ParsePositiveLong( _
        GetConfigValue(values, HEADER_MAX_BACKTRACK), _
        DEFAULT_MAX_BACKTRACK_COUNT, _
        HEADER_MAX_BACKTRACK)

    cfg.DebugLogLevel = ParseDebugLogLevel( _
        GetConfigValue(values, HEADER_DEBUG_LEVEL), _
        DEFAULT_DEBUG_LOG_LEVEL)

    cfg.LotCaseSensitive = ParseLotCaseSensitive( _
        GetConfigValue(values, HEADER_LOT_MODE), _
        DEFAULT_LOT_MODE)

    cfg.NoExpirySentinel = ParseRequiredText( _
        GetConfigValue(values, HEADER_NO_EXPIRY_SENTINEL), _
        DEFAULT_NO_EXPIRY_SENTINEL, _
        HEADER_NO_EXPIRY_SENTINEL)

    cfg.DetailedLogLimit = ParsePositiveLong( _
        GetConfigValue(values, HEADER_DETAILED_LOG_LIMIT), _
        DEFAULT_DETAILED_LOG_LIMIT, _
        HEADER_DETAILED_LOG_LIMIT)

    BuildConfigFromKeyValue = cfg
End Function

' -----------------------------------------------------------------------------
' 配置行解析
' -----------------------------------------------------------------------------

Private Function BuildConfigFromRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As ConfigStruct
    Dim cfg As ConfigStruct
    cfg = BuildDefaultConfig()

    If rowIndex < 2 Or rowIndex > GetLastUsedRow(ws) Then
        RaiseConfigError "读取配置失败：配置行号无效 [" & CStr(rowIndex) & "]。"
    End If

    Dim maxBacktrackCol As Long
    Dim debugLevelCol As Long
    Dim lotModeCol As Long
    Dim sentinelCol As Long
    Dim detailedLimitCol As Long

    maxBacktrackCol = FindHeaderColumn(ws, HEADER_MAX_BACKTRACK, True)
    debugLevelCol = FindHeaderColumn(ws, HEADER_DEBUG_LEVEL, True)
    lotModeCol = FindHeaderColumn(ws, HEADER_LOT_MODE, True)
    sentinelCol = FindHeaderColumn(ws, HEADER_NO_EXPIRY_SENTINEL, True)
    detailedLimitCol = FindHeaderColumn(ws, HEADER_DETAILED_LOG_LIMIT, False)

    cfg.MaxBacktrackCount = ParsePositiveLong( _
        ws.Cells(rowIndex, maxBacktrackCol).Value, _
        DEFAULT_MAX_BACKTRACK_COUNT, _
        HEADER_MAX_BACKTRACK)

    cfg.DebugLogLevel = ParseDebugLogLevel( _
        ws.Cells(rowIndex, debugLevelCol).Value, _
        DEFAULT_DEBUG_LOG_LEVEL)

    cfg.LotCaseSensitive = ParseLotCaseSensitive( _
        ws.Cells(rowIndex, lotModeCol).Value, _
        DEFAULT_LOT_MODE)

    cfg.NoExpirySentinel = ParseRequiredText( _
        ws.Cells(rowIndex, sentinelCol).Value, _
        DEFAULT_NO_EXPIRY_SENTINEL, _
        HEADER_NO_EXPIRY_SENTINEL)

    If detailedLimitCol > 0 Then
        cfg.DetailedLogLimit = ParsePositiveLong( _
            ws.Cells(rowIndex, detailedLimitCol).Value, _
            DEFAULT_DETAILED_LOG_LIMIT, _
            HEADER_DETAILED_LOG_LIMIT)
    End If

    BuildConfigFromRow = cfg
End Function

Private Function GetConfigValue(ByVal values As Object, ByVal fieldName As String) As Variant
    If values Is Nothing Then
        GetConfigValue = vbNullString
    ElseIf values.Exists(fieldName) Then
        GetConfigValue = values(fieldName)
    Else
        GetConfigValue = vbNullString
    End If
End Function

Private Function ParsePositiveLong(ByVal rawValue As Variant, ByVal defaultValue As Long, ByVal fieldName As String) As Long
    Dim textValue As String
    textValue = Trim$(CStr(rawValue))

    If textValue = vbNullString Then
        ParsePositiveLong = defaultValue
        Exit Function
    End If

    If Not IsNumeric(textValue) Then
        RaiseConfigError "配置项 [" & fieldName & "] 必须是正整数，当前值=[" & textValue & "]。"
    End If

    If CLng(CDbl(textValue)) <> CDbl(textValue) Or CLng(CDbl(textValue)) <= 0 Then
        RaiseConfigError "配置项 [" & fieldName & "] 必须是正整数，当前值=[" & textValue & "]。"
    End If

    ParsePositiveLong = CLng(CDbl(textValue))
End Function

Private Function ParseDebugLogLevel(ByVal rawValue As Variant, ByVal defaultValue As String) As String
    Dim textValue As String
    textValue = Trim$(CStr(rawValue))

    If textValue = vbNullString Then
        ParseDebugLogLevel = defaultValue
        Exit Function
    End If

    Select Case textValue
        Case DEBUG_LEVEL_OFF, DEBUG_LEVEL_SIMPLE, DEBUG_LEVEL_DETAIL
            ParseDebugLogLevel = textValue
        Case Else
            RaiseConfigError "配置项 [" & HEADER_DEBUG_LEVEL & "] 只能是 关闭/简版/详细，当前值=[" & textValue & "]。"
    End Select
End Function

Private Function ParseLotCaseSensitive(ByVal rawValue As Variant, ByVal defaultValue As String) As Boolean
    Dim textValue As String
    textValue = Trim$(CStr(rawValue))

    If textValue = vbNullString Then
        textValue = defaultValue
    End If

    Select Case textValue
        Case LOT_MODE_INSENSITIVE
            ParseLotCaseSensitive = False
        Case LOT_MODE_SENSITIVE
            ParseLotCaseSensitive = True
        Case Else
            RaiseConfigError "配置项 [" & HEADER_LOT_MODE & "] 只能是 不敏感/敏感，当前值=[" & textValue & "]。"
    End Select
End Function

Private Function ParseRequiredText(ByVal rawValue As Variant, ByVal defaultValue As String, ByVal fieldName As String) As String
    Dim textValue As String
    textValue = Trim$(CStr(rawValue))

    If textValue = vbNullString Then
        ParseRequiredText = defaultValue
    Else
        ParseRequiredText = textValue
    End If
End Function

' -----------------------------------------------------------------------------
' 表头和工作表工具
' -----------------------------------------------------------------------------

Private Sub ValidateWorksheet(ByVal ws As Worksheet)
    If ws Is Nothing Then
        RaiseConfigError "读取配置失败：工作表对象为空。"
    End If

    If GetLastUsedRow(ws) < 2 Then
        RaiseConfigError "读取配置失败：输入_配置 至少需要 1 行表头和 1 行配置数据。"
    End If
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String, ByVal required As Boolean) As Long
    Dim lastCol As Long
    lastCol = GetLastUsedColumn(ws)

    Dim colIndex As Long
    For colIndex = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, colIndex).Value)) = headerName Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex

    If required Then
        RaiseConfigError "读取配置失败：输入_配置 缺少表头 [" & headerName & "]。"
    End If

    FindHeaderColumn = 0
End Function

Private Function IsKeyValueConfigSheet(ByVal ws As Worksheet) As Boolean
    IsKeyValueConfigSheet = (FindHeaderColumn(ws, HEADER_PARAM_NAME, False) > 0 _
                             And FindHeaderColumn(ws, HEADER_PARAM_VALUE, False) > 0)
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    GetLastUsedRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    GetLastUsedColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub RaiseConfigError(ByVal message As String)
    Err.Raise ERR_BASE + 1, "modConfig", message
End Sub
