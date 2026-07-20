Option Explicit

' =============================================================================
' M01_基础数据模型（modTypes）
' =============================================================================
' 本模块只定义全系统共用的常量和 Type 结构体，不包含任何业务逻辑。
' 简单理解：这里是 VBA 版本的“数据字典”，后续模块都按这些结构传递数据。
'
' 设计原则：
' 1. M03 负责把 Excel 单元格原样读入 Raw* 结构。
' 2. M04 负责把 Raw* 标准化为 Normalized*，同时记录字段问题。
' 3. M05 以后再把字段问题转换成正式错误码 E01~E11。
' =============================================================================

' -----------------------------------------------------------------------------
' 1. 通用常量
' -----------------------------------------------------------------------------

' 效期单元格类型。M03 通过 VarType(cell.Value) 判断后写入 RawInventoryRow.ExpiryCellKind。
Public Const CELL_KIND_EXCEL_DATE As String = "ExcelDate"
Public Const CELL_KIND_TEXT       As String = "TextValue"
Public Const CELL_KIND_BLANK      As String = "Blank"
Public Const CELL_KIND_OTHER      As String = "Other"

' 来源表名称。用于 FieldNormalizeIssue / ValidationIssue 等问题定位。
Public Const SOURCE_RETURN_TABLE    As String = "退单表"
Public Const SOURCE_INVENTORY_TABLE As String = "质检库存表"

' 字段标准化问题类型。
Public Const ISSUE_KIND_EMPTY        As String = "Empty"
Public Const ISSUE_KIND_FORMAT_ERROR As String = "FormatError"
Public Const ISSUE_KIND_RANGE_ERROR  As String = "RangeError"

' 调试日志级别。生产配置和测试配置都统一使用三档。
Public Const DEBUG_LEVEL_OFF    As String = "关闭"
Public Const DEBUG_LEVEL_SIMPLE As String = "简版"
Public Const DEBUG_LEVEL_DETAIL As String = "详细"

' 批号比较模式。
Public Const LOT_MODE_INSENSITIVE As String = "不敏感"
Public Const LOT_MODE_SENSITIVE   As String = "敏感"

' QC 情况合法值。
Public Const QC_ZP As String = "ZP"
Public Const QC_QC As String = "QC"
Public Const QC_NG As String = "NG"

' 默认配置值。
Public Const DEFAULT_MAX_BACKTRACK_COUNT As Long = 200
Public Const DEFAULT_DEBUG_LOG_LEVEL     As String = DEBUG_LEVEL_OFF
Public Const DEFAULT_DETAILED_LOG_LIMIT  As Long = 100000
Public Const DEFAULT_LOT_MODE            As String = LOT_MODE_INSENSITIVE
Public Const DEFAULT_NO_EXPIRY_SENTINEL  As String = "2099/01/01"

' 分配前校验错误码（M05 产出，对应 §4.1）。
Public Const ERR_E01 As String = "E01"
Public Const ERR_E02 As String = "E02"
Public Const ERR_E03 As String = "E03"
Public Const ERR_E04 As String = "E04"
Public Const ERR_E05 As String = "E05"
Public Const ERR_E06 As String = "E06"
Public Const ERR_E07 As String = "E07"
Public Const ERR_E08 As String = "E08"
Public Const ERR_E11 As String = "E11"

' 分配阶段错误码（M09 产出）。
' E09：分配前预检测或回溯耗尽选项，确认无法分配。
' E10：回溯次数超过 MaxBacktrackCount 配置上限。
' E99：工程守卫发现库存守恒等式被破坏，立即停止运行。
Public Const ERR_E09 As String = "E09"
Public Const ERR_E10 As String = "E10"
Public Const ERR_E99 As String = "E99"

' 汇总表/异常明细占位符。
Public Const NA_PLACEHOLDER As String = "[N/A]"

' 行/退单号状态（M11 产出，供 M12/M13 写入输出表）。
Public Const STATUS_BATCH_IMPORT As String = "批量导入"
Public Const STATUS_MANUAL         As String = "手工操作"
Public Const STATUS_UNALLOCATED    As String = "无法分配"
Public Const LINE_STATUS_FAILED    As String = "分配失败"

' M09 短路后未实际分配的 SKU 组标记；M11 用于区分"直接失败"与"连带回滚"。
Public Const ERROR_CASCADE_ROLLBACK As String = "连带回滚"

' M11 汇总表单条记录（AggregateWMSStatus 返回 WMSStatusEntry 数组，即 WMSStatusMap）。
Public Type WMSStatusEntry
    ShipmentNo As String
    WMSOrderNo As String
    Status     As String
    Reason     As String
End Type

' -----------------------------------------------------------------------------
' 2. 原始层：M03 数据加载输出
' -----------------------------------------------------------------------------

' 退单表原始行。
' 字段使用 Variant 是为了保留 Excel 原始值，例如行号可能被 Excel 读成数值 1。
Public Type RawReturnRow
    ExcelRowNum As Long
    ShipmentNo  As Variant
    WMSOrderNo  As Variant
    SKU         As Variant
    LineNo      As Variant
    Qty         As Variant
End Type

' 质检库存表原始行。
' Expiry 保留原始单元格值；ExpiryCellKind 记录它在 Excel 内部到底是日期、文本、空值还是其他类型。
Public Type RawInventoryRow
    ExcelRowNum    As Long
    ShipmentNo     As Variant
    SKU            As Variant
    QC             As Variant
    LotNo          As Variant
    Expiry         As Variant
    ExpiryCellKind As String
    Qty            As Variant
End Type

' -----------------------------------------------------------------------------
' 3. 标准化层：M04 数据标准化输出
' -----------------------------------------------------------------------------

' 标准化后的退单行。
' Valid 字段只说明“这个字段本身是否合法”，真正的错误码由 M05 统一生成。
Public Type NormalizedReturnLine
    ExcelRowNum As Long
    ShipmentNo  As String
    WMSOrderNo  As String
    SKU         As String
    LineNo      As String
    Qty         As Long
    LineNoValid As Boolean
    QtyValid    As Boolean
    EmptyFields As String
End Type

' 标准化后的质检库存行。
Public Type NormalizedInventoryLine
    ExcelRowNum As Long
    ShipmentNo  As String
    SKU         As String
    QC          As String
    LotNo       As String
    Expiry      As String
    Qty         As Long
    QCValid     As Boolean
    ExpiryValid As Boolean
    QtyValid    As Boolean
    EmptyFields As String
End Type

' 单字段标准化问题记录。
' M04 发现字段空值、格式非法或范围非法时记录在这里，M05 再把它转换成 E01~E05 等正式错误码。
Public Type FieldNormalizeIssue
    ExcelRowNum As Long
    SourceTable As String
    FieldName   As String
    RawValue    As String
    IssueKind   As String
End Type

' -----------------------------------------------------------------------------
' 4. 领域层：库存、分配、日志、校验和运行统计
' -----------------------------------------------------------------------------

' 库存五元组键：物流单号 + SKU + QC + 批号 + 效期。
Public Type InventoryKey
    ShipmentNo As String
    SKU        As String
    QC         As String
    LotNo      As String
    Expiry     As String
End Type

' 库存账本中的单行数据（M06 GetFiveTupleRows 返回，M07/M08 用于构建候选列表）。
' OriginalQty：建账本时的原始数量，整个分配过程中不变，供 M10 守恒断言使用。
' CurrentQty ：当前可用数量，随 Deduct/Undo 操作变化。
Public Type InventoryRow
    ShipmentNo  As String
    SKU         As String
    QC          As String
    LotNo       As String
    Expiry      As String
    OriginalQty As Long
    CurrentQty  As Long
End Type

' 单条分配明细。
Public Type AllocationDetail
    ShipmentNo As String
    WMSOrderNo As String
    SKU        As String
    LineNo     As String
    OrderQty   As Long
    QC         As String
    LotNo      As String
    Expiry     As String
    AllocQty   As Long
    LineStatus As String
    StrategyUsed As String
End Type

' 单个 SKU 组的回溯统计。即使调试日志关闭，也要保留统计值用于运行历史。
Public Type GroupStats
    ShipmentNo     As String
    SKU            As String
    BacktrackCount As Long
    PreCheckHit    As String
End Type

' 调试日志事件（19 列输出，见 调试日志19列规格说明.md）。
' IsFinalResult=True 的行在「简版」模式下输出；「详细」模式输出全部事件。
Public Type AllocationEvent
    ShipmentNo         As String
    SKU                As String
    WMSOrderNo         As String
    LineNo             As String
    DemandD            As Long
    ProcessOrder       As String
    DynamicNextMinQty  As String
    CandidateQCCount   As String
    ExcludedQCList     As String
    StrategyUsed       As String
    UsedQC             As String
    QCBefore           As String
    QCAfter            As String
    LotExpiryComboCount As String
    IsBacktrackRetry   As String
    BacktrackNo        As Long
    LineStatus         As String
    ErrorCode          As String
    FailSubType        As String
    IsFinalResult      As Boolean
    IsRevoked          As Boolean
End Type

' 单个 SKU 组的分配结果。
' 说明：VBA 的 Type 不适合在这里直接嵌入动态数组；Details/Events 由 M09 以后用独立数组或集合配套承载。
Public Type GroupAllocResult
    ShipmentNo As String
    SKU        As String
    Success    As Boolean
    ErrorCode  As String
    Stats      As GroupStats
End Type

' 单个物流单号的分配结果。
' GroupResults 动态数组由 M09/M11 按需要另外维护。
Public Type ShipmentAllocResult
    ShipmentNo As String
End Type

' 校验问题。M05 产出，供 M11/M13 汇总与异常明细使用。
Public Type ValidationIssue
    ShipmentNo  As String
    WMSOrderNo  As String
    SKU         As String
    ErrorCode   As String
    SourceTable As String
    ExcelRowNum As Long
    FieldName   As String
    RawValue    As String
    Reason      As String
End Type

' M05 校验汇总结果。
' 说明：VBA 的 Type 不能内嵌动态数组，具体 ValidationIssue 列表通过 ValidatePre 的 ByRef 参数返回。
Public Type ValidationResult
    HasFailures         As Boolean
    FailedShipmentCount As Long
End Type

' 数据异常明细行（§5.4）。由 M05 BuildAnomalyRows 产出，供 M13 写 Excel。
Public Type AnomalyRow
    SourceTable As String
    ExcelRowNum As Long
    ShipmentNo  As String
    WMSOrderNo  As String
    SKU         As String
    FieldName   As String
    RawValue    As String
    ErrorCode   As String
    Reason      As String
End Type

' 配置结构体。
' M02 从生产配置表读取一个全局配置；M16 测试时可按物流单号构造临时配置。
Public Type ConfigStruct
    MaxBacktrackCount As Long
    DebugLogLevel     As String
    DetailedLogLimit  As Long
    LotCaseSensitive  As Boolean
    NoExpirySentinel  As String
End Type

' 本次运行汇总统计。M15 BuildRunStats 统一构造，Dry Run 时分配相关字段为 0。
Public Type RunStats
    TotalBacktrackCount As Long
    MaxGroupBacktrack   As Long
    ValidationFailCount As Long
    AllocSuccessCount   As Long
    AllocFailCount      As Long
    InputReturnRows     As Long
    InputInventoryRows  As Long
    InputShipmentCount  As Long   ' 两表去重合并后的物流单号总数（需求 §5.6）
End Type

' M13 输出构建统一行结构。
' 说明：不同输出表列数不同（汇总/明细/异常/调试/运行历史），
' 因此用 Variant 数组承载整行值，保持接口统一且不绑定固定列宽。
Public Type OutputRow
    Values As Variant   ' 一维数组（建议 1-based），按目标表头顺序填充
End Type

' M07 候选库存行：FilterCandidatePool 返回、M08 TryAllocate 接收的数据结构。
' 字段与 InventoryRow 一致，语义上代表"当前退单行此刻可选用的库存格"。
' M08 的 CompareByPriority 按 QC 优先级（ZP>QC>NG）→效期→批号 对候选排序。
Public Type CandidateRow
    ShipmentNo  As String
    SKU         As String
    QC          As String
    LotNo       As String
    Expiry      As String
    OriginalQty As Long
    CurrentQty  As Long
End Type

' M07 RunPrecheck 的预检测结论。
' 任一字段为 True，代表该 SKU 组在分配前已确定失败（E09），M09 可直接短路。
Public Type PrecheckResult
    PrecheckAHit As Boolean  ' 预检测A：排序后某行初始可用QC数=0，无论如何分配都会失败
    PrecheckBHit As Boolean  ' 预检测B：多行均锁定到同一QC，合计需求 > 该QC当前供应量
End Type
