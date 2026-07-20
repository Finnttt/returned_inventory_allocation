# 生成可直接交给业务用户使用的 Excel 宏工作簿。
# 使用前请关闭同名工作簿；若文件已存在，脚本会先在项目上级目录创建备份。

param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "退货入库分配系统.xlsm")
)

$ErrorActionPreference = "Stop"
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

function Confirm-TargetIsClosed([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None)
        $stream.Close()
    } catch {
        throw "目标工作簿正在使用中，请先关闭后重试：$Path"
    }
}

function Backup-ExistingWorkbook([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $parent = Split-Path (Split-Path $Path -Parent) -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $backupPath = Join-Path $parent ("{0}_backup_{1}.xlsm" -f $name, $stamp)
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    return $backupPath
}

function Set-TextCell($ws, [int]$Row, [int]$Column, [string]$Value) {
    $cell = $ws.Cells($Row, $Column)
    $cell.NumberFormat = "@"
    $cell.Value2 = $Value
}

function Set-Headers($ws, [string[]]$Headers, [int]$Color = 0xD9EAF7) {
    for ($column = 1; $column -le $Headers.Count; $column++) {
        Set-TextCell $ws 1 $column $Headers[$column - 1]
    }

    $headerRange = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $Headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = $Color
    $headerRange.Borders.LineStyle = 1
    $headerRange.AutoFilter() | Out-Null
    $ws.Rows(1).RowHeight = 24
    $ws.Columns.AutoFit() | Out-Null
}

function Add-Worksheet($wb, [string]$Name) {
    $ws = $wb.Worksheets.Add(
        [System.Reflection.Missing]::Value,
        $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $Name
    return $ws
}

function Add-ActionButton($ws, [string]$Caption, [string]$MacroName, [double]$Top, [int]$Color) {
    $shape = $ws.Shapes.AddShape(1, 30, $Top, 180, 42)
    $shape.Name = "btn_" + $MacroName
    $shape.TextFrame.Characters().Text = $Caption
    $shape.TextFrame.Characters().Font.Bold = $true
    $shape.Fill.ForeColor.RGB = $Color
    $shape.Line.ForeColor.RGB = 0x808080
    $shape.OnAction = $MacroName
}

# VBA 编辑器导入 .bas 时使用 Windows ANSI，而项目源码使用 UTF-8。
# 因此只为导入生成临时 ANSI 副本，并补齐模块名；不改动 UTF-8 源文件。
function New-VbaImportFile([System.IO.FileInfo]$SourceFile, [string]$TempDirectory) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($SourceFile.Name)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $text = [System.IO.File]::ReadAllText($SourceFile.FullName, $utf8)
    $text = [System.Text.RegularExpressions.Regex]::Replace(
        $text,
        '(?m)^Attribute VB_Name = ".*"\r?\n',
        '')
    $text = "Attribute VB_Name = `"$moduleName`"`r`n" + $text

    $tempPath = Join-Path $TempDirectory $SourceFile.Name
    [System.IO.File]::WriteAllText($tempPath, $text, [System.Text.Encoding]::Default)
    return $tempPath
}

Confirm-TargetIsClosed $OutputPath
$backupPath = Backup-ExistingWorkbook $OutputPath
$vbaTempDirectory = Join-Path $env:TEMP ("ria_vba_import_" + [System.Guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory($vbaTempDirectory) | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Add()

    # 不依赖用户的“新建工作簿默认 Sheet 数量”，统一只保留第一张。
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $panel = $wb.Worksheets.Item(1)
    $panel.Name = "操作面板"
    Set-TextCell $panel 1 1 "退货入库分配系统"
    Set-TextCell $panel 3 1 "推荐流程：①粘贴两张输入表 → ②仅运行校验 → ③修正异常 → ④开始分配"
    Set-TextCell $panel 4 1 "注意：行号必须以文本录入五位数字，例如 00001；系统不会自动补零。"
    Set-TextCell $panel 5 1 "首次打开时请点击“启用内容/启用宏”，否则按钮无法运行。"
    $panel.Cells(1, 1).Font.Bold = $true
    $panel.Cells(1, 1).Font.Size = 20
    $panel.Columns("A:A").ColumnWidth = 88
    Add-ActionButton $panel "仅运行校验" "StartValidationOnly" 130 0xF4B183
    Add-ActionButton $panel "开始分配" "StartFullAllocation" 185 0xA9D18E
    Add-ActionButton $panel "清空结果" "ClearAllocationResults" 240 0xBDD7EE

    $returnWs = Add-Worksheet $wb "输入_退单表"
    Set-Headers $returnWs @("物流单号", "WMS退单号", "SKU", "行号", "数量")
    $returnWs.Columns(4).NumberFormat = "@"

    $inventoryWs = Add-Worksheet $wb "输入_质检库存表"
    Set-Headers $inventoryWs @("物流单号", "SKU", "QC情况", "批号", "效期", "数量")
    $inventoryWs.Columns(4).NumberFormat = "@"
    $inventoryWs.Columns(5).NumberFormat = "@"

    $configWs = Add-Worksheet $wb "输入_配置"
    Set-Headers $configWs @("参数名", "值", "说明") 0xFFF2CC
    $configRows = @(
        @("最大回溯次数", "200", "单个物流单号允许的最大回溯次数，须为正整数"),
        @("调试日志级别", "关闭", "允许值：关闭、简版、详细"),
        @("详细日志单表上限", "100000", "详细日志超过此行数时自动分表"),
        @("批号比较模式", "不敏感", "允许值：不敏感、敏感"),
        @("无保质期哨兵值", "2099/01/01", "无保质期商品统一使用的占位效期")
    )
    for ($row = 0; $row -lt $configRows.Count; $row++) {
        for ($column = 0; $column -lt 3; $column++) {
            Set-TextCell $configWs ($row + 2) ($column + 1) $configRows[$row][$column]
        }
    }
    $configWs.Columns.AutoFit() | Out-Null

    $summaryWs = Add-Worksheet $wb "分配状态汇总表"
    Set-Headers $summaryWs @("物流单号", "WMS退单号", "退单号状态", "原因") 0xE2F0D9

    $detailWs = Add-Worksheet $wb "成功分配明细表"
    Set-Headers $detailWs @(
        "物流单号", "WMS退单号", "SKU", "行号", "退单数量", "QC情况",
        "批号", "效期", "分配数量", "行状态", "退单号状态") 0xE2F0D9
    $detailWs.Columns(4).NumberFormat = "@"

    $anomalyWs = Add-Worksheet $wb "数据异常明细表"
    Set-Headers $anomalyWs @(
        "来源表", "Excel行号", "物流单号", "WMS退单号", "SKU",
        "字段名", "原始值", "错误码", "原因说明") 0xFCE4D6

    $debugWs = Add-Worksheet $wb "调试日志"
    Set-Headers $debugWs @(
        "物流单号", "SKU", "WMS退单号", "行号", "D", "处理序",
        "动态nextMinQty", "候选QC数", "被排除QC列表", "策略", "分配QC",
        "分配前QC剩余", "分配后QC剩余", "批号/效期组合数", "是否回溯重试",
        "实际回溯次数", "行状态", "错误码", "分配失败子类型") 0xEDEDED

    $historyWs = Add-Worksheet $wb "运行历史记录表"
    Set-Headers $historyWs @(
        "运行编号", "运行时间", "运行类型", "输入：退单表行数", "输入：质检库存表行数",
        "输入：物流单号数", "校验耗时（秒）", "分配耗时（秒）", "总耗时（秒）",
        "校验失败物流单号数", "分配成功物流单号数", "分配失败物流单号数",
        "错误码分布", "总回溯次数", "最大单组回溯次数", "调试日志级别",
        "备注", "最大回溯次数", "批号比较模式", "无保质期哨兵值") 0xD9EAD3

    # 生产工作簿只导入 M01～M15；测试运行器留在开发测试工作簿中。
    $productionModules = @(
        "modTypes.bas",
        "modConfig.bas",
        "modExcelInput.bas",
        "modNormalize.bas",
        "modValidate.bas",
        "modInventoryLedger.bas",
        "modSortFilter.bas",
        "modStrategies.bas",
        "modBacktracking.bas",
        "modGuards.bas",
        "modStatus.bas",
        "modPostValidate.bas",
        "modOutputBuilder.bas",
        "modExcelOutput.bas",
        "modRunner.bas"
    )

    foreach ($moduleName in $productionModules) {
        $modulePath = Join-Path $PSScriptRoot $moduleName
        if (-not (Test-Path -LiteralPath $modulePath)) {
            throw "缺少生产模块：$modulePath"
        }
        $importPath = New-VbaImportFile (Get-Item -LiteralPath $modulePath) $vbaTempDirectory
        $wb.VBProject.VBComponents.Import($importPath) | Out-Null
    }

    $wb.Worksheets.Item("操作面板").Activate()
    $wb.SaveAs($OutputPath, 52)
    $wb.Close($true)
    $wb = $null

    Write-Output "OUTPUT=$OutputPath"
    if ($backupPath) { Write-Output "BACKUP=$backupPath" }
    Write-Output "MODULES=$($productionModules.Count)"
    Write-Output "SHEETS=9"
} finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (Test-Path -LiteralPath $vbaTempDirectory) {
        Remove-Item -LiteralPath $vbaTempDirectory -Recurse -Force
    }
    [System.GC]::Collect()
}
