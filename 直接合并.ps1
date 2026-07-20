# ==========================================================
# 直接合并.ps1  （纯 ZipFile 方案，不启动任何 Excel 进程）
# 功能:
#   1. 修复 xlsm 的 预期_数据异常明细（9列全写）
#   2. 追加 batch1 独立xlsx 的输入表数据
# ==========================================================
Add-Type -AssemblyName System.IO.Compression.FileSystem

$workDir  = $PSScriptRoot
$xlsmPath = Join-Path $workDir "测试用例部分汇总.xlsm"

$batch1Files = @(
    "SF0013_测试数据.xlsx","SF0014_测试数据.xlsx","SF0015_测试数据.xlsx",
    "SF0037_测试数据.xlsx","SF0051_测试数据.xlsx","SF0056_测试数据.xlsx",
    "SF0057_测试数据.xlsx","SF0058_测试数据.xlsx","SF0059_测试数据.xlsx",
    "SF0060_测试数据.xlsx"
)
$batchByShip = @{
    "SF3190000000013"="批次1";"SF3190000000014"="批次1";"SF3190000000015"="批次1"
    "SF3190000000037"="批次1";"SF3190000000051"="批次1";"SF3190000000056"="批次1"
    "SF3190000000057"="批次1";"SF3190000000058"="批次1";"SF3190000000059"="批次1"
    "SF3190000000060"="批次1"
}

# ── 列号转列字母（1→A, 26→Z, 27→AA ...）──────────────────────
function Col-Letter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}

# ── 读取 xlsx 内的 SharedStrings ──────────────────────────────
function Get-SS($zip) {
    $e = $zip.GetEntry("xl/sharedStrings.xml")
    if (-not $e) { return @() }
    $ms = New-Object System.IO.MemoryStream; $e.Open().CopyTo($ms)
    $xml = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
    $out = @()
    foreach ($m in [regex]::Matches($xml, '<si>(.*?)</si>', 'Singleline')) {
        $ts = [regex]::Matches($m.Value, '<t[^>]*>([^<]*)</t>')
        $out += ($ts | ForEach-Object { $_.Groups[1].Value }) -join ""
    }
    return $out
}

# ── 从 sheetXml 中读取所有行（跳过表头行1） ──────────────────
function Read-AllRows($sheetXml, $ss) {
    $rows = @()
    $rowMatches = [regex]::Matches($sheetXml, '<row r="(\d+)"[^>]*>(.*?)</row>', 'Singleline')
    foreach ($rm in $rowMatches) {
        $rn = [int]$rm.Groups[1].Value
        if ($rn -lt 2) { continue }
        $cells = [ordered]@{}
        foreach ($c in [regex]::Matches($rm.Groups[2].Value,
                '<c r="([A-Z]+)' + $rn + '"[^>]*>(.*?)</c>', 'Singleline')) {
            $col = $c.Groups[1].Value
            $cv  = [regex]::Match($c.Value, '<v>([^<]*)</v>').Groups[1].Value
            # inline string
            $it  = [regex]::Match($c.Value, '<is><t[^>]*>([^<]*)</t></is>').Groups[1].Value
            if ($it -ne "") { $cells[$col] = $it }
            elseif ($c.Value -match 't="s"' -and $cv -ne "") {
                $i = [int]$cv; $cells[$col] = if ($i -lt $ss.Count) { $ss[$i] } else { "" }
            } else { $cells[$col] = $cv }
        }
        $rows += ,$cells
    }
    return $rows
}

# ── 读取 xlsx 某工作表所有数据行 ──────────────────────────────
function Get-SheetRows($zip, $sheetName) {
    # 找 sheetIndex
    $msWb = New-Object System.IO.MemoryStream
    $zip.GetEntry("xl/workbook.xml").Open().CopyTo($msWb)
    $wbXml = [System.Text.Encoding]::UTF8.GetString($msWb.ToArray())
    $nodes = [regex]::Matches($wbXml, '<sheet name="([^"]+)"[^/]*/>')
    $idx = -1; $i2 = 0
    foreach ($n in $nodes) { $i2++; if ($n.Groups[1].Value -eq $sheetName) { $idx = $i2 } }
    if ($idx -lt 0) { return @() }

    $ss = Get-SS $zip
    $msS = New-Object System.IO.MemoryStream
    $zip.GetEntry("xl/worksheets/sheet${idx}.xml").Open().CopyTo($msS)
    $shXml = [System.Text.Encoding]::UTF8.GetString($msS.ToArray())
    return Read-AllRows $shXml $ss
}

# ── 生成单元格 XML（inline string 或数字） ────────────────────
function Make-CellXml([string]$colLetter, [int]$rowNum, $val, [bool]$forceText = $false) {
    $ref = "${colLetter}${rowNum}"
    if ($val -eq $null -or "$val" -eq "") { return "" }
    $s = "$val"
    if (-not $forceText -and $s -match "^-?[0-9]+(\.[0-9]+)?$") {
        return "<c r=`"$ref`"><v>$s</v></c>"
    } else {
        $esc = $s -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" `
                  -replace '"',"&quot;" -replace "'","&apos;"
        return "<c r=`"$ref`" t=`"inlineStr`"><is><t>$esc</t></is></c>"
    }
}

# ── 在 sheetXml 中找当前最大行号 ─────────────────────────────
function Get-MaxRowNum([string]$sheetXml) {
    $nums = [regex]::Matches($sheetXml, '<row r="(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value }
    if ($nums.Count -eq 0) { return 1 }
    return ($nums | Measure-Object -Maximum).Maximum
}

# ── 删除 sheetXml 中所有 r>=2 的行 ───────────────────────────
function Clear-DataRows([string]$sheetXml) {
    return [regex]::Replace($sheetXml, '<row r="([2-9]|\d{2,})"[^>]*>.*?</row>', '',
        [System.Text.RegularExpressions.RegexOptions]::Singleline)
}

# ── 在 sheetXml 的 </sheetData> 前插入行 XML ─────────────────
function Insert-RowsBefore($sheetXml, $rowsXml) {
    return $sheetXml -replace "</sheetData>", "$rowsXml</sheetData>"
}

# ── 把一组数据行（有序字典数组）转成 XML 字符串 ───────────────
# colHeaders: 目标列的顺序列表（不含运行批次，那是第0列）
# batchName:  若 $null 则不写批次列
function Rows-To-Xml($dataRows, $startRowNum, $colHeaders, $batchName) {
    $sb = New-Object System.Text.StringBuilder
    $r = $startRowNum
    foreach ($row in $dataRows) {
        $cellsXml = ""
        if ($batchName -ne $null) {
            $cellsXml += Make-CellXml "A" $r $batchName
        }
        $colOffset = if ($batchName -ne $null) { 2 } else { 1 }
        for ($ci = 0; $ci -lt $colHeaders.Count; $ci++) {
            $cLetter = Col-Letter ($ci + $colOffset)
            $val = if ($row.Contains($colHeaders[$ci])) { $row[$colHeaders[$ci]] } else { "" }
            $cellsXml += Make-CellXml $cLetter $r $val
        }
        [void]$sb.Append("<row r=`"$r`">$cellsXml</row>")
        $r++
    }
    return $sb.ToString()
}

# ── 读取 xlsx sheetXml 并返回字符串 ──────────────────────────
function Get-SheetXml($zip, $idx) {
    $ms = New-Object System.IO.MemoryStream
    $zip.GetEntry("xl/worksheets/sheet${idx}.xml").Open().CopyTo($ms)
    return [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
}

# ── 获取 xlsm 的工作表名称→索引 映射 ─────────────────────────
function Get-SheetIndex($zip, $sheetName) {
    $ms = New-Object System.IO.MemoryStream
    $zip.GetEntry("xl/workbook.xml").Open().CopyTo($ms)
    $wbXml = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
    $nodes = [regex]::Matches($wbXml, '<sheet name="([^"]+)"[^/]*/>')
    for ($i = 0; $i -lt $nodes.Count; $i++) {
        if ($nodes[$i].Groups[1].Value -eq $sheetName) { return $i + 1 }
    }
    return -1
}

# ── 修改 ZipArchive 里某个 entry 的内容 ──────────────────────
function Update-ZipEntry($archive, $entryPath, $newContent) {
    $entry = $archive.GetEntry($entryPath)
    if ($entry -ne $null) { $entry.Delete() }
    $newEntry = $archive.CreateEntry($entryPath)
    $sw = [System.IO.StreamWriter]::new($newEntry.Open(), [System.Text.Encoding]::UTF8)
    $sw.Write($newContent)
    $sw.Flush()
    $sw.Close()
}

# ==========================================================
# 主程序开始
# ==========================================================
Write-Host "=== 直接合并（纯 ZipFile 方案）===" -ForegroundColor Cyan

# ── A. 读取所有源数据 ────────────────────────────────────────
Write-Host "`n[A] 读取源 xlsx 数据..."
$srcAnomalData  = @()  # for 预期_数据异常明细
$srcReturnData  = @()  # for 输入_退单表
$srcInventData  = @()  # for 输入_质检库存表
$srcConfigData  = @()  # for 输入_配置

foreach ($fname in $batch1Files) {
    $fpath = Join-Path $workDir $fname
    if (-not (Test-Path $fpath)) { Write-Host "  未找到 $fname，跳过"; continue }
    $z = [System.IO.Compression.ZipFile]::OpenRead($fpath)

    # 读取物流单号（config 第2行 A列）
    $cfgRows = Get-SheetRows $z "输入_配置"
    $shipNo = if ($cfgRows.Count -gt 0 -and $cfgRows[0].Contains("A")) { $cfgRows[0]["A"] } else { "?" }
    $batchName = if ($batchByShip.ContainsKey($shipNo)) { $batchByShip[$shipNo] } else { "批次1" }
    Write-Host "  $fname => 物流单号=$shipNo [$batchName]"

    # 预期_数据异常明细（可能叫"预期_数据异常明细"或"预期_数据异常明细表"）
    $anomRows = Get-SheetRows $z "预期_数据异常明细"
    if ($anomRows.Count -eq 0) { $anomRows = Get-SheetRows $z "预期_数据异常明细表" }
    foreach ($row in $anomRows) { $row["__batch"] = $batchName }
    $srcAnomalData += $anomRows
    Write-Host "    异常明细: $($anomRows.Count) 行"

    # 输入表
    $rtnRows = Get-SheetRows $z "输入_退单表"
    foreach ($row in $rtnRows) {
        # 忽略仅有格式/示例值但没有数量的残留空行。
        if ($row.Contains("A") -and $row.Contains("E") -and "$($row["E"])" -ne "") {
            $srcReturnData += ,$row
        }
    }

    $invRows = Get-SheetRows $z "输入_质检库存表"
    foreach ($row in $invRows) {
        if ($row.Contains("A") -and $row.Contains("F") -and "$($row["F"])" -ne "") {
            $srcInventData += ,$row
        }
    }

    $cfgDataRows = Get-SheetRows $z "输入_配置"
    foreach ($row in $cfgDataRows) { $srcConfigData += ,$row }

    Write-Host "    退单表:$($rtnRows.Count) 库存表:$($invRows.Count) 配置:$($cfgDataRows.Count)"
    $z.Dispose()
}
Write-Host "  汇总 — 异常明细=$($srcAnomalData.Count) 退单表=$($srcReturnData.Count) 库存表=$($srcInventData.Count) 配置=$($srcConfigData.Count)"

# ── B. 修改 xlsm ────────────────────────────────────────────
Write-Host "`n[B] 修改 xlsm..."
$tmpPath = $xlsmPath + ".tmp"
Copy-Item $xlsmPath $tmpPath -Force

$archive = [System.IO.Compression.ZipFile]::Open($tmpPath,
    [System.IO.Compression.ZipArchiveMode]::Update)

# 读取 xlsm 的 sheet 索引
$msWb2 = New-Object System.IO.MemoryStream
$archive.GetEntry("xl/workbook.xml").Open().CopyTo($msWb2)
$wbXml2 = [System.Text.Encoding]::UTF8.GetString($msWb2.ToArray())
$xlsmNodes = [regex]::Matches($wbXml2, '<sheet name="([^"]+)"[^/]*/>')
Write-Host "  xlsm 工作表数: $($xlsmNodes.Count)"

function Get-XlsmSheetIdx($name) {
    for ($i = 0; $i -lt $xlsmNodes.Count; $i++) {
        if ($xlsmNodes[$i].Groups[1].Value -eq $name) { return $i + 1 }
    }
    return -1
}

# ── B1. 修复 预期_数据异常明细 ──────────────────────────────
Write-Host "`n  [B1] 修复 预期_数据异常明细..."
$anomIdx = Get-XlsmSheetIdx "预期_数据异常明细"
Write-Host "    Sheet index: $anomIdx"

if ($anomIdx -gt 0) {
    $anomXml = Get-SheetXml $archive $anomIdx
    # 清除 r>=2 的行
    $anomXml = Clear-DataRows $anomXml
    # 写新数据
    # 目标列顺序（对应源表的列字母）：A=来源表 B=原始行号 C=物流单号 D=WMS退单号 E=SKU F=异常字段名 G=原始值 H=错误码 I=原因说明
    $anomColMap = @("A","B","C","D","E","F","G","H","I")
    $sb = New-Object System.Text.StringBuilder
    $r = 2
    foreach ($row in $srcAnomalData) {
        $bName = if ($row.Contains("__batch")) { $row["__batch"] } else { "批次1" }
        $cellXml = Make-CellXml "A" $r $bName  # 运行批次列（目标A列）
        for ($ci = 0; $ci -lt $anomColMap.Count; $ci++) {
            $srcCol = $anomColMap[$ci]
            $dstCol = Col-Letter ($ci + 2)  # 目标从B列开始
            $val = if ($row.Contains($srcCol)) { $row[$srcCol] } else { "" }
            $cellXml += Make-CellXml $dstCol $r $val
        }
        [void]$sb.Append("<row r=`"$r`">$cellXml</row>")
        $r++
    }
    $anomXml = Insert-RowsBefore $anomXml $sb.ToString()
    Update-ZipEntry $archive "xl/worksheets/sheet${anomIdx}.xml" $anomXml
    Write-Host "    写入 $($srcAnomalData.Count) 行（A~J列）"
}

# ── B2. 追加 输入_退单表 ──────────────────────────────────────
Write-Host "`n  [B2] 追加 输入_退单表..."
$rtnIdx = Get-XlsmSheetIdx "输入_退单表"
if ($rtnIdx -gt 0 -and $srcReturnData.Count -gt 0) {
    $rtnXml = Get-SheetXml $archive $rtnIdx
    $maxR = Get-MaxRowNum $rtnXml
    # 退单表列: A=物流单号 B=WMS退单号 C=SKU D=行号 E=数量
    $colMap = @("A","B","C","D","E")
    $sb = New-Object System.Text.StringBuilder
    $r = $maxR + 1
    foreach ($row in $srcReturnData) {
        $cellXml = ""
        foreach ($ci in 0..($colMap.Count-1)) {
            $val = if ($row.Contains($colMap[$ci])) { $row[$colMap[$ci]] } else { "" }
            # 行号必须保持五位文本；否则 "00001" 会被脚本写成数值 1，触发 E01。
            $cellXml += Make-CellXml (Col-Letter ($ci+1)) $r $val ($colMap[$ci] -eq "D")
        }
        [void]$sb.Append("<row r=`"$r`">$cellXml</row>")
        $r++
    }
    $rtnXml = Insert-RowsBefore $rtnXml $sb.ToString()
    Update-ZipEntry $archive "xl/worksheets/sheet${rtnIdx}.xml" $rtnXml
    Write-Host "    追加 $($srcReturnData.Count) 行（从第 $($maxR+1) 行）"
}

# ── B3. 追加 输入_质检库存表 ─────────────────────────────────
Write-Host "`n  [B3] 追加 输入_质检库存表..."
$invIdx = Get-XlsmSheetIdx "输入_质检库存表"
if ($invIdx -gt 0 -and $srcInventData.Count -gt 0) {
    $invXml = Get-SheetXml $archive $invIdx
    $maxR = Get-MaxRowNum $invXml
    # 库存表列: A=物流单号 B=SKU C=QC情况 D=批号 E=效期 F=数量 G=备注
    $colMap = @("A","B","C","D","E","F","G")
    $sb = New-Object System.Text.StringBuilder
    $r = $maxR + 1
    foreach ($row in $srcInventData) {
        $cellXml = ""
        foreach ($ci in 0..($colMap.Count-1)) {
            $val = if ($row.Contains($colMap[$ci])) { $row[$colMap[$ci]] } else { "" }
            $cellXml += Make-CellXml (Col-Letter ($ci+1)) $r $val
        }
        [void]$sb.Append("<row r=`"$r`">$cellXml</row>")
        $r++
    }
    $invXml = Insert-RowsBefore $invXml $sb.ToString()
    Update-ZipEntry $archive "xl/worksheets/sheet${invIdx}.xml" $invXml
    Write-Host "    追加 $($srcInventData.Count) 行（从第 $($maxR+1) 行）"
}

# ── B4. 追加 输入_配置 ────────────────────────────────────────
Write-Host "`n  [B4] 追加 输入_配置..."
$cfgIdx = Get-XlsmSheetIdx "输入_配置"
if ($cfgIdx -gt 0 -and $srcConfigData.Count -gt 0) {
    $cfgXml = Get-SheetXml $archive $cfgIdx
    $maxR = Get-MaxRowNum $cfgXml
    # 配置列: A=物流单号 B=TC编号 C=最大回溯次数 D=调试日志级别 E=批号比较模式 F=无保质期哨兵值 G=备注
    $colMap = @("A","B","C","D","E","F","G")
    $sb = New-Object System.Text.StringBuilder
    $r = $maxR + 1
    foreach ($row in $srcConfigData) {
        $cellXml = ""
        foreach ($ci in 0..($colMap.Count-1)) {
            $val = if ($row.Contains($colMap[$ci])) { $row[$colMap[$ci]] } else { "" }
            $cellXml += Make-CellXml (Col-Letter ($ci+1)) $r $val
        }
        [void]$sb.Append("<row r=`"$r`">$cellXml</row>")
        $r++
    }
    $cfgXml = Insert-RowsBefore $cfgXml $sb.ToString()
    Update-ZipEntry $archive "xl/worksheets/sheet${cfgIdx}.xml" $cfgXml
    Write-Host "    追加 $($srcConfigData.Count) 行（从第 $($maxR+1) 行）"
}

# ── 关闭并替换 ────────────────────────────────────────────────
$archive.Dispose()
Write-Host "`n[C] 替换原文件..."
Copy-Item $tmpPath $xlsmPath -Force
Remove-Item $tmpPath -Force
Write-Host "✓ 全部完成！" -ForegroundColor Green
Write-Host "=== 脚本结束 ===" -ForegroundColor Cyan
