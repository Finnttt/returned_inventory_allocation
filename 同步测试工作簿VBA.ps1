# 将当前目录的全部 VBA 模块同步到测试汇总工作簿。
# 脚本会先确认目标文件未被占用，并在项目上级目录保留备份。

param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "测试用例部分汇总.xlsm")
)

$ErrorActionPreference = "Stop"
$WorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "测试汇总工作簿不存在：$WorkbookPath"
}

try {
    $stream = [System.IO.File]::Open(
        $WorkbookPath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None)
    $stream.Close()
} catch {
    throw "测试汇总工作簿正在使用中，请先关闭后重试：$WorkbookPath"
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Split-Path (Split-Path $WorkbookPath -Parent) -Parent
$backupPath = Join-Path $backupDir ("测试用例部分汇总_VBA同步前_{0}.xlsm" -f $stamp)
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force

$moduleFiles = Get-ChildItem -LiteralPath $PSScriptRoot -Filter "mod*.bas" | Sort-Object Name
if ($moduleFiles.Count -ne 17) {
    throw "预期找到 17 个 VBA 模块，实际找到 $($moduleFiles.Count) 个；为避免漏导入，已中止。"
}

# Excel VBA 导入器按 Windows ANSI 读取 .bas，而项目源码是 UTF-8。
# 临时转换只用于导入，避免中文字符串乱码和引号被破坏。
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

$vbaTempDirectory = Join-Path $env:TEMP ("ria_vba_import_" + [System.Guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory($vbaTempDirectory) | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Open($WorkbookPath, 0, $false)

    # 将旧接口名迁移到需求和 TC 统一使用的新名称。
    $oldHistory = $null
    $newHistory = $null
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq "运行历史记录") { $oldHistory = $ws }
        if ($ws.Name -eq "运行历史记录表") { $newHistory = $ws }
    }
    if ($oldHistory -ne $null -and $newHistory -eq $null) {
        $oldHistory.Name = "运行历史记录表"
    }

    # 这是“全量同步”脚本：先移除全部标准模块，防止旧版“模块1”等匿名模块残留。
    # ThisWorkbook 和工作表代码模块的类型不是 1，因此不会被触碰。
    $standardModules = @()
    foreach ($component in $wb.VBProject.VBComponents) {
        if ($component.Type -eq 1) { $standardModules += $component }
    }
    foreach ($component in $standardModules) {
        $wb.VBProject.VBComponents.Remove($component)
    }

    foreach ($file in $moduleFiles) {
        $importPath = New-VbaImportFile $file $vbaTempDirectory
        $wb.VBProject.VBComponents.Import($importPath) | Out-Null
    }

    $wb.Save()
    $wb.Close($false)
    $wb = $null

    Write-Output "WORKBOOK=$WorkbookPath"
    Write-Output "BACKUP=$backupPath"
    Write-Output "MODULES=$($moduleFiles.Count)"
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
