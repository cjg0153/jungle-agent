param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [string]$OutputPath = "",
    [int]$MaxCells = 5000
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "Excel 파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
if ($ext -notin @(".xlsx",".xlsm",".xls")) {
    throw "Excel 전용 추출기입니다: $ext"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $tempDir = "C:\AI\work\temp"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $OutputPath = Join-Path $tempDir "${base}_xlsx_qa_extract_$(Get-Date -Format yyyyMMdd_HHmmss).json"
}

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($FilePath,0,$true)

    $sheetObjs = New-Object System.Collections.Generic.List[object]
    $allLines = New-Object System.Collections.Generic.List[string]
    $totalCells = 0
    $formulaCells = 0
    $formulaErrors = 0
    $truncated = $false

    foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        $rows = [int]$used.Rows.Count
        $cols = [int]$used.Columns.Count

        $allLines.Add("SHEET: $($ws.Name) USED_RANGE=$($used.Address()) ROWS=$rows COLS=$cols")
        $cellsForSheet = New-Object System.Collections.Generic.List[object]

        for ($r=1; $r -le $rows; $r++) {
            for ($c=1; $c -le $cols; $c++) {
                if ($totalCells -ge $MaxCells) {
                    $truncated = $true
                    break
                }

                $cell = $used.Cells.Item($r,$c)
                $text = [string]$cell.Text
                $formula = [string]$cell.Formula
                $value = $cell.Value2

                if ([string]::IsNullOrWhiteSpace($text) -and [string]::IsNullOrWhiteSpace($formula)) {
                    continue
                }

                $totalCells++
                if (-not [string]::IsNullOrWhiteSpace($formula) -and $formula.StartsWith("=")) {
                    $formulaCells++
                }

                if ($text -match '^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!)$') {
                    $formulaErrors++
                }

                $addr = [string]$cell.Address($false,$false)
                $fmt = [string]$cell.NumberFormat

                $cellsForSheet.Add([PSCustomObject]@{
                    address = $addr
                    text = $text
                    value = $value
                    formula = $formula
                    number_format = $fmt
                })

                $suffix = ""
                if (-not [string]::IsNullOrWhiteSpace($formula) -and $formula.StartsWith("=")) {
                    $suffix = " | FORMULA=$formula"
                }

                $allLines.Add("CELL($($ws.Name)!$addr): $text$suffix | FORMAT=$fmt")
            }

            if ($truncated) { break }
        }

        $sheetObjs.Add([PSCustomObject]@{
            name = [string]$ws.Name
            used_range = [string]$used.Address()
            rows = $rows
            columns = $cols
            cells = @($cellsForSheet)
        })

        if ($truncated) { break }
    }

    $result = [PSCustomObject]@{
        source = $FilePath
        artifact = "XLSX"
        sheet_count = [int]$wb.Worksheets.Count
        captured_cells = $totalCells
        formula_cells = $formulaCells
        formula_errors = $formulaErrors
        truncated = $truncated
        max_cells = $MaxCells
        coverage = [PSCustomObject]@{
            values = $true
            formulas = $true
            number_formats = $true
            charts_images = $false
            rendered_visual_layout = $false
        }
        sheets = @($sheetObjs)
        all_qa_text = ($allLines -join "`r`n")
    }

    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText(
        $OutputPath,
        ($result | ConvertTo-Json -Depth 12),
        $enc
    )

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE XLSX QA EXTRACTOR"
    Write-Host "========================================"
    Write-Host "Sheets         :" $result.sheet_count
    Write-Host "Captured cells :" $result.captured_cells
    Write-Host "Formula cells  :" $result.formula_cells
    Write-Host "Formula errors :" $result.formula_errors
    Write-Host "Truncated      :" $result.truncated
    Write-Host "File           :" $OutputPath
    Write-Host "STATUS: XLSX_QA_EXTRACT_COMPLETE"
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }

    if ($null -ne $wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
