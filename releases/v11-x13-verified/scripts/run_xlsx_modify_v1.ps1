param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw "Excel 파일 누락: $FilePath" }

$tempDir = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$ext = [System.IO.Path]::GetExtension($FilePath)
$extractPath = Join-Path $tempDir "xlsx_modify_extract_${stamp}.json"
$instructionPath = Join-Path $tempDir "xlsx_modify_instruction_${stamp}.txt"
$planPath = Join-Path $tempDir "xlsx_modify_plan_${stamp}.json"
$outputPath = Join-Path $outputDir "${baseName}_modified_${stamp}${ext}"

& "C:\AI\scripts\extract_xlsx_for_qa.ps1" -FilePath $FilePath -OutputPath $extractPath

$instruction = @"
You are Jungle XLSX edit planner.

Workbook extract:
$extractPath

User request:
$UserRequest

Create a MINIMAL cell-edit plan.

Rules:
- Never invent unsupported business facts or numeric values.
- Preserve formulas unless a correction is unambiguous.
- No sheet add/delete/rename.
- No row/column insertion/deletion.
- Prefer text/label/number-format fixes.
- Formula edits only when clearly supported.
- Maximum 100 edits.
- Exact sheet names and A1 addresses.
- Include ONLY fields that should change.
- If no safe edit exists, edits=[].

Save JSON:
$planPath

Schema:
{
  "edits":[
    {
      "sheet":"Sheet1",
      "address":"B3",
      "new_value":"text or number",
      "new_formula":"=SUM(B4:B8)",
      "number_format":"0.0%"
    }
  ]
}

After saving output only XLSX_EDIT_PLAN_CREATED.
"@

$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($instructionPath,$instruction,$enc)

$req = "Read $instructionPath and create the JSON plan exactly. Output only XLSX_EDIT_PLAN_CREATED."
claude -p $req --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath -PathType Leaf)) { throw "XLSX edit plan 생성 실패" }

$plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
$edits = @($plan.edits)

Copy-Item -LiteralPath $FilePath -Destination $outputPath -Force

if ($edits.Count -gt 0) {
    $excel = $null
    $wb = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($outputPath)

        foreach ($e in $edits) {
            $ws = $wb.Worksheets.Item([string]$e.sheet)
            $cell = $ws.Range([string]$e.address)

            $hasFormula =
                ($null -ne $e.PSObject.Properties["new_formula"]) -and
                (-not [string]::IsNullOrWhiteSpace([string]$e.new_formula))

            $hasValue =
                ($null -ne $e.PSObject.Properties["new_value"]) -and
                ($null -ne $e.new_value)

            if ($hasFormula) {
                $cell.Formula = [string]$e.new_formula
            }
            elseif ($hasValue) {
                $cell.Value2 = $e.new_value
            }

            if (
                $null -ne $e.PSObject.Properties["number_format"] -and
                -not [string]::IsNullOrWhiteSpace([string]$e.number_format)
            ) {
                $cell.NumberFormat = [string]$e.number_format
            }
        }

        $wb.Save()
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($true) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        if ($null -ne $wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE XLSX MODIFY RESULT"
Write-Host "========================================"
Write-Host "Edits :" $edits.Count
Write-Host "File  :" $outputPath
Write-Host "STATUS: XLSX_MODIFY_COMPLETE"
