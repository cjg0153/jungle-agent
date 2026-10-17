param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$LayoutLibraryPath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    throw "원본 PPTX를 찾을 수 없습니다: $SourcePath"
}
if (!(Test-Path -LiteralPath $LayoutLibraryPath -PathType Leaf)) {
    throw "레이아웃 라이브러리를 찾을 수 없습니다: $LayoutLibraryPath"
}

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath      = Join-Path $outputDir "${baseName}_design_v2_${stamp}.pptx"
$contentPath     = Join-Path $tempDir "ppt_v2_source_content_${stamp}.json"
$planPath        = Join-Path $tempDir "ppt_v2_deck_plan_${stamp}.json"
$instructionPath = Join-Path $tempDir "ppt_v2_instruction_${stamp}.txt"

$extractor = "C:\AI\scripts\extract_office_content.ps1"
& $extractor -FilePath $SourcePath -OutputPath $contentPath

if (!(Test-Path -LiteralPath $contentPath -PathType Leaf)) {
    throw "원본 PPT 내용 추출에 실패했습니다."
}

$instruction = @"
You are the PowerPoint design planner for Jungle.

Use ONLY:
SOURCE CONTENT JSON:
$contentPath

LAYOUT LIBRARY JSON:
$LayoutLibraryPath

USER REQUEST:
$UserRequest

Create a NEW executive-quality PowerPoint by reusing the strongest existing slide designs.

Rules:
- Do not make plain title-and-bullet slides unless the library has no better option.
- Select the most suitable layout for each slide.
- Prefer executive summary, cards, comparison, KPI, process, roadmap, table, chart, diagram and rich detail layouts.
- Reuse the source deck's visual language.
- Do not invent business facts absent from the source.
- Keep Korean concise and natural.
- Respect each layout's content_capacity.
- Use exact shape_name values from replaceable_fields only.
- Every replaceable field must receive a replacement.
- Use empty text or empty table_rows when intentionally blank.
- reference_slide must match the selected layout.
- Normally create 4 to 8 slides unless the user requests another count.
- For field_type="table", use kind="table" and provide table_rows as a 2D string array sized to the table when possible.
- For field_type="text", use kind="text" and provide text.

Save valid UTF-8 JSON to:
$planPath

Schema:
{
  "deck_title": "presentation title",
  "slides": [
    {
      "slide_number": 1,
      "layout_id": "L01",
      "reference_slide": 1,
      "purpose": "why this layout fits",
      "replacements": [
        {
          "shape_name": "exact shape name",
          "kind": "text",
          "text": "new text"
        },
        {
          "shape_name": "exact table shape name",
          "kind": "table",
          "table_rows": [
            ["A", "B"],
            ["C", "D"]
          ]
        }
      ]
    }
  ]
}

After saving, output only:
PPT_DESIGN_V2_PLAN_CREATED
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($instructionPath, $instruction, $utf8Bom)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified there. Do not use web search, Google Drive, connectors, or external sources. When finished, output only PPT_DESIGN_V2_PLAN_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath -PathType Leaf)) {
    throw "PPT Design v2 생성계획 JSON이 만들어지지 않았습니다."
}

$plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
$library = Get-Content -LiteralPath $LayoutLibraryPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ($null -eq $plan.slides -or $plan.slides.Count -lt 1) {
    throw "PPT Design v2 생성계획에 slides가 없습니다."
}

$layoutById = @{}
foreach ($layout in $library.layouts) {
    $layoutById[[string]$layout.layout_id] = $layout
}

foreach ($slidePlan in $plan.slides) {
    $layoutId = [string]$slidePlan.layout_id
    if (!$layoutById.ContainsKey($layoutId)) {
        throw "알 수 없는 layout_id입니다: $layoutId"
    }
    $expectedRef = [int]$layoutById[$layoutId].reference_slide
    $actualRef = [int]$slidePlan.reference_slide
    if ($expectedRef -ne $actualRef) {
        throw "layout_id $layoutId 의 reference_slide 불일치"
    }
}

Copy-Item -LiteralPath $SourcePath -Destination $outputPath -Force

function Get-SlideById {
    param($Presentation, [int]$SlideId)
    foreach ($s in $Presentation.Slides) {
        if ([int]$s.SlideID -eq $SlideId) { return $s }
    }
    return $null
}

function Clear-ShapeContent {
    param($Shape)

    try {
        if ($Shape.HasTable -eq -1) {
            for ($r = 1; $r -le $Shape.Table.Rows.Count; $r++) {
                for ($c = 1; $c -le $Shape.Table.Columns.Count; $c++) {
                    $Shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = ""
                }
            }
            return $true
        }
    }
    catch {}

    try {
        if ($Shape.HasTextFrame -eq -1) {
            $Shape.TextFrame.TextRange.Text = ""
            return $true
        }
    }
    catch {}

    return $false
}

function Set-TextContent {
    param($Shape, [string]$Text)

    try {
        if ($Shape.HasTextFrame -eq -1) {
            $Shape.TextFrame.TextRange.Text = $Text
            try { $Shape.TextFrame2.AutoSize = 2 } catch {}
            return $true
        }
    }
    catch {}

    return $false
}

function Set-TableContent {
    param($Shape, $Rows)

    try {
        if ($Shape.HasTable -ne -1) { return $false }

        $maxRows = [int]$Shape.Table.Rows.Count
        $maxCols = [int]$Shape.Table.Columns.Count

        for ($r = 1; $r -le $maxRows; $r++) {
            for ($c = 1; $c -le $maxCols; $c++) {
                $text = ""
                if ($null -ne $Rows -and ($r - 1) -lt $Rows.Count) {
                    $row = @($Rows[$r - 1])
                    if (($c - 1) -lt $row.Count) {
                        $text = [string]$row[$c - 1]
                    }
                }

                $cellShape = $Shape.Table.Cell($r, $c).Shape
                $cellShape.TextFrame.TextRange.Text = $text
                try { $cellShape.TextFrame2.AutoSize = 2 } catch {}
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

$ppt = $null
$pres = $null
$createdCount = 0
$replacementCount = 0
$replacementFailures = @()

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($outputPath, $false, $false, $false)

    $originalCount = [int]$pres.Slides.Count
    $originalSlideIds = @{}
    $allOriginalIds = @()

    for ($i = 1; $i -le $originalCount; $i++) {
        $sid = [int]$pres.Slides.Item($i).SlideID
        $originalSlideIds[$i] = $sid
        $allOriginalIds += $sid
    }

    foreach ($slidePlan in $plan.slides) {
        $layoutId = [string]$slidePlan.layout_id
        $layout = $layoutById[$layoutId]
        $referenceSlideNumber = [int]$slidePlan.reference_slide

        if (!$originalSlideIds.ContainsKey($referenceSlideNumber)) {
            throw "참조 슬라이드를 찾을 수 없습니다: $referenceSlideNumber"
        }

        $referenceId = [int]$originalSlideIds[$referenceSlideNumber]
        $referenceSlide = Get-SlideById -Presentation $pres -SlideId $referenceId
        if ($null -eq $referenceSlide) {
            throw "참조 슬라이드 ID를 찾을 수 없습니다: $referenceId"
        }

        $dupRange = $referenceSlide.Duplicate()
        $newSlide = $dupRange.Item(1)
        $newSlide.MoveTo($pres.Slides.Count)
        $createdCount++

        foreach ($field in $layout.replaceable_fields) {
            $shapeName = [string]$field.shape_name
            try {
                $shape = $newSlide.Shapes.Item($shapeName)
                [void](Clear-ShapeContent -Shape $shape)
            }
            catch {
                $replacementFailures += "Slide $createdCount clear failed: $shapeName"
            }
        }

        foreach ($replacement in $slidePlan.replacements) {
            $shapeName = [string]$replacement.shape_name
            $kind = [string]$replacement.kind

            try {
                $shape = $newSlide.Shapes.Item($shapeName)

                $ok = $false
                if ($kind -eq "table") {
                    $ok = Set-TableContent -Shape $shape -Rows $replacement.table_rows
                }
                else {
                    $ok = Set-TextContent -Shape $shape -Text ([string]$replacement.text)

                    if (-not $ok) {
                        try {
                            if ($shape.HasTable -eq -1) {
                                $lines = @(([string]$replacement.text) -split "`r?`n")
                                $rows = @()
                                foreach ($line in $lines) {
                                    $rows += ,@($line -split '\s*\|\s*')
                                }
                                $ok = Set-TableContent -Shape $shape -Rows $rows
                            }
                        }
                        catch {}
                    }
                }

                if ($ok) { $replacementCount++ }
                else { $replacementFailures += "Slide $createdCount replace failed: $shapeName" }
            }
            catch {
                $replacementFailures += "Slide $createdCount shape not found: $shapeName"
            }
        }
    }

    foreach ($sid in $allOriginalIds) {
        $s = Get-SlideById -Presentation $pres -SlideId ([int]$sid)
        if ($null -ne $s) { $s.Delete() }
    }

    $pres.Save()
}
finally {
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "PPT Design v2 파일 생성에 실패했습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT DESIGN v2 RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Create       : OK"
Write-Host "Slides       :" $createdCount
Write-Host "Replacements :" $replacementCount
Write-Host "Failures     :" $replacementFailures.Count
Write-Host "File         :" $outputPath
Write-Host "Size         :" (Get-Item $outputPath).Length

if ($replacementFailures.Count -gt 0) {
    Write-Host ""
    Write-Host "Replacement warnings:"
    $replacementFailures | ForEach-Object { Write-Host " - $_" }
}

Write-Host ""
Write-Host "STATUS: PPT_DESIGN_V2_COMPLETE"
