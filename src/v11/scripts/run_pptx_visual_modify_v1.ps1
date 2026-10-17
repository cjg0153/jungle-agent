param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest,

    [int]$SlideLimit = 0,

    [int]$ChunkSize = 8
)

$ErrorActionPreference = "Stop"

$outputDir =
    "C:\AI\work\output"

$tempDir =
    "C:\AI\work\temp"

$claudePath =
    "C:\Users\Zephyrus G14\.local\bin\claude.exe"


# ============================================================
# JUNGLE_PPT_VISUAL_EXECUTOR_V1
#
# Claude = Visual Designer
# PowerPoint COM = Executor
#
# IMPORTANT:
# - Existing slides are preserved.
# - Existing text/content is NOT rewritten.
# - Existing images/media are NOT deleted.
# - Existing shapes are NOT deleted.
# - Claude can only change geometry/style of existing shapes.
# ============================================================


# ============================================================
# 1. Validate
# ============================================================

if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "PPT file not found: $FilePath"
}

if (-not (Test-Path -LiteralPath $claudePath -PathType Leaf)) {
    throw "Claude executable not found: $claudePath"
}

$ext =
    [System.IO.Path]::GetExtension(
        $FilePath
    ).ToLowerInvariant()

if ($ext -notin @(".ppt",".pptx")) {
    throw "Only PPT/PPTX is supported."
}

if ($ChunkSize -lt 1) {
    throw "ChunkSize must be >= 1."
}

foreach ($dir in @(
    $outputDir,
    $tempDir
)) {

    if (-not (Test-Path -LiteralPath $dir)) {

        New-Item `
            -ItemType Directory `
            -Path $dir `
            -Force |
            Out-Null
    }
}


# ============================================================
# 2. Paths
# ============================================================

$stamp =
    Get-Date -Format "yyyyMMdd_HHmmss"

$baseName =
    [System.IO.Path]::GetFileNameWithoutExtension(
        $FilePath
    )

$outputPath =
    Join-Path `
        $outputDir `
        "${baseName}_visual_v1_${stamp}.pptx"

$workDir =
    Join-Path `
        $tempDir `
        "ppt_visual_v1_${stamp}"

New-Item `
    -ItemType Directory `
    -Path $workDir `
    -Force |
    Out-Null

$imageDir =
    Join-Path `
        $workDir `
        "slides"

New-Item `
    -ItemType Directory `
    -Path $imageDir `
    -Force |
    Out-Null

$metadataPath =
    Join-Path `
        $workDir `
        "shape_metadata.json"

$combinedPlanPath =
    Join-Path `
        $workDir `
        "combined_visual_plan.json"


# ============================================================
# 3. Copy source
# ============================================================

Copy-Item `
    -LiteralPath $FilePath `
    -Destination $outputPath `
    -Force


# ============================================================
# Helpers
# ============================================================

function Release-ComObjectSafe {

    param($Object)

    if ($null -ne $Object) {

        try {

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Object
            ) | Out-Null
        }
        catch {}
    }
}


function Get-SafeShapeText {

    param($Shape)

    try {

        if (
            $Shape.HasTextFrame -eq -1 `
            -and $Shape.TextFrame.HasText -eq -1
        ) {

            $t =
                [string]$Shape.TextFrame.TextRange.Text

            if ($t.Length -gt 1500) {

                $t =
                    $t.Substring(0,1500) +
                    "...[TRUNCATED]"
            }

            return $t
        }
    }
    catch {}

    return ""
}


function Get-ShapeById {

    param(
        $Slide,
        [int]$ShapeId
    )

    foreach ($shape in $Slide.Shapes) {

        try {

            if ([int]$shape.Id -eq $ShapeId) {
                return $shape
            }
        }
        catch {}
    }

    return $null
}


function Has-PlanProperty {

    param(
        $Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    $p =
        $Object.PSObject.Properties[$Name]

    if ($null -eq $p) {
        return $false
    }

    return $null -ne $p.Value
}


# ============================================================
# 4. Extract slide images + shape geometry
# ============================================================

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE VISUAL ANALYSIS"
Write-Host "========================================"

$ppt =
    $null

$pres =
    $null

$slidesMeta =
    @()

try {

    $ppt =
        New-Object -ComObject PowerPoint.Application

    $pres =
        $ppt.Presentations.Open(
            $outputPath,
            $true,
            $false,
            $false
        )

    $slideCount =
        [int]$pres.Slides.Count

    $slideWidth =
        [double]$pres.PageSetup.SlideWidth

    $slideHeight =
        [double]$pres.PageSetup.SlideHeight

    if (
        $SlideLimit -gt 0 `
        -and $SlideLimit -lt $slideCount
    ) {

        $analysisCount =
            $SlideLimit
    }
    else {

        $analysisCount =
            $slideCount
    }

    Write-Host "Slides total :" $slideCount
    Write-Host "Slides target:" $analysisCount
    Write-Host "Slide size   :" $slideWidth "x" $slideHeight


    for ($s = 1; $s -le $analysisCount; $s++) {

        $slide =
            $pres.Slides.Item($s)

        $pngPath =
            Join-Path `
                $imageDir `
                ("slide_{0:D3}.png" -f $s)

        $slide.Export(
            $pngPath,
            "PNG",
            1600,
            900
        )

        $shapeMeta =
            @()

        foreach ($shape in $slide.Shapes) {

            $item =
                [ordered]@{
                    shape_id   = [int]$shape.Id
                    shape_name = [string]$shape.Name
                    shape_type = [int]$shape.Type

                    left   = [math]::Round([double]$shape.Left,2)
                    top    = [math]::Round([double]$shape.Top,2)
                    width  = [math]::Round([double]$shape.Width,2)
                    height = [math]::Round([double]$shape.Height,2)

                    rotation =
                        [math]::Round(
                            [double]$shape.Rotation,
                            2
                        )

                    text =
                        Get-SafeShapeText `
                            -Shape $shape
                }


            # ------------------------------------------------
            # Text style
            # ------------------------------------------------

            try {

                if (
                    $shape.HasTextFrame -eq -1 `
                    -and $shape.TextFrame.HasText -eq -1
                ) {

                    $tr =
                        $shape.TextFrame.TextRange

                    $item.font_size =
                        [math]::Round(
                            [double]$tr.Font.Size,
                            2
                        )

                    $item.font_bold =
                        ([int]$tr.Font.Bold -eq -1)

                    $item.font_name =
                        [string]$tr.Font.Name

                    $item.font_rgb =
                        [int]$tr.Font.Color.RGB

                    $item.alignment =
                        [int]$tr.ParagraphFormat.Alignment

                    $item.margin_left =
                        [math]::Round(
                            [double]$shape.TextFrame.MarginLeft,
                            2
                        )

                    $item.margin_right =
                        [math]::Round(
                            [double]$shape.TextFrame.MarginRight,
                            2
                        )

                    $item.margin_top =
                        [math]::Round(
                            [double]$shape.TextFrame.MarginTop,
                            2
                        )

                    $item.margin_bottom =
                        [math]::Round(
                            [double]$shape.TextFrame.MarginBottom,
                            2
                        )
                }
            }
            catch {}


            # ------------------------------------------------
            # Fill
            # ------------------------------------------------

            try {

                $item.fill_visible =
                    ([int]$shape.Fill.Visible -eq -1)

                $item.fill_rgb =
                    [int]$shape.Fill.ForeColor.RGB

                $item.fill_transparency =
                    [math]::Round(
                        [double]$shape.Fill.Transparency,
                        3
                    )
            }
            catch {}


            # ------------------------------------------------
            # Line
            # ------------------------------------------------

            try {

                $item.line_visible =
                    ([int]$shape.Line.Visible -eq -1)

                $item.line_rgb =
                    [int]$shape.Line.ForeColor.RGB

                $item.line_weight =
                    [math]::Round(
                        [double]$shape.Line.Weight,
                        2
                    )
            }
            catch {}


            $shapeMeta +=
                [pscustomobject]$item
        }


        $slidesMeta +=
            [pscustomobject]@{
                slide_number = $s
                image_path   = $pngPath
                shapes       = $shapeMeta
            }

        Write-Host (
            "Exported slide {0}/{1} | shapes={2}" `
            -f `
            $s,
            $analysisCount,
            $shapeMeta.Count
        )
    }
}
finally {

    if ($null -ne $pres) {

        try {
            $pres.Close()
        }
        catch {}

        Release-ComObjectSafe $pres
    }

    if ($null -ne $ppt) {

        try {
            $ppt.Quit()
        }
        catch {}

        Release-ComObjectSafe $ppt
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


$metadata =
    [ordered]@{
        source_file  = $FilePath
        output_file  = $outputPath
        slide_count  = $slideCount
        target_count = $analysisCount
        slide_width  = $slideWidth
        slide_height = $slideHeight
        slides       = $slidesMeta
    }


$utf8Bom =
    New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllText(
    $metadataPath,
    (
        $metadata |
        ConvertTo-Json -Depth 12
    ),
    $utf8Bom
)


# ============================================================
# 5. Claude visual planning - chunked
# ============================================================

$allOperations =
    @()

$chunkIndex =
    0


for (
    $chunkStart = 1;
    $chunkStart -le $analysisCount;
    $chunkStart += $ChunkSize
) {

    $chunkIndex++

    $chunkEnd =
        [Math]::Min(
            $analysisCount,
            $chunkStart + $ChunkSize - 1
        )

    $chunkSlides =
        @(
            $slidesMeta |
            Where-Object {
                $_.slide_number -ge $chunkStart `
                -and $_.slide_number -le $chunkEnd
            }
        )

    $chunkMetaPath =
        Join-Path `
            $workDir `
            (
                "chunk_{0:D2}_metadata.json" `
                -f $chunkIndex
            )

    $chunkPlanPath =
        Join-Path `
            $workDir `
            (
                "chunk_{0:D2}_plan.json" `
                -f $chunkIndex
            )

    $instructionPath =
        Join-Path `
            $workDir `
            (
                "chunk_{0:D2}_instruction.txt" `
                -f $chunkIndex
            )


    $chunkMeta =
        [ordered]@{
            slide_width  = $slideWidth
            slide_height = $slideHeight
            slides       = $chunkSlides
        }


    [System.IO.File]::WriteAllText(
        $chunkMetaPath,
        (
            $chunkMeta |
            ConvertTo-Json -Depth 12
        ),
        $utf8Bom
    )


    $imageList =
        (
            $chunkSlides |
            ForEach-Object {
                "Slide $($_.slide_number): $($_.image_path)"
            }
        ) -join "`r`n"


    $instruction = @"
You are Jungle's PowerPoint visual designer.

Your job is to IMPROVE THE EXISTING SLIDES visually.

This is NOT a rewrite.
This is NOT a summary.
This is NOT a new short deck.

USER REQUEST:
$UserRequest


SLIDES:
$chunkStart through $chunkEnd

SLIDE SIZE:
width=$slideWidth
height=$slideHeight


READ THIS METADATA JSON:
$chunkMetaPath


READ THESE SLIDE PNG IMAGES:
$imageList


IMPORTANT SAFETY RULES:

1. Preserve every slide.
2. Preserve ALL substantive text.
3. Preserve all images, screenshots, diagrams, tables, videos and media.
4. Do NOT delete shapes.
5. Do NOT add shapes.
6. Do NOT change shape text.
7. Do NOT change slide count.
8. Do NOT redesign the content concept.
9. Work only with existing shape IDs.

Your role is visual refinement only.

Improve where useful:
- information hierarchy
- alignment
- spacing
- margins
- size balance
- visual grouping
- headline emphasis
- typography hierarchy
- readability
- color consistency
- box/fill styling
- line styling
- image sizing / positioning
- whitespace

Avoid unnecessary changes.
Do not move a shape simply because you can.
Preserve good existing design.

Use the PNGs as the primary visual reference.
Use the metadata for exact geometry and Shape IDs.


OUTPUT:

Save VALID UTF-8 JSON exactly to:

$chunkPlanPath


SCHEMA:

{
  "operations": [
    {
      "slide_number": 1,
      "shape_id": 5,

      "left": 100.0,
      "top": 120.0,
      "width": 500.0,
      "height": 80.0,

      "font_size": 28.0,
      "font_bold": true,
      "font_rgb": 16777215,
      "alignment": 1,

      "fill_rgb": 255,
      "fill_transparency": 0.0,

      "line_visible": false,
      "line_rgb": 0,
      "line_weight": 1.0,

      "margin_left": 8.0,
      "margin_right": 8.0,
      "margin_top": 4.0,
      "margin_bottom": 4.0,

      "reason": "short visual-design reason"
    }
  ]
}


RULES FOR JSON:

- Only include properties that actually need changing.
- slide_number and shape_id are always required.
- reason is always required.
- Never output a text replacement property.
- Never invent a Shape ID.
- Never move shapes outside the slide bounds.
- width and height must remain positive.
- Prefer restrained professional banking/executive presentation design.
- If a slide is already strong, it is valid to return no operations for it.


When complete output only:

PPT_VISUAL_PLAN_CREATED
"@


    [System.IO.File]::WriteAllText(
        $instructionPath,
        $instruction,
        $utf8Bom
    )


    Write-Host ""
    Write-Host "----------------------------------------"
    Write-Host (
        " CLAUDE VISUAL PLAN {0} | Slides {1}-{2}" `
        -f `
        $chunkIndex,
        $chunkStart,
        $chunkEnd
    )
    Write-Host "----------------------------------------"


    $claudeRequest =
        "Read the instruction file at $instructionPath and execute it exactly. Read the local PNG slide images and JSON metadata specified there. Do not use web search, Google Drive, connectors, or external sources. Write the requested JSON plan. When finished output only PPT_VISUAL_PLAN_CREATED."


    & $claudePath `
        -p $claudeRequest `
        --allowedTools "Read" "Write"


    if (-not (Test-Path -LiteralPath $chunkPlanPath -PathType Leaf)) {

        throw "Claude visual plan was not created: $chunkPlanPath"
    }


    $chunkPlan =
        Get-Content `
            -LiteralPath $chunkPlanPath `
            -Raw `
            -Encoding UTF8 |
        ConvertFrom-Json


    if ($null -eq $chunkPlan.operations) {

        throw "Visual plan does not contain operations: $chunkPlanPath"
    }


    foreach ($op in @($chunkPlan.operations)) {

        $allOperations +=
            $op
    }


    Write-Host (
        "Plan operations : {0}" `
        -f `
        @($chunkPlan.operations).Count
    )
}


$combinedPlan =
    [ordered]@{
        source_file =
            $FilePath

        output_file =
            $outputPath

        operations =
            $allOperations
    }


[System.IO.File]::WriteAllText(
    $combinedPlanPath,
    (
        $combinedPlan |
        ConvertTo-Json -Depth 12
    ),
    $utf8Bom
)


Write-Host ""
Write-Host "TOTAL VISUAL OPERATIONS :" $allOperations.Count


# ============================================================
# 6. Apply visual plan with PowerPoint COM
# ============================================================

$ppt =
    $null

$pres =
    $null

$applied =
    0

$failed =
    @()


try {

    $ppt =
        New-Object -ComObject PowerPoint.Application

    $pres =
        $ppt.Presentations.Open(
            $outputPath,
            $false,
            $false,
            $false
        )


    if ([int]$pres.Slides.Count -ne $slideCount) {

        throw "Slide count changed before apply."
    }


    foreach ($op in $allOperations) {

        try {

            $slideNumber =
                [int]$op.slide_number

            $shapeId =
                [int]$op.shape_id


            if (
                $slideNumber -lt 1 `
                -or $slideNumber -gt $analysisCount
            ) {

                throw "Invalid slide number"
            }


            $slide =
                $pres.Slides.Item(
                    $slideNumber
                )


            $shape =
                Get-ShapeById `
                    -Slide $slide `
                    -ShapeId $shapeId


            if ($null -eq $shape) {

                throw "Shape ID not found"
            }


            # ------------------------------------------------
            # Geometry
            # ------------------------------------------------

            $newLeft =
                [double]$shape.Left

            $newTop =
                [double]$shape.Top

            $newWidth =
                [double]$shape.Width

            $newHeight =
                [double]$shape.Height


            if (Has-PlanProperty $op "left") {

                $newLeft =
                    [double]$op.left
            }

            if (Has-PlanProperty $op "top") {

                $newTop =
                    [double]$op.top
            }

            if (Has-PlanProperty $op "width") {

                $newWidth =
                    [double]$op.width
            }

            if (Has-PlanProperty $op "height") {

                $newHeight =
                    [double]$op.height
            }


            $newWidth =
                [Math]::Max(
                    5,
                    [Math]::Min(
                        $newWidth,
                        $slideWidth
                    )
                )

            $newHeight =
                [Math]::Max(
                    5,
                    [Math]::Min(
                        $newHeight,
                        $slideHeight
                    )
                )

            $newLeft =
                [Math]::Max(
                    0,
                    [Math]::Min(
                        $newLeft,
                        $slideWidth - $newWidth
                    )
                )

            $newTop =
                [Math]::Max(
                    0,
                    [Math]::Min(
                        $newTop,
                        $slideHeight - $newHeight
                    )
                )


            if (Has-PlanProperty $op "left") {
                $shape.Left = $newLeft
            }

            if (Has-PlanProperty $op "top") {
                $shape.Top = $newTop
            }

            if (Has-PlanProperty $op "width") {
                $shape.Width = $newWidth
            }

            if (Has-PlanProperty $op "height") {
                $shape.Height = $newHeight
            }


            # ------------------------------------------------
            # Text style
            # ------------------------------------------------

            if (
                $shape.HasTextFrame -eq -1 `
                -and $shape.TextFrame.HasText -eq -1
            ) {

                $tr =
                    $shape.TextFrame.TextRange


                if (Has-PlanProperty $op "font_size") {

                    $fontSize =
                        [double]$op.font_size

                    if (
                        $fontSize -ge 7 `
                        -and $fontSize -le 72
                    ) {

                        $tr.Font.Size =
                            $fontSize
                    }
                }


                if (Has-PlanProperty $op "font_bold") {

                    if ([bool]$op.font_bold) {

                        $tr.Font.Bold =
                            -1
                    }
                    else {

                        $tr.Font.Bold =
                            0
                    }
                }


                if (Has-PlanProperty $op "font_rgb") {

                    $tr.Font.Color.RGB =
                        [int]$op.font_rgb
                }


                if (Has-PlanProperty $op "alignment") {

                    $alignment =
                        [int]$op.alignment

                    if (
                        $alignment -ge 1 `
                        -and $alignment -le 4
                    ) {

                        $tr.ParagraphFormat.Alignment =
                            $alignment
                    }
                }


                if (Has-PlanProperty $op "margin_left") {

                    $shape.TextFrame.MarginLeft =
                        [double]$op.margin_left
                }

                if (Has-PlanProperty $op "margin_right") {

                    $shape.TextFrame.MarginRight =
                        [double]$op.margin_right
                }

                if (Has-PlanProperty $op "margin_top") {

                    $shape.TextFrame.MarginTop =
                        [double]$op.margin_top
                }

                if (Has-PlanProperty $op "margin_bottom") {

                    $shape.TextFrame.MarginBottom =
                        [double]$op.margin_bottom
                }
            }


            # ------------------------------------------------
            # Fill style
            # ------------------------------------------------

            if (Has-PlanProperty $op "fill_rgb") {

                try {

                    $shape.Fill.Visible =
                        -1

                    $shape.Fill.Solid()

                    $shape.Fill.ForeColor.RGB =
                        [int]$op.fill_rgb

                    if (Has-PlanProperty $op "fill_transparency") {

                        $transparency =
                            [double]$op.fill_transparency

                        if (
                            $transparency -ge 0 `
                            -and $transparency -le 1
                        ) {

                            $shape.Fill.Transparency =
                                $transparency
                        }
                    }
                }
                catch {}
            }


            # ------------------------------------------------
            # Line style
            # ------------------------------------------------

            if (Has-PlanProperty $op "line_visible") {

                if ([bool]$op.line_visible) {

                    $shape.Line.Visible =
                        -1
                }
                else {

                    $shape.Line.Visible =
                        0
                }
            }


            if (Has-PlanProperty $op "line_rgb") {

                try {

                    $shape.Line.ForeColor.RGB =
                        [int]$op.line_rgb
                }
                catch {}
            }


            if (Has-PlanProperty $op "line_weight") {

                try {

                    $weight =
                        [double]$op.line_weight

                    if (
                        $weight -ge 0.25 `
                        -and $weight -le 10
                    ) {

                        $shape.Line.Weight =
                            $weight
                    }
                }
                catch {}
            }


            $applied++
        }
        catch {

            $failed +=
                (
                    "Slide {0} / Shape {1} : {2}" `
                    -f `
                    $op.slide_number,
                    $op.shape_id,
                    $_.Exception.Message
                )
        }
    }


    $pres.Save()
}
finally {

    if ($null -ne $pres) {

        try {
            $pres.Close()
        }
        catch {}

        Release-ComObjectSafe $pres
    }

    if ($null -ne $ppt) {

        try {
            $ppt.Quit()
        }
        catch {}

        Release-ComObjectSafe $ppt
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ============================================================
# 7. Validate output
# ============================================================

if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {

    throw "Visual PPTX was not created."
}


$outputInfo =
    Get-Item -LiteralPath $outputPath


if ($outputInfo.Length -lt 10000) {

    throw "Generated PPTX is unexpectedly small."
}


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT VISUAL MODIFY RESULT"
Write-Host "========================================"
Write-Host ""

Write-Host "Visual Modify : OK"
Write-Host "Slides total  :" $slideCount
Write-Host "Slides edited :" $analysisCount
Write-Host "Operations    :" $allOperations.Count
Write-Host "Applied       :" $applied
Write-Host "Failed        :" $failed.Count
Write-Host "Original      :" $FilePath
Write-Host "File          :" $outputPath
Write-Host "Size          :" $outputInfo.Length
Write-Host "Plan          :" $combinedPlanPath

if ($failed.Count -gt 0) {

    Write-Host ""
    Write-Host "Failed operations:"

    foreach ($item in $failed) {

        Write-Host " -" $item
    }
}


if ($failed.Count -eq 0) {

    Write-Host ""
    Write-Host "STATUS: PPT_VISUAL_MODIFY_COMPLETE"
}
else {

    Write-Host ""
    Write-Host "STATUS: PPT_VISUAL_MODIFY_PARTIAL"
}