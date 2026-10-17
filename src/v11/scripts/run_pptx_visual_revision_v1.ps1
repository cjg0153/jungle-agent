param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$OriginalPath,

    [Parameter(Mandatory=$true)]
    [string]$VisualQaPath,

    [string]$UserRequest = "",

    [int]$ChunkSize = 6
)

$ErrorActionPreference = "Stop"

# ============================================================
# JUNGLE_PPT_NATIVE_VISUAL_REVISION_V1
#
# Claude:
#   Reads ORIGINAL + CURRENT + Visual QA comments
#   and creates targeted visual-revision operations.
#
# PowerPoint COM:
#   Applies those operations to a COPY of CURRENT.
#
# Safety:
#   - no slide creation/deletion
#   - no shape creation/deletion
#   - no text replacement
#   - Shape ID targeting only
#   - source/current PPT never overwritten
# ============================================================

$outputDir =
    "C:\AI\work\output"

$tempDir =
    "C:\AI\work\temp"

$claude =
    "C:\Users\Zephyrus G14\.local\bin\claude.exe"


# ============================================================
# 1. Validate
# ============================================================

foreach ($file in @(
    $SourcePath,
    $OriginalPath,
    $VisualQaPath,
    $claude
)) {

    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "Required file missing: $file"
    }
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


$sourceHashBefore =
    (Get-FileHash `
        -LiteralPath $SourcePath `
        -Algorithm SHA256).Hash

$originalHashBefore =
    (Get-FileHash `
        -LiteralPath $OriginalPath `
        -Algorithm SHA256).Hash


# ============================================================
# 2. QA
# ============================================================

$qa =
    [System.IO.File]::ReadAllText(
        $VisualQaPath
    ) |
    ConvertFrom-Json


if ($null -eq $qa.slides) {
    throw "Visual QA JSON has no slides."
}


$qaSlides =
    @($qa.slides)


$targetQa =
    @(
        $qaSlides |
        Where-Object {

            $severity =
                ([string]$_.severity).ToLowerInvariant()

            $better =
                [bool]$_.better_than_before


            ($severity -in @(
                "critical",
                "major",
                "minor"
            )) `
            -or (-not $better)
        } |
        Sort-Object {
            [int]$_.slide
        }
    )


if ($targetQa.Count -lt 1) {

    Write-Host ""
    Write-Host "No revision targets found."
    Write-Host "STATUS: PPT_VISUAL_REVISION_NOT_REQUIRED"
    exit 0
}


$targetSlideNumbers =
    @(
        $targetQa |
        ForEach-Object {
            [int]$_.slide
        } |
        Sort-Object -Unique
    )


Write-Host ""
Write-Host "========================================"
Write-Host " VISUAL REVISION TARGETS"
Write-Host "========================================"
Write-Host "QA average     :" $qa.average_score
Write-Host "Target slides  :" $targetSlideNumbers.Count
Write-Host "Slide numbers  :" ($targetSlideNumbers -join ", ")


# ============================================================
# 3. Paths
# ============================================================

$stamp =
    Get-Date -Format "yyyyMMdd_HHmmss"

$baseName =
    [System.IO.Path]::GetFileNameWithoutExtension(
        $SourcePath
    )


$outputPath =
    Join-Path `
        $outputDir `
        "${baseName}_visual_revision_v1_${stamp}.pptx"


$workDir =
    Join-Path `
        $tempDir `
        "ppt_visual_revision_v1_${stamp}"


$currentImageDir =
    Join-Path `
        $workDir `
        "current"


$originalImageDir =
    Join-Path `
        $workDir `
        "original"


New-Item `
    -ItemType Directory `
    -Path $workDir `
    -Force |
    Out-Null

New-Item `
    -ItemType Directory `
    -Path $currentImageDir `
    -Force |
    Out-Null

New-Item `
    -ItemType Directory `
    -Path $originalImageDir `
    -Force |
    Out-Null


Copy-Item `
    -LiteralPath $SourcePath `
    -Destination $outputPath `
    -Force


# ============================================================
# Helpers
# ============================================================

function Release-ComSafe {

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


function Has-Property {

    param(
        $Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    $p =
        $Object.PSObject.Properties[$Name]

    return (
        $null -ne $p `
        -and $null -ne $p.Value
    )
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


function Get-SafeText {

    param($Shape)

    try {

        if (
            $Shape.HasTextFrame -eq -1 `
            -and $Shape.TextFrame.HasText -eq -1
        ) {

            $t =
                [string]$Shape.TextFrame.TextRange.Text

            if ($t.Length -gt 1500) {

                return (
                    $t.Substring(0,1500) +
                    "...[TRUNCATED]"
                )
            }

            return $t
        }
    }
    catch {}

    return ""
}


function Get-FullTextSignature {

    param($Presentation)

    $items =
        @()

    for (
        $slideNo = 1;
        $slideNo -le $Presentation.Slides.Count;
        $slideNo++
    ) {

        $slide =
            $Presentation.Slides.Item($slideNo)

        foreach ($shape in $slide.Shapes) {

            try {

                if (
                    $shape.HasTextFrame -eq -1 `
                    -and $shape.TextFrame.HasText -eq -1
                ) {

                    $items +=
                        (
                            "{0}|{1}|{2}" `
                            -f `
                            $slideNo,
                            [int]$shape.Id,
                            [string]$shape.TextFrame.TextRange.Text
                        )
                }
            }
            catch {}


            try {

                if ($shape.HasTable -eq -1) {

                    for (
                        $r = 1;
                        $r -le $shape.Table.Rows.Count;
                        $r++
                    ) {

                        for (
                            $c = 1;
                            $c -le $shape.Table.Columns.Count;
                            $c++
                        ) {

                            $cellText =
                                [string]$shape.Table.Cell(
                                    $r,
                                    $c
                                ).Shape.TextFrame.TextRange.Text


                            $items +=
                                (
                                    "{0}|{1}|T{2},{3}|{4}" `
                                    -f `
                                    $slideNo,
                                    [int]$shape.Id,
                                    $r,
                                    $c,
                                    $cellText
                                )
                        }
                    }
                }
            }
            catch {}
        }
    }


    $joined =
        $items -join "`n"


    $bytes =
        [System.Text.Encoding]::UTF8.GetBytes(
            $joined
        )


    $sha =
        [System.Security.Cryptography.SHA256]::Create()

    try {

        return (
            [System.BitConverter]::ToString(
                $sha.ComputeHash($bytes)
            ).Replace("-","")
        )
    }
    finally {

        $sha.Dispose()
    }
}


function Export-SelectedSlides {

    param(
        [string]$FilePath,
        [string]$Destination,
        [int[]]$Slides
    )

    $ppt =
        $null

    $pres =
        $null


    try {

        $ppt =
            New-Object -ComObject PowerPoint.Application


        $pres =
            $ppt.Presentations.Open(
                $FilePath,
                $true,
                $false,
                $false
            )


        foreach ($slideNo in $Slides) {

            if (
                $slideNo -lt 1 `
                -or $slideNo -gt $pres.Slides.Count
            ) {

                throw "Invalid slide: $slideNo"
            }


            $pngPath =
                Join-Path `
                    $Destination `
                    ("slide_{0:D3}.png" -f $slideNo)


            $pres.Slides.Item(
                $slideNo
            ).Export(
                $pngPath,
                "PNG",
                1600,
                900
            )
        }


        return [int]$pres.Slides.Count
    }
    finally {

        if ($null -ne $pres) {

            try {
                $pres.Close()
            }
            catch {}

            Release-ComSafe $pres
        }


        if ($null -ne $ppt) {

            try {
                $ppt.Quit()
            }
            catch {}

            Release-ComSafe $ppt
        }


        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


# ============================================================
# 4. Render ORIGINAL and CURRENT
# ============================================================

Write-Host ""
Write-Host "----------------------------------------"
Write-Host " RENDER ORIGINAL TARGETS"
Write-Host "----------------------------------------"


$originalSlideCount =
    Export-SelectedSlides `
        -FilePath $OriginalPath `
        -Destination $originalImageDir `
        -Slides $targetSlideNumbers


Write-Host ""
Write-Host "----------------------------------------"
Write-Host " RENDER CURRENT TARGETS"
Write-Host "----------------------------------------"


$currentSlideCount =
    Export-SelectedSlides `
        -FilePath $SourcePath `
        -Destination $currentImageDir `
        -Slides $targetSlideNumbers


if ($originalSlideCount -ne $currentSlideCount) {
    throw "Slide count mismatch."
}


# ============================================================
# 5. Extract exact shape metadata from CURRENT
# ============================================================

$ppt =
    $null

$pres =
    $null

$slideWidth =
    0

$slideHeight =
    0

$metadataBySlide =
    @{}
# JUNGLE_PPT_SAFE_CAPABILITY_GUARD_V1
$safeCapabilityBySlide = @{}
$safeUnsupportedEvents = @()


try {

    $ppt =
        New-Object -ComObject PowerPoint.Application


    $pres =
        $ppt.Presentations.Open(
            $SourcePath,
            $true,
            $false,
            $false
        )


    $slideWidth =
        [double]$pres.PageSetup.SlideWidth

    $slideHeight =
        [double]$pres.PageSetup.SlideHeight


    foreach ($slideNo in $targetSlideNumbers) {

        $slide =
            $pres.Slides.Item(
                $slideNo
            )


        $shapes =
            @()


        foreach ($shape in $slide.Shapes) {

            $item =
                [ordered]@{
                    shape_id =
                        [int]$shape.Id

                    shape_name =
                        [string]$shape.Name

                    shape_type =
                        [int]$shape.Type

                    left =
                        [math]::Round(
                            [double]$shape.Left,
                            2
                        )

                    top =
                        [math]::Round(
                            [double]$shape.Top,
                            2
                        )

                    width =
                        [math]::Round(
                            [double]$shape.Width,
                            2
                        )

                    height =
                        [math]::Round(
                            [double]$shape.Height,
                            2
                        )

                    rotation =
                        [math]::Round(
                            [double]$shape.Rotation,
                            2
                        )

                    text =
                        Get-SafeText $shape

                    is_picture =
                        (
                            [int]$shape.Type -in @(
                                11,
                                13
                            )
                        )
                }


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


            if ($item.is_picture) {

                try {

                    $item.crop_left =
                        [math]::Round(
                            [double]$shape.PictureFormat.CropLeft,
                            2
                        )

                    $item.crop_right =
                        [math]::Round(
                            [double]$shape.PictureFormat.CropRight,
                            2
                        )

                    $item.crop_top =
                        [math]::Round(
                            [double]$shape.PictureFormat.CropTop,
                            2
                        )

                    $item.crop_bottom =
                        [math]::Round(
                            [double]$shape.PictureFormat.CropBottom,
                            2
                        )
                }
                catch {}
            }


            $shapes +=
                [pscustomobject]$item
        }


        $metadataBySlide[$slideNo] =
            $shapes

        $hasPictureShape = $false

        foreach ($metaShape in @($shapes)) {
            try {
                $pictureProperty =
                    $metaShape.PSObject.Properties["is_picture"]

                if (
                    $null -ne $pictureProperty `
                    -and [bool]$metaShape.is_picture
                ) {
                    $hasPictureShape = $true
                    break
                }
            }
            catch {
            }
        }

        $backgroundFillType = $null

        try {
            $backgroundFillType =
                [int]$slide.Background.Fill.Type
        }
        catch {
        }

        $backgroundPictureOnly =
            (
                (-not $hasPictureShape) `
                -and ($backgroundFillType -eq 6)
            )

        $safeCapabilityBySlide[[int]$slideNo] =
            [pscustomobject]@{
                has_picture_shape =
                    [bool]$hasPictureShape
                background_fill_type =
                    $backgroundFillType
                background_picture_only =
                    [bool]$backgroundPictureOnly
            }
    }
}
finally {

    if ($null -ne $pres) {

        try {
            $pres.Close()
        }
        catch {}

        Release-ComSafe $pres
    }


    if ($null -ne $ppt) {

        try {
            $ppt.Quit()
        }
        catch {}

        Release-ComSafe $ppt
    }


    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ============================================================
# 6. Claude revision planning
# ============================================================

$utf8Bom =
    New-Object System.Text.UTF8Encoding($true)


$allOperations =
    @()


$chunkIndex =
    0


for (
    $offset = 0;
    $offset -lt $targetSlideNumbers.Count;
    $offset += $ChunkSize
) {

    $chunkIndex++


    $endIndex =
        [Math]::Min(
            $targetSlideNumbers.Count - 1,
            $offset + $ChunkSize - 1
        )


    $chunkSlides =
        @(
            $targetSlideNumbers[
                $offset..$endIndex
            ]
        )


    $chunkData =
        @()


    foreach ($slideNo in $chunkSlides) {

        $qaItem =
            $targetQa |
            Where-Object {
                [int]$_.slide -eq $slideNo
            } |
            Select-Object -First 1


        $chunkData +=
            [pscustomobject]@{
                slide_number =
                    $slideNo

                better_than_before =
                    [bool]$qaItem.better_than_before

                severity =
                    [string]$qaItem.severity

                qa_comment =
                    [string]$qaItem.comment

                original_image =
                    (
                        Join-Path `
                            $originalImageDir `
                            ("slide_{0:D3}.png" -f $slideNo)
                    )

                current_image =
                    (
                        Join-Path `
                            $currentImageDir `
                            ("slide_{0:D3}.png" -f $slideNo)
                    )

                shapes =
                    $metadataBySlide[$slideNo]
            }
    }


    # JUNGLE_PPT_REVISION_CHECKPOINT_HARDENING_V1
    $checkpointVersion =
        "JUNGLE_PPT_REVISION_CHECKPOINT_V1"

    $promptVersion =
        "JUNGLE_PPT_VISUAL_REVISION_PROMPT_V1"


    function Get-JungleCheckpointSha256Text {

        param(
            [string]$Text
        )


        $bytes =
            [System.Text.Encoding]::UTF8.GetBytes(
                $Text
            )


        $sha =
            [System.Security.Cryptography.SHA256]::Create()


        try {

            return (
                [System.BitConverter]::ToString(
                    $sha.ComputeHash($bytes)
                ).Replace(
                    "-",
                    ""
                )
            )
        }
        finally {

            $sha.Dispose()
        }
    }


    function ConvertTo-JungleCheckpointCanonicalValue {

        param(
            [object]$Value
        )


        if ($null -eq $Value) {
            return $null
        }


        if ($Value -is [string]) {

            $stringValue =
                [string]$Value


            if ($stringValue -match '(?i)\.png$') {

                if (-not (
                    Test-Path `
                        -LiteralPath $stringValue `
                        -PathType Leaf
                )) {

                    throw (
                        "Checkpoint PNG input missing: " +
                        $stringValue
                    )
                }


                $leaf =
                    [System.IO.Path]::GetFileName(
                        $stringValue
                    )


                $fileHash =
                    (Get-FileHash `
                        -LiteralPath $stringValue `
                        -Algorithm SHA256).Hash


                return (
                    "PNG_SHA256:{0}:{1}" -f
                    $leaf,
                    $fileHash
                )
            }


            return $stringValue
        }


        if (
            $Value -is
            [System.Collections.IDictionary]
        ) {

            $orderedDictionary =
                [ordered]@{}


            $keys =
                @(
                    $Value.Keys |
                    ForEach-Object {
                        [string]$_
                    } |
                    Sort-Object
                )


            foreach ($key in $keys) {

                $orderedDictionary[$key] =
                    ConvertTo-JungleCheckpointCanonicalValue `
                        -Value $Value[$key]
            }


            return [pscustomobject]$orderedDictionary
        }


        if (
            $Value -is
            [System.Management.Automation.PSCustomObject]
        ) {

            $orderedObject =
                [ordered]@{}


            $propertyNames =
                @(
                    $Value.PSObject.Properties |
                    ForEach-Object {
                        $_.Name
                    } |
                    Sort-Object
                )


            foreach ($propertyName in $propertyNames) {

                $propertyValue =
                    $Value.PSObject.Properties[
                        $propertyName
                    ].Value


                $orderedObject[$propertyName] =
                    ConvertTo-JungleCheckpointCanonicalValue `
                        -Value $propertyValue
            }


            return [pscustomobject]$orderedObject
        }


        if (
            $Value -is
            [System.Collections.IEnumerable]
        ) {

            $canonicalArray =
                @()


            foreach ($item in $Value) {

                $canonicalItem =
                    ConvertTo-JungleCheckpointCanonicalValue `
                        -Value $item


                $canonicalArray +=
                    ,$canonicalItem
            }


            return ,$canonicalArray
        }


        return $Value
    }


    $canonicalChunkData =
        ConvertTo-JungleCheckpointCanonicalValue `
            -Value @($chunkData)


    $checkpointInputObject =
        [ordered]@{

            schema =
                "JUNGLE_PPT_REVISION_CHUNK_INPUT_V1"

            user_request =
                [string]$UserRequest

            chunk_slides =
                @($chunkSlides)

            chunk_data =
                $canonicalChunkData
        }


    $checkpointInputJson =
        $checkpointInputObject |
        ConvertTo-Json `
            -Depth 30 `
            -Compress


    $chunkInputHash =
        Get-JungleCheckpointSha256Text `
            -Text $checkpointInputJson

    $chunkDataPath =
        Join-Path `
            $workDir `
            ("revision_chunk_{0:D2}_data.json" -f $chunkIndex)


    $planPath =
        Join-Path `
            $workDir `
            ("revision_chunk_{0:D2}_plan.json" -f $chunkIndex)


    $instructionPath =
        Join-Path `
            $workDir `
            ("revision_chunk_{0:D2}_instruction.txt" -f $chunkIndex)


    [System.IO.File]::WriteAllText(
        $chunkDataPath,
        (
            $chunkData |
            ConvertTo-Json -Depth 15
        ),
        $utf8Bom
    )


    $instruction = @"
You are Jungle's senior PowerPoint visual revision specialist.

This is SECOND-PASS REVISION.

The first automated visual pass was technically successful,
but visual QA found slides that still need improvement.

USER REQUEST:
$UserRequest


READ THIS JSON:
$chunkDataPath


For every item in that JSON you have:

- ORIGINAL image
- CURRENT image after first visual pass
- visual QA severity
- visual QA comment
- exact CURRENT PowerPoint Shape IDs and geometry/style metadata


GOAL:

Fix the visual QA issues while preserving everything that is already good.

The CURRENT slide is the slide you are revising.
The ORIGINAL slide is only a comparison reference.


IMPORTANT:

1. Do not change slide count.
2. Do not create slides.
3. Do not delete slides.
4. Do not add shapes.
5. Do not delete shapes.
6. Do not change any substantive text.
7. Do not rewrite wording.
8. Do not remove screenshots, pictures, diagrams, tables, or media.
9. Use only existing Shape IDs.
10. Fix the exact QA complaint first.
11. Do not make cosmetic changes unrelated to the QA issue.
12. Avoid regressions on areas already improved.
13. Keep all objects inside slide bounds.
14. Preserve picture aspect ratios when resizing.
15. For cropped screenshots, crop values may be adjusted when necessary.
16. If QA says the CURRENT slide is identical or not better than ORIGINAL,
    make a meaningful but restrained improvement.
17. For minor-only issues, use minimal changes.
18. For major issues, make enough change to resolve the issue fully.

SLIDE SIZE:
width=$slideWidth
height=$slideHeight


Allowed operation properties:

- left
- top
- width
- height
- font_size
- font_bold
- font_rgb
- alignment
- fill_rgb
- fill_transparency
- line_visible
- line_rgb
- line_weight
- margin_left
- margin_right
- margin_top
- margin_bottom
- crop_left
- crop_right
- crop_top
- crop_bottom


PICTURES:

If is_picture=true:

- Preserve aspect ratio whenever resizing.
- Prefer changing only width OR height rather than both.
- Use crop properties only if crop is part of the problem.
- Do not distort screenshots or charts.


SAVE VALID UTF-8 JSON TO:

$planPath


SCHEMA:

{
  "operations": [
    {
      "slide_number": 27,
      "shape_id": 123,
      "left": 100.0,
      "top": 120.0,
      "width": 500.0,
      "reason": "QA issue being fixed"
    }
  ]
}


Rules:

- slide_number and shape_id are mandatory.
- reason is mandatory.
- Include only properties that actually need changing.
- Never invent a Shape ID.
- Never output text replacement.
- Never output slide creation/deletion.
- It is acceptable to emit multiple operations for one slide.
- Do not output operations for a slide if no safe improvement can be made.

After writing the JSON output only:

PPT_VISUAL_REVISION_PLAN_CREATED
"@


    [System.IO.File]::WriteAllText(
        $instructionPath,
        $instruction,
        $utf8Bom
    )


    Write-Host ""
    Write-Host "----------------------------------------"
    Write-Host (
        " CLAUDE VISUAL REVISION {0} | Slides {1}" `
        -f `
        $chunkIndex,
        ($chunkSlides -join ",")
    )
    Write-Host "----------------------------------------"


    $claudeRequest =
        "Read the instruction file at $instructionPath and execute it exactly. Read every local ORIGINAL and CURRENT PNG listed in the JSON and use the QA comments. Do not use web search, Google Drive, connectors, or external sources. Write the requested revision JSON. When finished output only PPT_VISUAL_REVISION_PLAN_CREATED."


    # JUNGLE_PPT_VISUAL_REVISION_RESUME_V1
    $resumeHit =
        $false

    $resumeRoot =
        $env:JUNGLE_VISUAL_REVISION_RESUME_DIR


    $planLeaf =
        [System.IO.Path]::GetFileName(
            $planPath
        )


    $checkpointPath =
        Join-Path `
            $workDir `
            ("revision_chunk_{0:D2}_checkpoint.json" -f $chunkIndex)


    $checkpointLeaf =
        [System.IO.Path]::GetFileName(
            $checkpointPath
        )


    $chunkLabel =
        "UNKNOWN"


    if (
        $planLeaf -match
        '^revision_chunk_(\d+)_plan\.json$'
    ) {

        $chunkLabel =
            [int]$Matches[1]
    }


    if (
        -not [string]::IsNullOrWhiteSpace(
            $resumeRoot
        )
    ) {

        if (
            Test-Path `
                -LiteralPath $resumeRoot `
                -PathType Container
        ) {

            $resumeCandidate =
                Join-Path `
                    $resumeRoot `
                    $planLeaf


            $resumeCheckpointCandidate =
                Join-Path `
                    $resumeRoot `
                    $checkpointLeaf


            if (
                Test-Path `
                    -LiteralPath $resumeCandidate `
                    -PathType Leaf
            ) {

                $resumeStatus =
                    $null

                $resumeCheckpoint =
                    $null


                if (-not (
                    Test-Path `
                        -LiteralPath $resumeCheckpointCandidate `
                        -PathType Leaf
                )) {

                    $resumeStatus =
                        "STALE_METADATA"
                }
                else {

                    try {

                        $resumeCheckpoint =
                            [System.IO.File]::ReadAllText(
                                $resumeCheckpointCandidate
                            ) |
                            ConvertFrom-Json
                    }
                    catch {

                        $resumeStatus =
                            "STALE_METADATA"
                    }
                }


                if ($null -eq $resumeStatus) {

                    if (
                        [string]$resumeCheckpoint.checkpoint_version `
                        -ne `
                        [string]$checkpointVersion
                    ) {

                        $resumeStatus =
                            "STALE_METADATA"
                    }
                    elseif (
                        [string]$resumeCheckpoint.prompt_version `
                        -ne `
                        [string]$promptVersion
                    ) {

                        $resumeStatus =
                            "STALE_PROMPT"
                    }
                    elseif (
                        [string]$resumeCheckpoint.chunk_input_hash `
                        -ne `
                        [string]$chunkInputHash
                    ) {

                        $resumeStatus =
                            "STALE_INPUT"
                    }
                    else {

                        $resumePlanHash =
                            (Get-FileHash `
                                -LiteralPath $resumeCandidate `
                                -Algorithm SHA256).Hash


                        if (
                            [string]$resumeCheckpoint.plan_sha256 `
                            -ne `
                            [string]$resumePlanHash
                        ) {

                            $resumeStatus =
                                "STALE_PLAN"
                        }
                        else {

                            $resumeStatus =
                                "REUSE"
                        }
                    }
                }


                if ($resumeStatus -eq "REUSE") {

                    Copy-Item `
                        -LiteralPath $resumeCandidate `
                        -Destination $planPath `
                        -Force


                    Copy-Item `
                        -LiteralPath $resumeCheckpointCandidate `
                        -Destination $checkpointPath `
                        -Force


                    $resumeHit =
                        $true
                }


                Write-Host (
                    "[REVISION RESUME] Chunk {0:D2} -> {1}" -f
                    $chunkLabel,
                    $resumeStatus
                )
            }
        }
    }

    $guardUnsupportedSlides = @()

    foreach ($guardSlideNo in @($chunkSlides)) {

        $guardCap =
            $safeCapabilityBySlide[[int]$guardSlideNo]

        if ($null -eq $guardCap) {
            continue
        }

        if (-not [bool]$guardCap.background_picture_only) {
            continue
        }

        $guardQa =
            @(
                $qaSlides |
                Where-Object {
                    [int]$_.slide -eq [int]$guardSlideNo
                } |
                Select-Object -First 1
            )

        $guardComment = ""

        if ($guardQa.Count -gt 0) {
            $guardComment =
                [string]$guardQa[0].comment
        }

        # JUNGLE_PPT_SAFE_CAPABILITY_GUARD_ASCII_REGEX_FIX_V1
        $guardNeedsPictureEdit =
            $guardComment -match '(?i)(screenshot|picture|image|chart|crop|panel|background|visual|photo|graphic)'

        if ($guardNeedsPictureEdit) {
            $guardUnsupportedSlides +=
                [int]$guardSlideNo
        }
    }

    $guardUnsupportedSlides =
        @(
            $guardUnsupportedSlides |
            Sort-Object -Unique
        )

    if ($guardUnsupportedSlides.Count -gt 0) {

        $safeUnsupportedEvents +=
            [pscustomobject]@{
                chunk =
                    [int]$chunkIndex
                slides =
                    @($guardUnsupportedSlides)
                reason =
                    "SAFE_UNSUPPORTED_BACKGROUND_IMAGE"
            }

        $unsupportedText =
            ($guardUnsupportedSlides -join ",")

        if (
            $guardUnsupportedSlides.Count `
            -eq @($chunkSlides).Count
        ) {

            $guardPlan =
                [ordered]@{
                    operations = @()
                }

            $guardJson =
                $guardPlan |
                ConvertTo-Json -Depth 4

            [System.IO.File]::WriteAllText(
                $planPath,
                $guardJson,
                $utf8Bom
            )

            $resumeHit = $true

            Write-Host (
                "[SAFE CAPABILITY GUARD] Slides {0} -> SAFE_UNSUPPORTED_BACKGROUND_IMAGE -> CLAUDE SKIPPED" -f $unsupportedText
            ) -ForegroundColor Yellow

        }
        else {

            $claudeRequest +=
                " SAFE_CAPABILITY_GUARD: Slides $unsupportedText use a picture background without editable picture Shape IDs and their QA issue requires screenshot/image/chart/crop/panel editing. Do not emit operations for those slides."

            Write-Host (
                "[SAFE CAPABILITY GUARD] Slides {0} -> SAFE_UNSUPPORTED_BACKGROUND_IMAGE -> excluded from operations" -f $unsupportedText
            ) -ForegroundColor Yellow
        }
    }

    if (-not $resumeHit) {

        Write-Host (
            "[REVISION RESUME] Chunk {0:D2} -> CLAUDE" -f $chunkLabel
        )

        & $claude `
            -p $claudeRequest `
            --allowedTools "Read" "Write"
    }


    if (-not (Test-Path -LiteralPath $planPath -PathType Leaf)) {

        throw "Revision plan missing: $planPath"
    }


    $plan =
        [System.IO.File]::ReadAllText(
            $planPath
        ) |
        ConvertFrom-Json


    if ($null -eq $plan.operations) {

        throw "Revision plan has no operations property."
    }


    $currentPlanHash =
        (Get-FileHash `
            -LiteralPath $planPath `
            -Algorithm SHA256).Hash


    $checkpointRecord =
        [ordered]@{

            checkpoint_version =
                $checkpointVersion

            prompt_version =
                $promptVersion

            chunk_input_hash =
                $chunkInputHash

            plan_sha256 =
                $currentPlanHash
        }


    $checkpointJson =
        $checkpointRecord |
        ConvertTo-Json `
            -Depth 6


    [System.IO.File]::WriteAllText(
        $checkpointPath,
        $checkpointJson,
        $utf8Bom
    )

    foreach ($op in @($plan.operations)) {

        $slideNo =
            [int]$op.slide_number


        if ($slideNo -notin $targetSlideNumbers) {

            throw (
                "Claude attempted revision outside target slides: " +
                $slideNo
            )
        }


        $validIds =
            @(
                $metadataBySlide[$slideNo] |
                ForEach-Object {
                    [int]$_.shape_id
                }
            )


        if ([int]$op.shape_id -notin $validIds) {

            throw (
                "Claude invented Shape ID: " +
                "slide=$slideNo shape=$($op.shape_id)"
            )
        }


        $allOperations +=
            $op
    }


    Write-Host "Operations :" @($plan.operations).Count
}


Write-Host ""
Write-Host "TOTAL REVISION OPERATIONS :" $allOperations.Count


# ============================================================
# 7. Apply revision
# ============================================================

$ppt =
    $null

$pres =
    $null

$applied =
    0

$failed =
    @()

$textSignatureBefore =
    $null


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


    if ([int]$pres.Slides.Count -ne $currentSlideCount) {
        throw "Slide count changed before revision apply."
    }


    $textSignatureBefore =
        Get-FullTextSignature $pres


    foreach ($op in $allOperations) {

        try {

            $slideNo =
                [int]$op.slide_number

            $shapeId =
                [int]$op.shape_id


            $slide =
                $pres.Slides.Item(
                    $slideNo
                )


            $shape =
                Get-ShapeById `
                    -Slide $slide `
                    -ShapeId $shapeId


            if ($null -eq $shape) {
                throw "Shape ID not found."
            }


            # =================================================
            # Geometry
            # =================================================

            $isPicture =
                ([int]$shape.Type -in @(11,13))


            $newLeft =
                [double]$shape.Left

            $newTop =
                [double]$shape.Top

            $newWidth =
                [double]$shape.Width

            $newHeight =
                [double]$shape.Height


            if (Has-Property $op "left") {
                $newLeft = [double]$op.left
            }

            if (Has-Property $op "top") {
                $newTop = [double]$op.top
            }

            if (Has-Property $op "width") {
                $newWidth = [double]$op.width
            }

            if (Has-Property $op "height") {
                $newHeight = [double]$op.height
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


            if ($isPicture) {

                $hasW =
                    Has-Property $op "width"

                $hasH =
                    Has-Property $op "height"


                if ($hasW -xor $hasH) {

                    try {
                        $shape.LockAspectRatio = -1
                    }
                    catch {}
                }
            }


            if (Has-Property $op "left") {
                $shape.Left = $newLeft
            }

            if (Has-Property $op "top") {
                $shape.Top = $newTop
            }

            if (Has-Property $op "width") {
                $shape.Width = $newWidth
            }

            if (Has-Property $op "height") {
                $shape.Height = $newHeight
            }


            # =================================================
            # Text style only — NEVER text content
            # =================================================

            if (
                $shape.HasTextFrame -eq -1 `
                -and $shape.TextFrame.HasText -eq -1
            ) {

                $tr =
                    $shape.TextFrame.TextRange


                if (Has-Property $op "font_size") {

                    $value =
                        [double]$op.font_size


                    if (
                        $value -ge 7 `
                        -and $value -le 72
                    ) {

                        $tr.Font.Size =
                            $value
                    }
                }


                if (Has-Property $op "font_bold") {

                    if ([bool]$op.font_bold) {

                        $tr.Font.Bold =
                            -1
                    }
                    else {

                        $tr.Font.Bold =
                            0
                    }
                }


                if (Has-Property $op "font_rgb") {

                    $tr.Font.Color.RGB =
                        [int]$op.font_rgb
                }


                if (Has-Property $op "alignment") {

                    $value =
                        [int]$op.alignment


                    if ($value -ge 1 -and $value -le 4) {

                        $tr.ParagraphFormat.Alignment =
                            $value
                    }
                }


                foreach ($marginName in @(
                    "margin_left",
                    "margin_right",
                    "margin_top",
                    "margin_bottom"
                )) {

                    if (-not (Has-Property $op $marginName)) {
                        continue
                    }


                    $value =
                        [double]$op.$marginName


                    switch ($marginName) {

                        "margin_left" {
                            $shape.TextFrame.MarginLeft = $value
                        }

                        "margin_right" {
                            $shape.TextFrame.MarginRight = $value
                        }

                        "margin_top" {
                            $shape.TextFrame.MarginTop = $value
                        }

                        "margin_bottom" {
                            $shape.TextFrame.MarginBottom = $value
                        }
                    }
                }
            }


            # =================================================
            # Fill
            # =================================================

            if (Has-Property $op "fill_rgb") {

                try {

                    $shape.Fill.Visible = -1
                    $shape.Fill.Solid()

                    $shape.Fill.ForeColor.RGB =
                        [int]$op.fill_rgb
                }
                catch {}
            }


            if (Has-Property $op "fill_transparency") {

                try {

                    $value =
                        [double]$op.fill_transparency


                    if ($value -ge 0 -and $value -le 1) {

                        $shape.Fill.Transparency =
                            $value
                    }
                }
                catch {}
            }


            # =================================================
            # Line
            # =================================================

            if (Has-Property $op "line_visible") {

                if ([bool]$op.line_visible) {

                    $shape.Line.Visible =
                        -1
                }
                else {

                    $shape.Line.Visible =
                        0
                }
            }


            if (Has-Property $op "line_rgb") {

                try {

                    $shape.Line.ForeColor.RGB =
                        [int]$op.line_rgb
                }
                catch {}
            }


            if (Has-Property $op "line_weight") {

                try {

                    $value =
                        [double]$op.line_weight


                    if (
                        $value -ge 0.25 `
                        -and $value -le 10
                    ) {

                        $shape.Line.Weight =
                            $value
                    }
                }
                catch {}
            }


            # =================================================
            # Picture crop
            # =================================================

            if ($isPicture) {

                foreach ($cropName in @(
                    "crop_left",
                    "crop_right",
                    "crop_top",
                    "crop_bottom"
                )) {

                    if (-not (Has-Property $op $cropName)) {
                        continue
                    }


                    $value =
                        [double]$op.$cropName


                    switch ($cropName) {

                        "crop_left" {
                            $shape.PictureFormat.CropLeft = $value
                        }

                        "crop_right" {
                            $shape.PictureFormat.CropRight = $value
                        }

                        "crop_top" {
                            $shape.PictureFormat.CropTop = $value
                        }

                        "crop_bottom" {
                            $shape.PictureFormat.CropBottom = $value
                        }
                    }
                }
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


    # ========================================================
    # Text integrity gate
    # ========================================================

    $textSignatureAfter =
        Get-FullTextSignature $pres


    if ($textSignatureBefore -ne $textSignatureAfter) {

        throw (
            "SAFETY FAILURE: substantive PPT text changed " +
            "during visual revision."
        )
    }


    if ([int]$pres.Slides.Count -ne $currentSlideCount) {

        throw "SAFETY FAILURE: slide count changed."
    }


    $pres.Save()
}
finally {

    if ($null -ne $pres) {

        try {
            $pres.Close()
        }
        catch {}

        Release-ComSafe $pres
    }


    if ($null -ne $ppt) {

        try {
            $ppt.Quit()
        }
        catch {}

        Release-ComSafe $ppt
    }


    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ============================================================
# 8. Post safety
# ============================================================

$sourceHashAfter =
    (Get-FileHash `
        -LiteralPath $SourcePath `
        -Algorithm SHA256).Hash


$originalHashAfter =
    (Get-FileHash `
        -LiteralPath $OriginalPath `
        -Algorithm SHA256).Hash


if ($sourceHashBefore -ne $sourceHashAfter) {

    throw "SAFETY FAILURE: Source visual PPT was modified."
}


if ($originalHashBefore -ne $originalHashAfter) {

    throw "SAFETY FAILURE: Original PPT was modified."
}


if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {

    throw "Revision PPTX was not created."
}


$outputInfo =
    Get-Item `
        -LiteralPath $outputPath


if ($outputInfo.Length -lt 10000) {

    throw "Revision PPTX is unexpectedly small."
}


# ============================================================
# 9. Result
# ============================================================

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT NATIVE VISUAL REVISION RESULT"
Write-Host "========================================"
Write-Host ""

Write-Host "Revision       : OK"
Write-Host "Slides total   :" $currentSlideCount
Write-Host "Target slides  :" $targetSlideNumbers.Count
Write-Host "Operations     :" $allOperations.Count
Write-Host "Applied        :" $applied
Write-Host "Failed         :" $failed.Count
Write-Host "Text integrity : PASSED"
Write-Host "Source hash    : UNCHANGED"
Write-Host "Original hash  : UNCHANGED"
Write-Host "File           :" $outputPath
Write-Host "Size           :" $outputInfo.Length
Write-Host "Work Dir       :" $workDir


if ($failed.Count -gt 0) {

    Write-Host ""
    Write-Host "Failed operations:"

    foreach ($item in $failed) {

        Write-Host " -" $item
    }
}


if ($failed.Count -eq 0) {

    Write-Host ""
    Write-Host "STATUS: PPT_NATIVE_VISUAL_REVISION_COMPLETE"
}
else {

    Write-Host ""
    Write-Host "STATUS: PPT_NATIVE_VISUAL_REVISION_PARTIAL"
}