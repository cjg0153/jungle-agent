param(
    [Parameter(Mandatory=$true)]
    [string]$BeforePath,

    [Parameter(Mandatory=$true)]
    [string]$AfterPath,

    [int]$ChunkSize = 8,

    [string]$OutputDir = "C:\AI\work\output",

    [string]$WorkRoot = "C:\AI\work\temp",

    [string]$ClaudePath =
        "C:\Users\Zephyrus G14\.local\bin\claude.exe",

    [string]$ResumeDir = ""
)

$ErrorActionPreference = "Stop"

# JUNGLE_PPT_VISUAL_REQA_V1


function Release-ComObject {

    param(
        [object]$Object
    )

    if ($null -ne $Object) {

        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $Object
            )
        }
        catch {
        }
    }
}


function Export-PptSlides {

    param(
        [string]$PptPath,
        [string]$TargetDir
    )

    if (-not (Test-Path -LiteralPath $PptPath -PathType Leaf)) {
        throw "PPT missing: $PptPath"
    }

    New-Item `
        -ItemType Directory `
        -Path $TargetDir `
        -Force |
        Out-Null

    $ppt = $null
    $pres = $null

    try {

        $ppt =
            New-Object `
                -ComObject PowerPoint.Application

        try {
            $ppt.DisplayAlerts = 1
        }
        catch {
        }

        $pres =
            $ppt.Presentations.Open(
                $PptPath,
                $true,
                $false,
                $false
            )

        $count =
            [int]$pres.Slides.Count

        for ($i = 1; $i -le $count; $i++) {

            $slide =
                $pres.Slides.Item($i)

            $png =
                Join-Path `
                    $TargetDir `
                    ("slide_{0:D3}.png" -f $i)

            $slide.Export(
                $png,
                "PNG",
                1600,
                900
            )

            Release-ComObject $slide
        }

        return $count
    }
    finally {

        if ($null -ne $pres) {

            try {
                $pres.Close()
            }
            catch {
            }

            Release-ComObject $pres
        }

        if ($null -ne $ppt) {

            try {
                $ppt.Quit()
            }
            catch {
            }

            Release-ComObject $ppt
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


# ============================================================
# PRECHECK
# ============================================================

foreach ($file in @(
    $BeforePath,
    $AfterPath,
    $ClaudePath
)) {

    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "Required file missing: $file"
    }
}


New-Item `
    -ItemType Directory `
    -Path $OutputDir `
    -Force |
    Out-Null

New-Item `
    -ItemType Directory `
    -Path $WorkRoot `
    -Force |
    Out-Null


$stamp =
    Get-Date -Format "yyyyMMdd_HHmmss"

$workDir =
    Join-Path `
        $WorkRoot `
        "ppt_visual_reqa_v1_$stamp"

$beforeDir =
    Join-Path $workDir "before"

$afterDir =
    Join-Path $workDir "after"

New-Item `
    -ItemType Directory `
    -Path $workDir `
    -Force |
    Out-Null


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT VISUAL RE-QA V1"
Write-Host "========================================"
Write-Host ""

Write-Host "Before : $BeforePath"
Write-Host "After  : $AfterPath"
Write-Host "Work   : $workDir"
Write-Host ""


# ============================================================
# RENDER
# ============================================================

Write-Host "----------------------------------------"
Write-Host " RENDER BEFORE"
Write-Host "----------------------------------------"

$beforeCount =
    Export-PptSlides `
        -PptPath $BeforePath `
        -TargetDir $beforeDir


Write-Host ""
Write-Host "----------------------------------------"
Write-Host " RENDER AFTER"
Write-Host "----------------------------------------"

$afterCount =
    Export-PptSlides `
        -PptPath $AfterPath `
        -TargetDir $afterDir


if ($beforeCount -ne $afterCount) {
    throw "Slide count mismatch: BEFORE=$beforeCount AFTER=$afterCount"
}


$slideCount =
    $beforeCount


Write-Host ""
Write-Host "Slides : $slideCount"
Write-Host ""


# ============================================================
# CLAUDE QA
# ============================================================

$allItems =
    @()

$chunkScores =
    @()

$chunkNo =
    0


for (
    $start = 1;
    $start -le $slideCount;
    $start += $ChunkSize
) {

    $chunkNo++

    $end =
        [Math]::Min(
            $slideCount,
            $start + $ChunkSize - 1
        )

    $slideNumbers =
        @($start..$end)

    Write-Host ""
    Write-Host "----------------------------------------"
    Write-Host (
        " VISUAL RE-QA {0} | Slides {1}" -f
        $chunkNo,
        ($slideNumbers -join ",")
    )
    Write-Host "----------------------------------------"


    $pairs =
        @()

    foreach ($slideNo in $slideNumbers) {

        $pairs +=
            [pscustomobject]@{
                slide =
                    $slideNo

                before =
                    Join-Path `
                        $beforeDir `
                        ("slide_{0:D3}.png" -f $slideNo)

                after =
                    Join-Path `
                        $afterDir `
                        ("slide_{0:D3}.png" -f $slideNo)
            }
    }


    $pairJson =
        $pairs |
        ConvertTo-Json -Depth 4


    $chunkPath =
        Join-Path `
            $workDir `
            ("visual_qa_chunk_{0:D2}.json" -f $chunkNo)


    $resumeHit =
        $false


    if (
        -not [string]::IsNullOrWhiteSpace($ResumeDir) -and
        (Test-Path -LiteralPath $ResumeDir -PathType Container)
    ) {

        $resumeCandidate =
            Join-Path `
                $ResumeDir `
                ([System.IO.Path]::GetFileName($chunkPath))

        if (
            Test-Path `
                -LiteralPath $resumeCandidate `
                -PathType Leaf
        ) {

            Copy-Item `
                -LiteralPath $resumeCandidate `
                -Destination $chunkPath `
                -Force

            $resumeHit =
                $true

            Write-Host (
                "[RE-QA RESUME] Chunk {0:D2} -> REUSE" -f
                $chunkNo
            )
        }
    }


    if (-not $resumeHit) {

        Write-Host (
            "[RE-QA RESUME] Chunk {0:D2} -> CLAUDE" -f
            $chunkNo
        )


        $prompt = @"
You are performing an INDEPENDENT visual QA of a PowerPoint deck.

This is a SECOND-PASS visual evaluation after an automated revision.

You must inspect EVERY BEFORE and AFTER PNG listed below using the Read tool.

IMPORTANT:
- BEFORE is the original slide.
- AFTER is the current revised slide.
- Evaluate visual quality, not whether change occurred.
- Do not reward change for the sake of change.
- If AFTER is identical to BEFORE and the slide was already strong:
  better_than_before=false, severity=none is valid.
- better_than_before=true only when AFTER is clearly visually stronger.

Evaluate:
1. information hierarchy
2. alignment
3. spacing
4. whitespace
5. typography hierarchy
6. visual balance
7. readability
8. professionalism
9. consistency
10. image/text balance
11. clipping or overlap
12. whether AFTER is genuinely better than BEFORE

Severity definitions:

critical
= slide is unusable or seriously broken.

major
= substantial readability/layout/design problem that still requires revision.

minor
= visible polish issue but slide remains usable.

none
= no meaningful visual defect requiring correction.

For the entire chunk assign chunk_score from 0 to 100.

Calibration:
90-100 = excellent / production-ready
85-89  = strong
75-84  = usable but noticeable issues remain
65-74  = substantial visual improvement still needed
below 65 = poor

PAIR LIST:
$pairJson

Write VALID JSON to exactly this path:

$chunkPath

Required schema:

{
  "chunk": $chunkNo,
  "chunk_score": 85,
  "items": [
    {
      "slide": 1,
      "better_than_before": true,
      "severity": "none",
      "comment": "Concise evidence-based visual assessment."
    }
  ]
}

Rules:
- Include exactly one item for every slide in this chunk.
- severity must be exactly:
  none | minor | major | critical
- Do not change either PPT.
- Do not create any other deliverable.
- Do not use markdown fences inside the JSON file.

After successfully writing the JSON, output only:

PPT_VISUAL_REQA_CHUNK_COMPLETE
"@


        & $ClaudePath `
            -p $prompt `
            --allowedTools "Read" "Write"
    }


    if (
        -not (
            Test-Path `
                -LiteralPath $chunkPath `
                -PathType Leaf
        )
    ) {

        throw "Re-QA result missing: $chunkPath"
    }


    $chunkObject =
        [System.IO.File]::ReadAllText(
            $chunkPath
        ) |
        ConvertFrom-Json


    if ($null -eq $chunkObject.chunk_score) {
        throw "chunk_score missing: $chunkPath"
    }


    if ($null -eq $chunkObject.items) {
        throw "items missing: $chunkPath"
    }


    $items =
        @($chunkObject.items)


    if ($items.Count -ne $slideNumbers.Count) {

        throw (
            "Item count mismatch in chunk {0}: expected={1} actual={2}" -f
            $chunkNo,
            $slideNumbers.Count,
            $items.Count
        )
    }


    foreach ($item in $items) {

        $severity =
            ([string]$item.severity).ToLowerInvariant()

        if (
            $severity -notin @(
                "none",
                "minor",
                "major",
                "critical"
            )
        ) {

            throw (
                "Invalid severity in chunk {0}, slide {1}: {2}" -f
                $chunkNo,
                $item.slide,
                $item.severity
            )
        }


        $allItems +=
            [pscustomobject]@{
                slide =
                    [int]$item.slide

                better_than_before =
                    [bool]$item.better_than_before

                severity =
                    $severity

                comment =
                    [string]$item.comment
            }
    }


    $chunkScores +=
        [double]$chunkObject.chunk_score


    Write-Host (
        "Chunk score : {0}" -f
        $chunkObject.chunk_score
    )
}


# ============================================================
# AGGREGATE
# ============================================================

$allItems =
    @(
        $allItems |
        Sort-Object -Property slide
    )


if ($allItems.Count -ne $slideCount) {

    throw (
        "Final slide QA count mismatch: expected={0} actual={1}" -f
        $slideCount,
        $allItems.Count
    )
}


$average =
    [Math]::Round(
        (
            $chunkScores |
            Measure-Object -Average
        ).Average,
        1
    )


$betterCount =
    @(
        $allItems |
        Where-Object {
            $_.better_than_before
        }
    ).Count


$criticalCount =
    @(
        $allItems |
        Where-Object {
            $_.severity -eq "critical"
        }
    ).Count


$majorCount =
    @(
        $allItems |
        Where-Object {
            $_.severity -eq "major"
        }
    ).Count


$minorCount =
    @(
        $allItems |
        Where-Object {
            $_.severity -eq "minor"
        }
    ).Count


$pass =
    (
        $average -ge 85 -and
        $criticalCount -eq 0 -and
        $majorCount -eq 0
    )


$result =
    [ordered]@{

        before_pptx =
            $BeforePath

        after_pptx =
            $AfterPath

        average_score =
            $average

        slides_total =
            $slideCount

        slides_better =
            $betterCount

        critical =
            $criticalCount

        major =
            $majorCount

        minor =
            $minorCount

        pass =
            $pass

        chunk_scores =
            @($chunkScores)

        work_dir =
            $workDir

        slides =
            @($allItems)
    }


$outputPath =
    Join-Path `
        $OutputDir `
        ("Jungle_PPT_Visual_ReQA_{0}.json" -f $stamp)


$result |
    ConvertTo-Json -Depth 8 |
    Set-Content `
        -LiteralPath $outputPath `
        -Encoding UTF8


# ============================================================
# RESULT
# ============================================================

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT VISUAL RE-QA RESULT"
Write-Host "========================================"
Write-Host ""

Write-Host "Slides total   : $slideCount"
Write-Host "Average score  : $average"
Write-Host "Slides better  : $betterCount"
Write-Host "Critical       : $criticalCount"
Write-Host "Major          : $majorCount"
Write-Host "Minor          : $minorCount"
Write-Host "Visual pass    : $pass"
Write-Host "JSON           : $outputPath"
Write-Host "Work Dir       : $workDir"

Write-Host ""

if (
    $criticalCount -eq 0 -and
    $majorCount -eq 0
) {

    Write-Host "MAJOR GATE     : PASSED"
}
else {

    Write-Host "MAJOR GATE     : FAILED"
}

Write-Host ""
Write-Host "STATUS: PPT_VISUAL_REQA_COMPLETE"