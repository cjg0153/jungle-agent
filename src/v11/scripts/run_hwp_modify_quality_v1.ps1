param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string]$UserRequest
)

$ErrorActionPreference = "Stop"

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "HWP 파일 누락: $FilePath"
}

if ([System.IO.Path]::GetExtension($FilePath).ToLowerInvariant() -ne ".hwp") {
    throw "HWP 파일만 지원합니다: $FilePath"
}

foreach ($d in @($outputDir,$tempDir)) {
    if (!(Test-Path -LiteralPath $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

$docxPipeline = "C:\AI\scripts\run_docx_modify_quality_v1.ps1"

if (!(Test-Path -LiteralPath $docxPipeline -PathType Leaf)) {
    throw "DOCX 품질 파이프라인 누락: $docxPipeline"
}

function Normalize-JungleText {
    param([string]$Text)

    if ($null -eq $Text) { return "" }

    return (($Text -replace '\s+', ' ').Trim())
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$sourceDir = [System.IO.Path]::GetDirectoryName($FilePath)
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

$workDocx = Join-Path $tempDir "${baseName}_HWP_WORK_${stamp}.docx"
$outHwp   = Join-Path $sourceDir "${baseName}_수정본.hwp"

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE HWP QUALITY PIPELINE v1"
Write-Host "========================================"
Write-Host "Source HWP :" $FilePath
Write-Host "Work DOCX  :" $workDocx
Write-Host "Output HWP :" $outHwp
Write-Host ""


# ============================================================
# STEP 1. HWP -> DOCX
# ============================================================

Write-Host "[1/5] HWP -> DOCX (OOXML)"

$hwp = New-Object -ComObject HWPFrame.HwpObject

try {

    $opened = $hwp.Open(
        $FilePath,
        "HWP",
        ""
    )

    if (-not $opened) {
        throw "원본 HWP 열기 실패"
    }

    $sourceText = [string]$hwp.GetTextFile(
        "UNICODE",
        ""
    )

    if (Test-Path -LiteralPath $workDocx) {
        Remove-Item -LiteralPath $workDocx -Force
    }

    [void]$hwp.SaveAs(
        $workDocx,
        "OOXML",
        ""
    )

    Start-Sleep -Milliseconds 800

    if (!(Test-Path -LiteralPath $workDocx -PathType Leaf)) {
        throw "HWP -> DOCX 변환 실패"
    }

    $hwp.Clear(1)

    $openedDocx = $hwp.Open(
        $workDocx,
        "OOXML",
        ""
    )

    if (-not $openedDocx) {
        throw "변환된 DOCX 재오픈 실패"
    }

    $workText = [string]$hwp.GetTextFile(
        "UNICODE",
        ""
    )

    $sourceNorm = Normalize-JungleText $sourceText
    $workNorm   = Normalize-JungleText $workText

    if ($sourceNorm -ne $workNorm) {
        throw "HWP -> DOCX 변환 과정에서 본문 텍스트가 달라졌습니다."
    }

    $workInfo = Get-Item -LiteralPath $workDocx

    Write-Host "      DOCX created :" $workInfo.Length "bytes"
    Write-Host "      Text integrity: EXACT"
}
finally {

    try { $hwp.Quit() } catch {}

    if ($null -ne $hwp) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
    }
}


# ============================================================
# STEP 2. Existing DOCX modify + QA pipeline
# ============================================================

Write-Host ""
Write-Host "[2/5] DOCX modify + Codex QA"

$qualityStart = Get-Date

$wrappedRequest = @"
[HWP WRAPPER MODE]

Original HWP:
$FilePath

Working DOCX:
$workDocx

Apply the user's requested edits to the WORKING DOCX.
Do not attempt to create HWP yourself.
Jungle will convert the final approved DOCX back to HWP.

[USER REQUEST]
$UserRequest
"@

$captured = New-Object System.Collections.Generic.List[string]

& $docxPipeline `
    -FilePath $workDocx `
    -UserRequest $wrappedRequest `
    6>&1 |
ForEach-Object {

    $line = [string]$_

    [void]$captured.Add($line)

    Write-Host $line
}


# ============================================================
# STEP 3. Resolve final approved DOCX
# ============================================================

Write-Host ""
Write-Host "[3/5] Resolve final approved DOCX"

$finalDocx = $null


# 3-A. Prefer "Final :" emitted by quality pipeline

foreach ($line in $captured) {

    if ($line -match '^\s*Final\s*:\s*(.+?\.docx)\s*$') {

        $candidatePath = $Matches[1].Trim()

        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            $finalDocx = $candidatePath
        }
    }
}


# 3-B. Acceptance JSON fallback

if ([string]::IsNullOrWhiteSpace($finalDocx)) {

    $acceptanceFiles = @(
        Get-ChildItem `
            -LiteralPath $outputDir `
            -Filter "*_acceptance.json" `
            -File `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LastWriteTime -ge $qualityStart.AddSeconds(-2)
        } |
        Sort-Object LastWriteTime -Descending
    )

    foreach ($acceptanceFile in $acceptanceFiles) {

        try {

            $a = Get-Content `
                -LiteralPath $acceptanceFile.FullName `
                -Raw `
                -Encoding UTF8 |
                ConvertFrom-Json

            $candidatePath = [string]$a.final_file

            if (
                -not [string]::IsNullOrWhiteSpace($candidatePath) `
                -and (Test-Path -LiteralPath $candidatePath -PathType Leaf) `
                -and ([System.IO.Path]::GetExtension($candidatePath).ToLowerInvariant() -eq ".docx")
            ) {
                $finalDocx = $candidatePath
                break
            }
        }
        catch {}
    }
}


# 3-C. FINAL file fallback

if ([string]::IsNullOrWhiteSpace($finalDocx)) {

    $finalCandidate = @(
        Get-ChildItem `
            -LiteralPath $outputDir `
            -Filter "*_FINAL_*.docx" `
            -File `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LastWriteTime -ge $qualityStart.AddSeconds(-2)
        } |
        Sort-Object LastWriteTime -Descending
    ) | Select-Object -First 1

    if ($null -ne $finalCandidate) {
        $finalDocx = $finalCandidate.FullName
    }
}


if ([string]::IsNullOrWhiteSpace($finalDocx)) {
    throw "DOCX 품질 파이프라인의 최종 승인 파일을 찾지 못했습니다."
}

if (!(Test-Path -LiteralPath $finalDocx -PathType Leaf)) {
    throw "최종 승인 DOCX 누락: $finalDocx"
}

Write-Host "      FINAL DOCX :" $finalDocx


# ============================================================
# STEP 4. Final DOCX -> HWP
# ============================================================

Write-Host ""
Write-Host "[4/5] Final DOCX -> HWP"

$hwp2 = New-Object -ComObject HWPFrame.HwpObject

try {

    $openedFinalDocx = $hwp2.Open(
        $finalDocx,
        "OOXML",
        ""
    )

    if (-not $openedFinalDocx) {
        throw "최종 승인 DOCX 열기 실패"
    }

    $finalDocxText = [string]$hwp2.GetTextFile(
        "UNICODE",
        ""
    )

    if (Test-Path -LiteralPath $outHwp) {
        Remove-Item -LiteralPath $outHwp -Force
    }

    [void]$hwp2.SaveAs(
        $outHwp,
        "HWP",
        ""
    )

    Start-Sleep -Milliseconds 800

    if (!(Test-Path -LiteralPath $outHwp -PathType Leaf)) {
        throw "최종 HWP 생성 실패"
    }


    # ========================================================
    # STEP 5. Re-open final HWP + text integrity
    # ========================================================

    Write-Host ""
    Write-Host "[5/5] Re-open final HWP + integrity check"

    $hwp2.Clear(1)

    $reopened = $hwp2.Open(
        $outHwp,
        "HWP",
        ""
    )

    if (-not $reopened) {
        throw "최종 HWP 재오픈 실패"
    }

    $finalHwpText = [string]$hwp2.GetTextFile(
        "UNICODE",
        ""
    )

    $docxNorm = Normalize-JungleText $finalDocxText
    $hwpNorm  = Normalize-JungleText $finalHwpText

    if ($docxNorm -ne $hwpNorm) {
        throw "최종 DOCX -> HWP 변환 후 본문 텍스트가 달라졌습니다."
    }

    $hwpInfo = Get-Item -LiteralPath $outHwp

    if ($hwpInfo.Length -lt 1000) {
        throw "최종 HWP 파일 크기가 비정상적입니다: $($hwpInfo.Length) bytes"
    }

    Write-Host ""
    Write-Host "[JUNGLE HWP ARTIFACT STATUS]"
    Write-Host "HWP_CREATED: TRUE"
    Write-Host "HWP_REOPEN_OK: TRUE"
    Write-Host "TEXT_EQUIVALENT: TRUE"
    Write-Host "SOURCE_HWP:" $FilePath
    Write-Host "WORK_DOCX:" $workDocx
    Write-Host "FINAL_DOCX:" $finalDocx
    Write-Host "OUTPUT_HWP:" $outHwp
    Write-Host "HWP_SIZE_BYTES:" $hwpInfo.Length

    Write-Host ""
    Write-Host "STATUS: HWP_MODIFY_QUALITY_COMPLETE"
}
finally {

    try { $hwp2.Quit() } catch {}

    if ($null -ne $hwp2) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp2) | Out-Null
    }
}