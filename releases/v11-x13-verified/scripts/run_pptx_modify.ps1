param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()

if ($ext -ne ".pptx" -and $ext -ne ".ppt") {
    throw "PPT/PPTX 파일만 지원합니다."
}

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"

if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if (!(Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath = Join-Path `
    $outputDir `
    "${baseName}_modified_${dateStamp}.pptx"

$extractPath = Join-Path `
    $tempDir `
    "pptx_extract_${dateStamp}.json"

$planPath = Join-Path `
    $tempDir `
    "pptx_edit_plan_${dateStamp}.json"


Write-Host ""
Write-Host "========================================"
Write-Host " Jungle PPTX Modify"
Write-Host "========================================"
Write-Host ""
Write-Host "ORIGINAL :" $FilePath
Write-Host "OUTPUT   :" $outputPath
Write-Host ""


# ------------------------------------------------------------
# 1. 원본 복사
# ------------------------------------------------------------

Copy-Item `
    -LiteralPath $FilePath `
    -Destination $outputPath `
    -Force


# ------------------------------------------------------------
# 2. PPT 내용 추출
# ------------------------------------------------------------

& "C:\AI\scripts\extract_office_content.ps1" `
    -FilePath $FilePath `
    -OutputPath $extractPath

if (!(Test-Path -LiteralPath $extractPath)) {
    throw "PPT 내용 추출에 실패했습니다."
}


# ------------------------------------------------------------
# 3. Claude 수정계획 생성
# ------------------------------------------------------------

$prompt = @"
다음 PowerPoint 내용을 직접 읽고,
사용자의 요청에 맞는 수정계획을 JSON으로 작성하세요.

[PowerPoint 추출 내용]
$extractPath

[사용자 요청]
$UserRequest

중요 원칙:

1. 사용자가 요청한 부분만 수정할 것.
2. 기존 슬라이드 구성과 디자인은 유지할 것.
3. 슬라이드를 임의로 삭제하거나 추가하지 말 것.
4. 기존 내용을 불필요하게 축약하거나 재작성하지 말 것.
5. old_text는 추출 JSON에 존재하는 텍스트를 정확히 그대로 사용할 것.
6. new_text에는 수정 후 들어갈 전체 텍스트를 작성할 것.
7. 수정할 필요가 없는 텍스트는 edits에 넣지 말 것.
8. 추측해서 새로운 사실을 추가하지 말 것.
9. 한국어는 자연스럽게 작성할 것.

다음 JSON 형식으로만 저장하세요.

{
  "edits": [
    {
      "slide_number": 1,
      "old_text": "기존 문구",
      "new_text": "수정 문구",
      "reason": "수정 이유"
    }
  ]
}

JSON 파일을 다음 위치에 UTF-8로 저장하세요.

$planPath

작업이 끝나면 다른 설명 없이 PPT_EDIT_PLAN_CREATED 만 출력하세요.
"@

# PowerShell 5.1의 긴 한국어/따옴표 인수 손상을 피하기 위해
# 전체 지시문을 UTF-8 파일로 저장한 뒤 Claude가 직접 읽게 한다.

$instructionPath = Join-Path `
    $tempDir `
    "pptx_instruction_${dateStamp}.txt"

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllText(
    $instructionPath,
    $prompt,
    $utf8Bom
)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Do not use Google Drive, connectors, web search, or any external source. Use only the local files specified in the instruction file. When finished, output only PPT_EDIT_PLAN_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath)) {
    throw "Claude 수정계획 JSON이 생성되지 않았습니다."
}


# ------------------------------------------------------------
# 4. 수정계획 검증
# ------------------------------------------------------------

$plan = Get-Content `
    -LiteralPath $planPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

if ($null -eq $plan.edits) {
    throw "수정계획에 edits가 없습니다."
}

Write-Host ""
Write-Host "EDIT PLAN :" $plan.edits.Count "건"
Write-Host ""


# ------------------------------------------------------------
# 5. PowerPoint COM으로 실제 수정
# ------------------------------------------------------------

$ppt  = $null
$pres = $null

$applied = 0
$failed  = @()

try {

    $ppt = New-Object -ComObject PowerPoint.Application

    $pres = $ppt.Presentations.Open(
        $outputPath,
        $false,
        $false,
        $false
    )

    foreach ($edit in $plan.edits) {

        $slideNumber = [int]$edit.slide_number
        $oldText = [string]$edit.old_text
        $newText = [string]$edit.new_text

        if (
            $slideNumber -lt 1 -or
            $slideNumber -gt $pres.Slides.Count
        ) {
            $failed += "Slide $slideNumber : 존재하지 않는 슬라이드"
            continue
        }

        $slide = $pres.Slides.Item($slideNumber)
        $matched = $false

        foreach ($shape in $slide.Shapes) {

            try {

                if (
                    $shape.HasTextFrame -eq -1 -and
                    $shape.TextFrame.HasText -eq -1
                ) {

                    $currentText =
                        [string]$shape.TextFrame.TextRange.Text

                    if ($currentText -eq $oldText) {

                        $shape.TextFrame.TextRange.Text = $newText

                        $matched = $true
                        $applied++

                        break
                    }
                }
            }
            catch {}
        }

        if (!$matched) {
            $failed += "Slide $slideNumber : 원문을 찾지 못함"
        }
    }

    $pres.Save()
}
finally {

    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) |
                Out-Null
        }
        catch {}
    }

    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) |
                Out-Null
        }
        catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ------------------------------------------------------------
# 6. 결과 검증
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $outputPath)) {
    throw "PPT 수정본이 생성되지 않았습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPTX MODIFY RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Modify  : OK"
Write-Host "Applied :" $applied
Write-Host "Failed  :" $failed.Count
Write-Host "Original:" $FilePath
Write-Host "File    :" $outputPath
Write-Host "Size    :" (Get-Item $outputPath).Length

if ($failed.Count -gt 0) {

    Write-Host ""
    Write-Host "미적용 항목:"

    foreach ($item in $failed) {
        Write-Host " -" $item
    }

    Write-Host ""
    Write-Host "STATUS: PPTX_MODIFY_PARTIAL"
}
else {

    Write-Host ""
    Write-Host "STATUS: PPTX_MODIFY_COMPLETE"
}