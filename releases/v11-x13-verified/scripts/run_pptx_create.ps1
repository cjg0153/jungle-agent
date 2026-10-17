param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    throw "참고 파일을 찾을 수 없습니다: $SourcePath"
}

$sourceExt = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()

if ($sourceExt -ne ".pptx") {
    throw "PPTX CREATE v1은 참고자료가 PPTX인 경우부터 지원합니다."
}

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath      = Join-Path $outputDir "${baseName}_created_${dateStamp}.pptx"
$extractPath     = Join-Path $tempDir "pptx_create_source_${dateStamp}.json"
$planPath        = Join-Path $tempDir "pptx_create_plan_${dateStamp}.json"
$instructionPath = Join-Path $tempDir "pptx_create_instruction_${dateStamp}.txt"

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle PPTX Create"
Write-Host "========================================"
Write-Host ""
Write-Host "SOURCE :" $SourcePath
Write-Host "OUTPUT :" $outputPath
Write-Host ""

# ------------------------------------------------------------
# 1. 참고 PPT 내용 추출
# ------------------------------------------------------------

& "C:\AI\scripts\extract_office_content.ps1" `
    -FilePath $SourcePath `
    -OutputPath $extractPath

if (!(Test-Path -LiteralPath $extractPath -PathType Leaf)) {
    throw "PPT 내용 추출에 실패했습니다."
}

# ------------------------------------------------------------
# 2. Claude가 새 PPT 구성계획 작성
# ------------------------------------------------------------

$prompt = @"
다음 PowerPoint의 실제 내용을 읽고,
사용자의 요청에 맞는 새로운 PowerPoint 구성을 설계하세요.

[참고 PowerPoint 추출 내용]
$extractPath

[사용자 요청]
$UserRequest

[작성 원칙]

1. 참고자료의 실제 내용에 근거할 것.
2. 새로운 사실을 임의로 추가하지 말 것.
3. 사용자가 지정한 대상과 목적을 최우선으로 고려할 것.
4. 한 슬라이드에 너무 많은 내용을 넣지 말 것.
5. 제목은 짧고 명확하게 작성할 것.
6. 본문 bullet은 슬라이드당 가급적 3~5개로 제한할 것.
7. 임원 보고자료처럼 간결하고 자연스러운 한국어를 사용할 것.
8. 불필요한 AI식 수식어와 반복 표현을 피할 것.
9. 기존 PPT의 테마와 마스터는 이후 Jungle이 재사용하므로 콘텐츠 구조에 집중할 것.
10. 첫 장은 title 형식으로 구성할 것.

아래 JSON 구조로 작성하세요.

{
  "deck_title": "전체 PPT 제목",
  "slides": [
    {
      "slide_type": "title",
      "title": "제목",
      "subtitle": "부제",
      "bullets": []
    },
    {
      "slide_type": "content",
      "title": "슬라이드 제목",
      "subtitle": "",
      "bullets": [
        "핵심 내용 1",
        "핵심 내용 2",
        "핵심 내용 3"
      ]
    }
  ]
}

JSON을 다음 위치에 UTF-8로 저장하세요.

$planPath

완료 후 다른 설명 없이
PPTX_CREATE_PLAN_CREATED
만 출력하세요.
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllText(
    $instructionPath,
    $prompt,
    $utf8Bom
)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified there. Do not use web search, Google Drive, connectors, or external sources. When finished, output only PPTX_CREATE_PLAN_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath -PathType Leaf)) {
    throw "PPTX 생성계획 JSON이 만들어지지 않았습니다."
}

$plan = Get-Content `
    -LiteralPath $planPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

if ($null -eq $plan.slides) {
    throw "PPTX 생성계획에 slides가 없습니다."
}

Write-Host ""
Write-Host "SLIDE PLAN :" $plan.slides.Count "장"
Write-Host ""

# ------------------------------------------------------------
# 3. 원본 PPT 복사
#    테마 / 마스터 / 페이지 크기 유지
# ------------------------------------------------------------

Copy-Item `
    -LiteralPath $SourcePath `
    -Destination $outputPath `
    -Force

# ------------------------------------------------------------
# 4. PowerPoint COM으로 새 슬라이드 구성
# ------------------------------------------------------------

$ppt  = $null
$pres = $null

try {

    $ppt = New-Object -ComObject PowerPoint.Application

    $pres = $ppt.Presentations.Open(
        $outputPath,
        $false,
        $false,
        $false
    )

    # 기존 슬라이드 삭제
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) {
        $pres.Slides.Item($i).Delete()
    }

    $slideIndex = 0

    foreach ($item in $plan.slides) {

        $slideIndex++

        $slideType = ([string]$item.slide_type).ToLowerInvariant()

        if ($slideType -eq "title") {

            # ppLayoutTitle = 1
            $slide = $pres.Slides.Add($slideIndex, 1)

            if ($null -ne $slide.Shapes.Title) {
                $slide.Shapes.Title.TextFrame.TextRange.Text =
                    [string]$item.title
            }

            try {
                $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text =
                    [string]$item.subtitle
            }
            catch {}
        }
        else {

            # ppLayoutText = 2
            $slide = $pres.Slides.Add($slideIndex, 2)

            if ($null -ne $slide.Shapes.Title) {
                $slide.Shapes.Title.TextFrame.TextRange.Text =
                    [string]$item.title
            }

            $bodyText = ""

            if ($null -ne $item.bullets) {
                $bodyText = ($item.bullets | ForEach-Object {
                    [string]$_
                }) -join "`r"
            }

            try {

                $bodyShape = $slide.Shapes.Placeholders.Item(2)

                $bodyShape.TextFrame.TextRange.Text = $bodyText

                if ($item.bullets.Count -gt 0) {

                    for ($p = 1; $p -le $item.bullets.Count; $p++) {

                        try {
                            $bodyShape.TextFrame.TextRange.Paragraphs(
                                $p,
                                1
                            ).ParagraphFormat.Bullet.Visible = -1
                        }
                        catch {}
                    }
                }
            }
            catch {}
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

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "PPTX 생성에 실패했습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPTX CREATE RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Create :" "OK"
Write-Host "Slides :" $plan.slides.Count
Write-Host "Source :" $SourcePath
Write-Host "File   :" $outputPath
Write-Host "Size   :" (Get-Item $outputPath).Length
Write-Host ""
Write-Host "STATUS: PPTX_CREATE_COMPLETE"