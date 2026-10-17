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

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
$ext       = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath      = Join-Path $outputDir "${baseName}_created_${dateStamp}.docx"
$planPath        = Join-Path $tempDir "docx_create_plan_${dateStamp}.json"
$instructionPath = Join-Path $tempDir "docx_create_instruction_${dateStamp}.txt"
$sourceForClaude = $SourcePath

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle DOCX Create"
Write-Host "========================================"
Write-Host ""
Write-Host "SOURCE :" $SourcePath
Write-Host "OUTPUT :" $outputPath
Write-Host ""

# Office 파일이면 먼저 텍스트 구조 추출
if ($ext -in @(".pptx", ".ppt", ".docx", ".doc")) {

    $extractPath = Join-Path $tempDir "create_source_${dateStamp}.json"

    & "C:\AI\scripts\extract_office_content.ps1" `
        -FilePath $SourcePath `
        -OutputPath $extractPath

    if (!(Test-Path -LiteralPath $extractPath)) {
        throw "참고자료 추출에 실패했습니다."
    }

    $sourceForClaude = $extractPath
}

$prompt = @"
다음 참고자료를 읽고 사용자의 요청에 맞는 새로운 Word 문서를 설계하세요.

[참고자료]
$sourceForClaude

[사용자 요청]
$UserRequest

[작성 원칙]
- 참고자료의 실제 내용을 기반으로 작성할 것
- 참고자료에 없는 사실을 임의로 추가하지 말 것
- 사용자가 지정한 목적과 독자를 우선할 것
- 기존 문서를 단순 복사하지 말고 요청에 맞게 새 문서 구조로 재구성할 것
- 자연스러운 한국어를 사용할 것
- 불필요한 AI식 표현, 과도한 수식어, 반복 표현을 피할 것
- 제목과 본문 구조를 명확히 할 것

아래 JSON 구조로 작성하세요.

{
  "title": "문서 제목",
  "sections": [
    {
      "heading": "1. 제목",
      "paragraphs": [
        "본문 문단 1",
        "본문 문단 2"
      ]
    }
  ]
}

JSON을 다음 경로에 UTF-8로 저장하세요.

$planPath

완료 후 다른 설명 없이
DOCX_CREATE_PLAN_CREATED
만 출력하세요.
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllText(
    $instructionPath,
    $prompt,
    $utf8Bom
)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified there. Do not use web search, Google Drive, connectors, or external sources. When finished, output only DOCX_CREATE_PLAN_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath)) {
    throw "DOCX 생성계획 JSON이 만들어지지 않았습니다."
}

$plan = Get-Content `
    -LiteralPath $planPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace([string]$plan.title)) {
    throw "문서 제목이 없습니다."
}


# ------------------------------------------------------------
# 5. OpenXML 기반 DOCX 생성
# Word COM을 사용하지 않으므로 SaveAs2 멈춤 문제를 피한다.
# ------------------------------------------------------------

$writerScript = "C:\AI\scripts\write_docx_from_plan.ps1"

if (!(Test-Path -LiteralPath $writerScript -PathType Leaf)) {
    throw "OpenXML DOCX Writer를 찾을 수 없습니다: $writerScript"
}

& $writerScript `
    -PlanPath $planPath `
    -OutputPath $outputPath

if (-not $?) {
    throw "OpenXML DOCX 생성에 실패했습니다."
}

if (!(Test-Path -LiteralPath $outputPath)) {
    throw "새 Word 문서 생성에 실패했습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE DOCX CREATE RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Create : OK"
Write-Host "Source :" $SourcePath
Write-Host "File   :" $outputPath
Write-Host "Size   :" (Get-Item $outputPath).Length
Write-Host ""
Write-Host "STATUS: DOCX_CREATE_COMPLETE"
