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

$outputDir = "C:\AI\work\output"

if (!(Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath = Join-Path `
    $outputDir `
    "${baseName}_review_${dateStamp}.md"

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle File Review"
Write-Host "========================================"
Write-Host ""
Write-Host "INPUT :" $FilePath
Write-Host "OUTPUT:" $outputPath
Write-Host ""

$prompt = @"
다음 파일을 직접 읽고 사용자의 요청에 따라 검토하세요.

[원본 파일]
$FilePath

[사용자 요청]
$UserRequest

[검토 원칙]
- 반드시 실제 원본 파일의 내용을 읽고 판단할 것
- 원본 파일에 없는 내용을 추측하지 말 것
- 사용자가 지정한 대상자와 목적을 최우선으로 고려할 것
- 이해하기 어려운 부분, 설명 부족, 난이도 변화, 불필요한 복잡성, 흐름 문제를 구체적으로 찾을 것
- 잘 된 부분도 함께 평가할 것
- 수정이 필요한 경우 무엇을 어떻게 바꿀지 구체적으로 제안할 것
- 자연스러운 한국어로 작성할 것
- 원본 파일은 절대 수정하지 말 것

검토 결과를 다음 경로에 UTF-8 Markdown 파일로 저장하세요.

$outputPath

저장이 완료되면 다른 설명 없이 REVIEW_FILE_CREATED 만 출력하세요.
"@

claude -p $prompt --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $outputPath)) {
    throw "검토 결과 파일이 생성되지 않았습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE FILE REVIEW RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Review : OK"
Write-Host "File   :" $outputPath
Write-Host "Size   :" (Get-Item $outputPath).Length
Write-Host ""
Write-Host "STATUS: FILE_REVIEW_COMPLETE"