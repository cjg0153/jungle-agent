param(
    [Parameter(Mandatory=$true)]
    [string]$DraftPath,

    [Parameter(Mandatory=$true)]
    [string]$QaPath,

    [string]$SourcePath = "",

    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $DraftPath -PathType Leaf)) {
    throw "수정할 PPTX를 찾을 수 없습니다: $DraftPath"
}

if (!(Test-Path -LiteralPath $QaPath -PathType Leaf)) {
    throw "QA JSON을 찾을 수 없습니다: $QaPath"
}

$qa = Get-Content -LiteralPath $QaPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ([string]$qa.overall_status -eq "PASS" -and -not [bool]$qa.revision_required) {
    Write-Host ""
    Write-Host "QA Status : PASS"
    Write-Host "Revision  : NOT REQUIRED"
    Write-Host "STATUS: PPT_QA_REVISION_SKIPPED"
    return
}

$tempDir = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$sourceExtractPath = ""

if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {

    if (!(Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "원본 참고 PPTX를 찾을 수 없습니다: $SourcePath"
    }

    $extractor = "C:\AI\scripts\extract_office_content.ps1"

    if (!(Test-Path -LiteralPath $extractor -PathType Leaf)) {
        throw "Office extractor를 찾을 수 없습니다: $extractor"
    }

    $sourceExtractPath = Join-Path $tempDir "ppt_qa_revision_source_${stamp}.json"

    & $extractor -FilePath $SourcePath -OutputPath $sourceExtractPath

    if (!(Test-Path -LiteralPath $sourceExtractPath -PathType Leaf)) {
        throw "원본 참고자료 추출에 실패했습니다."
    }
}

$modifier = "C:\AI\scripts\run_pptx_modify.ps1"

if (!(Test-Path -LiteralPath $modifier -PathType Leaf)) {
    throw "PPTX 수정 엔진을 찾을 수 없습니다: $modifier"
}

$groundingSection = ""

if (-not [string]::IsNullOrWhiteSpace($sourceExtractPath)) {
    $groundingSection = @"

[원본 참고자료]
아래 JSON을 반드시 읽고 사실 근거로 사용하세요.
$sourceExtractPath
"@
}

$revisionRequest = @"
이 PPT는 Codex 최종 QA에서 수정 필요 판정을 받았습니다.

[원래 사용자 요청]
$UserRequest

[Codex QA 결과]
아래 JSON을 반드시 읽고 Critical/Major 이슈를 우선 수정하세요.
$QaPath
$groundingSection

[수정 원칙]
1. 현재 PPT의 슬라이드 수, 레이아웃, 테마, 도형 배치, 디자인은 유지하세요.
2. 내용과 문장만 수정하세요.
3. QA의 Critical/Major 이슈를 우선 해결하고 Minor는 품질을 해치지 않는 범위에서 반영하세요.
4. QA가 요구한 정보가 원본 참고자료에 있으면 그 사실을 사용하세요.
5. 원본 참고자료에도 없는 숫자, 근거, 책임자, 일정, 성과를 임의로 만들지 마세요.
6. 근거가 없는 문장은 삭제하거나 사실 범위 안에서 표현을 낮추세요.
7. 임원 보고용 한국어로 짧고 명확하게 정리하세요.
8. humanizer-kr 관점에서 AI식 상투어, 과도한 수식어, 반복 결론, 번역투를 제거하세요.
9. 기존 디자인에 맞춰 현재 텍스트박스의 분량을 크게 초과하지 마세요.
10. 원래 사용자 요청의 핵심 목적을 유지하세요.

수정 결과는 새 PPTX 파일로 저장하고 원본 DraftPath는 보존하세요.
"@

$started = Get-Date

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE CLAUDE REVISION GATE"
Write-Host "========================================"
Write-Host ""
Write-Host "Draft :" $DraftPath
Write-Host "QA    :" $QaPath
if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
    Write-Host "Source:" $SourcePath
}
Write-Host "Status:" $qa.overall_status
Write-Host ""
Write-Host "Claude revision started..."

& $modifier -FilePath $DraftPath -UserRequest $revisionRequest

if (-not $?) {
    throw "Claude PPT revision 실행에 실패했습니다."
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($DraftPath)

$candidates = Get-ChildItem -LiteralPath $outputDir -Filter "${baseName}_modified_*.pptx" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending

$revisedPath = $null

$candidateArray = @($candidates)
if ($candidateArray.Count -gt 0) {
    $revisedPath = $candidateArray[0].FullName
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE CLAUDE REVISION RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Revision : COMPLETE"
if ($null -ne $revisedPath) {
    Write-Host "File     :" $revisedPath
}
else {
    Write-Host "File     : Check C:\AI\work\output for the newest modified PPTX."
}
Write-Host ""
Write-Host "STATUS: PPT_QA_REVISION_COMPLETE"
