param(
    [Parameter(Mandatory=$true)]
    [string]$DraftPath,

    [Parameter(Mandatory=$true)]
    [string]$TriagePath,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

foreach ($p in @($DraftPath, $TriagePath, $SourcePath)) {
    if (!(Test-Path -LiteralPath $p -PathType Leaf)) {
        throw "파일을 찾을 수 없습니다: $p"
    }
}

$revisionEngine = "C:\AI\scripts\run_pptx_qa_revision_v1.ps1"
if (!(Test-Path -LiteralPath $revisionEngine -PathType Leaf)) {
    throw "기존 PPT QA revision engine을 찾을 수 없습니다: $revisionEngine"
}

$tempDir = "C:\AI\work\temp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$filteredQaPath = Join-Path $tempDir "auto_fix_only_${stamp}.json"

$triage = Get-Content -LiteralPath $TriagePath -Raw -Encoding UTF8 | ConvertFrom-Json
$autoIssues = @($triage.issues | Where-Object { $_.classification -eq "AUTO_FIX" })
$factIssues = @($triage.issues | Where-Object { $_.classification -eq "FACT_DECISION" })
$overIssues = @($triage.issues | Where-Object { $_.classification -eq "DETAIL_OVERREACH" })

if ($autoIssues.Count -eq 0) {
    Write-Host ""
    Write-Host "AUTO_FIX issues : 0"
    Write-Host "STATUS: AUTO_REVISION_SKIPPED"
    return
}

$mappedIssues = @()

foreach ($i in $autoIssues) {
    $mappedIssues += [PSCustomObject]@{
        severity       = [string]$i.original_severity
        category       = [string]$i.original_category
        location       = [string]$i.location
        problem        = [string]$i.reason
        recommendation = [string]$i.safe_action
        classification = "AUTO_FIX"
    }
}

$filtered = [PSCustomObject]@{
    overall_status = "REVISE"
    review_scope   = "AUTO_FIX_ONLY"
    revision_policy = [PSCustomObject]@{
        allowed = @(
            "AUTO_FIX 항목만 수정",
            "원본 자료에 근거한 문구 복원",
            "중복 제거",
            "명칭·용어·표현 정리",
            "임원 보고용 간결성 개선"
        )
        forbidden = @(
            "FACT_DECISION 항목 수정",
            "DETAIL_OVERREACH 내용을 본문에 추가",
            "원본에 없는 숫자·일정·담당자·승인범위 생성",
            "295 FTE의 누적/연평균 여부 임의 확정",
            "100개 Agent의 후보/구축/운영 단계 임의 확정",
            "새로운 사실·근거·외부 정보 추가",
            "디자인 템플릿 변경"
        )
        fact_decision_count = $factIssues.Count
        detail_overreach_count = $overIssues.Count
    }
    issues = $mappedIssues
    final_summary = "이 파일은 Jungle QA Triage에서 AUTO_FIX로 분류된 항목만 포함한다. 이 목록 밖의 이슈는 수정하지 않는다."
}

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$json = $filtered | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText($filteredQaPath, $json, $utf8Bom)

$revisionRequest = @"
$UserRequest

[STRICT AUTO-REVISION RULE]
이번 수정에서는 제공된 AUTO_FIX 전용 QA JSON의 항목만 수정한다.

반드시 지킬 것:
1. FACT_DECISION 항목은 수정하지 않는다.
2. DETAIL_OVERREACH 항목은 본문에 추가하지 않는다.
3. 원본에 없는 숫자, 일정, 담당자, 승인범위, 산식, 의사결정 내용을 만들지 않는다.
4. 295 FTE의 집계 방식과 Agent 100개의 단계 정의를 임의로 확정하지 않는다.
5. 기존 디자인/레이아웃/템플릿은 최대한 보존한다.
6. AUTO_FIX 항목은 원본 참고자료에서 근거가 있는 내용만 복원하거나 문장을 간결하게 다듬는다.
7. 수정 범위를 넓히지 않는다.
"@

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT AUTO REVISION v1"
Write-Host "========================================"
Write-Host ""
Write-Host "AUTO_FIX        :" $autoIssues.Count
Write-Host "FACT_DECISION   :" $factIssues.Count "(locked)"
Write-Host "DETAIL_OVERREACH:" $overIssues.Count "(locked)"
Write-Host "Filtered QA     :" $filteredQaPath
Write-Host ""
Write-Host "Claude auto revision started..."

& $revisionEngine `
    -DraftPath $DraftPath `
    -QaPath $filteredQaPath `
    -SourcePath $SourcePath `
    -UserRequest $revisionRequest

Write-Host ""
Write-Host "STATUS: PPT_AUTO_REVISION_V1_COMPLETE"
