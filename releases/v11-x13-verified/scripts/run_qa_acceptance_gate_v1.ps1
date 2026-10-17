param(
    [Parameter(Mandatory=$true)]
    [string]$BaselinePath,

    [Parameter(Mandatory=$true)]
    [string]$BaselineQaPath,

    [Parameter(Mandatory=$true)]
    [string]$CandidatePath,

    [Parameter(Mandatory=$true)]
    [string]$CandidateQaPath,

    [string]$FinalName = ""
)

$ErrorActionPreference = "Stop"

foreach ($p in @($BaselinePath, $BaselineQaPath, $CandidatePath, $CandidateQaPath)) {
    if (!(Test-Path -LiteralPath $p -PathType Leaf)) {
        throw "파일을 찾을 수 없습니다: $p"
    }
}

function Get-QaMetrics {
    param([string]$Path)

    $qa = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json

    $critical = 0
    $major = 0
    $minor = 0
    $score = 0

    try { $critical = [int]$qa.issue_counts.critical } catch {}
    try { $major = [int]$qa.issue_counts.major } catch {}
    try { $minor = [int]$qa.issue_counts.minor } catch {}
    try { $score = [int]$qa.executive_readiness_score } catch {}

    [PSCustomObject]@{
        Score    = $score
        Critical = $critical
        Major    = $major
        Minor    = $minor
        Status   = [string]$qa.overall_status
        Scope    = [string]$qa.review_scope
    }
}

$baseline = Get-QaMetrics -Path $BaselineQaPath
$candidate = Get-QaMetrics -Path $CandidateQaPath

$reasons = New-Object System.Collections.Generic.List[string]
$accept = $true

if ($candidate.Critical -gt $baseline.Critical) {
    $accept = $false
    $reasons.Add("Critical 증가: $($baseline.Critical) -> $($candidate.Critical)")
}

if ($candidate.Major -gt $baseline.Major) {
    $accept = $false
    $reasons.Add("Major 증가: $($baseline.Major) -> $($candidate.Major)")
}

if ($candidate.Score -lt $baseline.Score) {
    $accept = $false
    $reasons.Add("Score 하락: $($baseline.Score) -> $($candidate.Score)")
}

$improved = (
    ($candidate.Critical -lt $baseline.Critical) -or
    ($candidate.Major -lt $baseline.Major) -or
    ($candidate.Score -gt $baseline.Score)
)

if ($accept -and -not $improved) {
    $accept = $false
    $reasons.Add("핵심 QA 지표 개선 없음")
}

$selectedPath = $BaselinePath
$selectedQaPath = $BaselineQaPath
$decision = "ROLLBACK_TO_BASELINE"

if ($accept) {
    $selectedPath = $CandidatePath
    $selectedQaPath = $CandidateQaPath
    $decision = "ACCEPT_CANDIDATE"
    $reasons.Add("Critical/Major/Score 기준에서 악화 없이 개선 확인")
}

$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ext = [System.IO.Path]::GetExtension($selectedPath)

if ([string]::IsNullOrWhiteSpace($FinalName)) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($selectedPath)
    $FinalName = "${baseName}_FINAL_${stamp}${ext}"
}
elseif ([System.IO.Path]::GetExtension($FinalName) -eq "") {
    $FinalName = $FinalName + $ext
}

$finalPath = Join-Path $outputDir $FinalName
Copy-Item -LiteralPath $selectedPath -Destination $finalPath -Force

$decisionPath = Join-Path $outputDir (
    [System.IO.Path]::GetFileNameWithoutExtension($FinalName) + "_acceptance.json"
)

$decisionObj = [PSCustomObject]@{
    decision = $decision
    accepted_candidate = $accept
    baseline = [PSCustomObject]@{
        file = $BaselinePath
        qa = $BaselineQaPath
        score = $baseline.Score
        critical = $baseline.Critical
        major = $baseline.Major
        minor = $baseline.Minor
        status = $baseline.Status
    }
    candidate = [PSCustomObject]@{
        file = $CandidatePath
        qa = $CandidateQaPath
        score = $candidate.Score
        critical = $candidate.Critical
        major = $candidate.Major
        minor = $candidate.Minor
        status = $candidate.Status
    }
    selected_file = $selectedPath
    final_file = $finalPath
    reasons = @($reasons)
    policy = [PSCustomObject]@{
        reject_if_critical_increases = $true
        reject_if_major_increases = $true
        reject_if_score_decreases = $true
        require_at_least_one_improvement = $true
        note = "현재 Gate는 content-only QA 지표를 비교하며 시각 품질은 별도 검토 대상"
    }
}

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$json = $decisionObj | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($decisionPath, $json, $utf8Bom)

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE QA ACCEPTANCE GATE v1"
Write-Host "========================================"
Write-Host ""
Write-Host "Baseline  : Score $($baseline.Score) | Critical $($baseline.Critical) | Major $($baseline.Major) | Minor $($baseline.Minor)"
Write-Host "Candidate : Score $($candidate.Score) | Critical $($candidate.Critical) | Major $($candidate.Major) | Minor $($candidate.Minor)"
Write-Host ""
Write-Host "Decision  :" $decision
Write-Host "Selected  :" $selectedPath
Write-Host "Final     :" $finalPath
Write-Host "Decision JSON :" $decisionPath
Write-Host ""
Write-Host "Reason:"
foreach ($r in $reasons) {
    Write-Host " -" $r
}
Write-Host ""
Write-Host "STATUS: QA_ACCEPTANCE_GATE_COMPLETE"
