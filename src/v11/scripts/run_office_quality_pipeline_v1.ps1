param(
    [Parameter(Mandatory=$true)][string]$DraftPath,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

foreach ($f in @($DraftPath,$SourcePath)) {
    if (!(Test-Path -LiteralPath $f -PathType Leaf)) { throw "파일 누락: $f" }
}

$ext = [System.IO.Path]::GetExtension($DraftPath).ToLowerInvariant()
$outputDir = "C:\AI\work\output"

if ($ext -eq ".docx") {
    $autoRevision = "C:\AI\scripts\run_docx_auto_revision_v1.ps1"
    $candidatePattern = ([System.IO.Path]::GetFileNameWithoutExtension($DraftPath)) + "_modified_*.docx"
}
elseif ($ext -in @(".xlsx",".xlsm",".xls")) {
    $autoRevision = "C:\AI\scripts\run_xlsx_auto_revision_v1.ps1"
    $candidatePattern = ([System.IO.Path]::GetFileNameWithoutExtension($DraftPath)) + "_modified_*" + $ext
}
else {
    throw "Office quality pipeline 미지원 형식: $ext"
}

$qaGate = "C:\AI\scripts\run_office_codex_qa_v1.ps1"
$triageGate = "C:\AI\scripts\run_quality_triage_v2.ps1"
$acceptGate = "C:\AI\scripts\run_qa_acceptance_gate_v1.ps1"

foreach ($d in @($qaGate,$triageGate,$autoRevision,$acceptGate)) {
    if (!(Test-Path -LiteralPath $d -PathType Leaf)) { throw "필수 엔진 누락: $d" }
}

function Get-NewestOutput {
    param([string]$Pattern,[datetime]$Since)
    @(Get-ChildItem -LiteralPath $outputDir -Filter $Pattern -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-2) } |
        Sort-Object LastWriteTime -Descending)[0]
}

$draftBase = [System.IO.Path]::GetFileNameWithoutExtension($DraftPath)

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE OFFICE QUALITY PIPELINE v1"
Write-Host "========================================"
Write-Host "Draft :" $DraftPath
Write-Host "Type  :" $ext

$step = Get-Date
& $qaGate -FilePath $DraftPath -UserRequest $UserRequest
$baselineQa = Get-NewestOutput -Pattern "${draftBase}_QA_office_v1_*.json" -Since $step
if ($null -eq $baselineQa) { throw "Baseline QA 결과 누락" }

$step = Get-Date
& $triageGate -QaPath $baselineQa.FullName -SourcePath $SourcePath -UserRequest $UserRequest
$qaBase = [System.IO.Path]::GetFileNameWithoutExtension($baselineQa.FullName)
$triageFile = Get-NewestOutput -Pattern "${qaBase}_triage_v2_*.json" -Since $step
if ($null -eq $triageFile) { throw "Triage 결과 누락" }

$t = Get-Content -LiteralPath $triageFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
$autoCount = [int]$t.triage_summary.auto_fix
$factCount = [int]$t.triage_summary.fact_decision
$overCount = [int]$t.triage_summary.detail_overreach

if ($autoCount -eq 0) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $finalPath = Join-Path $outputDir "${draftBase}_FINAL_${stamp}${ext}"
    Copy-Item -LiteralPath $DraftPath -Destination $finalPath -Force

    Write-Host ""
    Write-Host "Decision         : KEEP_BASELINE"
    Write-Host "AUTO_FIX         : 0"
    Write-Host "FACT_DECISION    :" $factCount
    Write-Host "DETAIL_OVERREACH :" $overCount
    Write-Host "Final            :" $finalPath
    Write-Host "STATUS: OFFICE_QUALITY_PIPELINE_COMPLETE"
    return
}

$step = Get-Date
& $autoRevision -DraftPath $DraftPath -TriagePath $triageFile.FullName -SourcePath $SourcePath -UserRequest $UserRequest
if (-not $?) { throw "AUTO_FIX pass 실패" }

$candidate = Get-NewestOutput -Pattern $candidatePattern -Since $step
if ($null -eq $candidate) { throw "AUTO_FIX 수정본 누락" }

$candidateBase = [System.IO.Path]::GetFileNameWithoutExtension($candidate.FullName)
$step = Get-Date
& $qaGate -FilePath $candidate.FullName -UserRequest $UserRequest
$candidateQa = Get-NewestOutput -Pattern "${candidateBase}_QA_office_v1_*.json" -Since $step
if ($null -eq $candidateQa) { throw "Candidate QA 결과 누락" }

$step = Get-Date
& $acceptGate `
    -BaselinePath $DraftPath `
    -BaselineQaPath $baselineQa.FullName `
    -CandidatePath $candidate.FullName `
    -CandidateQaPath $candidateQa.FullName

$acceptance = @(Get-ChildItem -LiteralPath $outputDir -Filter "*_acceptance.json" -File |
    Where-Object { $_.LastWriteTime -ge $step.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending)[0]

if ($null -eq $acceptance) { throw "Acceptance JSON 누락" }

$a = Get-Content -LiteralPath $acceptance.FullName -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE OFFICE QUALITY RESULT"
Write-Host "========================================"
Write-Host "Initial AUTO_FIX :" $autoCount
Write-Host "FACT_DECISION    :" $factCount
Write-Host "DETAIL_OVERREACH :" $overCount
Write-Host "Revision passes  : 1 / 1"
Write-Host "Decision         :" $a.decision
Write-Host "Final            :" $a.final_file
Write-Host "Baseline QA      :" $baselineQa.FullName
Write-Host "Candidate QA     :" $candidateQa.FullName
Write-Host "Triage           :" $triageFile.FullName
Write-Host "Acceptance       :" $acceptance.FullName
Write-Host "STATUS: OFFICE_QUALITY_PIPELINE_COMPLETE"
