param(
    [Parameter(Mandatory=$true)]
    [string]$DraftPath,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

foreach ($p in @($DraftPath, $SourcePath)) {
    if (!(Test-Path -LiteralPath $p -PathType Leaf)) {
        throw "파일을 찾을 수 없습니다: $p"
    }
}

$qaGate       = "C:\AI\scripts\run_codex_qa_gate_v3.ps1"
$triageGate   = "C:\AI\scripts\run_qa_triage_v1.ps1"
$revisionGate = "C:\AI\scripts\run_pptx_auto_revision_v1.ps1"
$acceptGate   = "C:\AI\scripts\run_qa_acceptance_gate_v1.ps1"

foreach ($dep in @($qaGate, $triageGate, $revisionGate, $acceptGate)) {
    if (!(Test-Path -LiteralPath $dep -PathType Leaf)) {
        throw "필수 엔진을 찾을 수 없습니다: $dep"
    }
}

$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

function Get-NewestFile {
    param(
        [string]$Pattern,
        [datetime]$Since
    )

    $f = Get-ChildItem -LiteralPath $outputDir -Filter $Pattern -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-2) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    return $f
}

function Show-TriageSummary {
    param([object]$Triage)

    Write-Host "AUTO_FIX        :" $Triage.triage_summary.auto_fix
    Write-Host "FACT_DECISION   :" $Triage.triage_summary.fact_decision
    Write-Host "DETAIL_OVERREACH:" $Triage.triage_summary.detail_overreach
    Write-Host "Next            :" $Triage.recommended_next_step
}

$draftBase = [System.IO.Path]::GetFileNameWithoutExtension($DraftPath)

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT QUALITY PIPELINE v1"
Write-Host "========================================"
Write-Host ""
Write-Host "Policy : QA -> Triage -> AUTO_FIX 1회 -> Re-QA -> Acceptance/Rollback"
Write-Host "Draft  :" $DraftPath
Write-Host "Source :" $SourcePath
Write-Host ""

# ------------------------------------------------------------
# 1. Baseline QA
# ------------------------------------------------------------
$stepStart = Get-Date

& $qaGate `
    -FilePath $DraftPath `
    -UserRequest $UserRequest

$baselineQaFile = Get-NewestFile -Pattern "${draftBase}_QA_v3_*.json" -Since $stepStart

if ($null -eq $baselineQaFile) {
    throw "Baseline QA 결과 파일을 찾지 못했습니다."
}

$baselineQaPath = $baselineQaFile.FullName

# ------------------------------------------------------------
# 2. Triage
# ------------------------------------------------------------
$stepStart = Get-Date

& $triageGate `
    -QaPath $baselineQaPath `
    -SourcePath $SourcePath `
    -UserRequest $UserRequest

$qaBase = [System.IO.Path]::GetFileNameWithoutExtension($baselineQaPath)
$triageFile = Get-NewestFile -Pattern "${qaBase}_triage_*.json" -Since $stepStart

if ($null -eq $triageFile) {
    throw "Triage 결과 파일을 찾지 못했습니다."
}

$triagePath = $triageFile.FullName
$triage = Get-Content -LiteralPath $triagePath -Raw -Encoding UTF8 | ConvertFrom-Json

$autoFixCount = [int]$triage.triage_summary.auto_fix
$factCount = [int]$triage.triage_summary.fact_decision
$overCount = [int]$triage.triage_summary.detail_overreach

# ------------------------------------------------------------
# 3. No AUTO_FIX -> finalize baseline without revision
# ------------------------------------------------------------
if ($autoFixCount -eq 0) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $ext = [System.IO.Path]::GetExtension($DraftPath)
    $finalPath = Join-Path $outputDir "${draftBase}_FINAL_${stamp}${ext}"

    Copy-Item -LiteralPath $DraftPath -Destination $finalPath -Force

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE PPT QUALITY RESULT"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Decision         : KEEP_BASELINE"
    Write-Host "AUTO_FIX         : 0"
    Write-Host "FACT_DECISION    :" $factCount
    Write-Host "DETAIL_OVERREACH :" $overCount
    Write-Host "Final            :" $finalPath
    Write-Host "Baseline QA      :" $baselineQaPath
    Write-Host "Triage           :" $triagePath
    Write-Host ""
    Write-Host "STATUS: PPT_QUALITY_PIPELINE_COMPLETE"
    return
}

# ------------------------------------------------------------
# 4. AUTO_FIX revision - EXACTLY ONCE
# ------------------------------------------------------------
$stepStart = Get-Date

& $revisionGate `
    -DraftPath $DraftPath `
    -TriagePath $triagePath `
    -SourcePath $SourcePath `
    -UserRequest $UserRequest

$candidateFile = Get-NewestFile -Pattern "${draftBase}_modified_*.pptx" -Since $stepStart

if ($null -eq $candidateFile) {
    throw "AUTO_FIX 수정본을 찾지 못했습니다."
}

$candidatePath = $candidateFile.FullName
$candidateBase = [System.IO.Path]::GetFileNameWithoutExtension($candidatePath)

# ------------------------------------------------------------
# 5. Re-QA candidate
# ------------------------------------------------------------
$stepStart = Get-Date

& $qaGate `
    -FilePath $candidatePath `
    -UserRequest $UserRequest

$candidateQaFile = Get-NewestFile -Pattern "${candidateBase}_QA_v3_*.json" -Since $stepStart

if ($null -eq $candidateQaFile) {
    throw "수정본 QA 결과 파일을 찾지 못했습니다."
}

$candidateQaPath = $candidateQaFile.FullName

# ------------------------------------------------------------
# 6. Acceptance / rollback
# ------------------------------------------------------------
$stepStart = Get-Date

& $acceptGate `
    -BaselinePath $DraftPath `
    -BaselineQaPath $baselineQaPath `
    -CandidatePath $candidatePath `
    -CandidateQaPath $candidateQaPath

$acceptanceFile = Get-ChildItem -LiteralPath $outputDir -Filter "*_acceptance.json" -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -ge $stepStart.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $acceptanceFile) {
    throw "Acceptance 결과 JSON을 찾지 못했습니다."
}

$acceptancePath = $acceptanceFile.FullName
$acceptance = Get-Content -LiteralPath $acceptancePath -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT QUALITY PIPELINE RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Initial AUTO_FIX :" $autoFixCount
Write-Host "FACT_DECISION    :" $factCount
Write-Host "DETAIL_OVERREACH :" $overCount
Write-Host "Revision passes  : 1 / 1"
Write-Host "Decision         :" $acceptance.decision
Write-Host "Final            :" $acceptance.final_file
Write-Host "Baseline QA      :" $baselineQaPath
Write-Host "Candidate QA     :" $candidateQaPath
Write-Host "Triage           :" $triagePath
Write-Host "Acceptance       :" $acceptancePath
Write-Host ""
if ($factCount -gt 0) {
    Write-Host "USER DECISION NEEDED :" $factCount "FACT_DECISION issue(s)"
}
if ($overCount -gt 0) {
    Write-Host "APPENDIX/DETAIL NOTE :" $overCount "DETAIL_OVERREACH issue(s)"
}
Write-Host ""
Write-Host "STATUS: PPT_QUALITY_PIPELINE_COMPLETE"
