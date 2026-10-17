param(
    [Parameter(Mandatory=$true)]
    [string]$QaPath,

    [string]$SourcePath = "",

    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $QaPath -PathType Leaf)) {
    throw "QA JSON을 찾을 수 없습니다: $QaPath"
}

$tempDir   = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$qaBase = [System.IO.Path]::GetFileNameWithoutExtension($QaPath)

$sourceExtractPath = ""
$instructionPath = Join-Path $tempDir "qa_triage_instruction_${stamp}.txt"
$outputPath = Join-Path $outputDir "${qaBase}_triage_${stamp}.json"

if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
    if (!(Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "원본 참고자료를 찾을 수 없습니다: $SourcePath"
    }

    $extractor = "C:\AI\scripts\extract_ppt_for_qa.ps1"

    if (!(Test-Path -LiteralPath $extractor -PathType Leaf)) {
        throw "PPT QA 추출기를 찾을 수 없습니다: $extractor"
    }

    $sourceExtractPath = Join-Path $tempDir "qa_triage_source_${stamp}.json"

    & $extractor -FilePath $SourcePath -OutputPath $sourceExtractPath

    if (!(Test-Path -LiteralPath $sourceExtractPath -PathType Leaf)) {
        throw "원본 참고자료 추출에 실패했습니다."
    }
}

$sourceSection = ""

if (-not [string]::IsNullOrWhiteSpace($sourceExtractPath)) {
    $sourceSection = @"

[ORIGINAL SOURCE EXTRACT]
$sourceExtractPath
"@
}

$instruction = @"
You are the QA TRIAGE stage for Jungle.

Your job is NOT to revise the PowerPoint.
Your job is to classify each QA issue so later automation knows what may be fixed automatically and what must not be invented.

[QA RESULT]
Read this JSON:
$QaPath

[ORIGINAL USER REQUEST]
$UserRequest
$sourceSection

Use only the local files listed above.
Do not use web search, connectors, Google Drive, or external sources.
Do not modify any source artifact.

CLASSIFY EVERY QA ISSUE INTO EXACTLY ONE OF THESE:

1. AUTO_FIX
Use when the issue can be safely improved without inventing new business facts.
Examples:
- awkward Korean
- duplicated wording
- inconsistent terminology already resolvable from source
- overly long titles
- unclear labels
- redundant notes
- obvious template residue
- wording that can be narrowed to match available evidence

2. FACT_DECISION
Use when fixing the issue requires a business definition, owner, target, date, numeric basis, approved source, or decision that may not exist in the material.
Examples:
- whether "100 Agent" means candidates, pilots, builds, or operations
- exact FTE calculation basis when not fully documented
- accountable department/owner not present in source
- unapproved deadlines or numeric targets
- information that would become fabrication if guessed

3. DETAIL_OVERREACH
Use when the QA request is reasonable in a detailed appendix or operating document but too detailed for the stated executive/main-deck purpose.
Examples:
- full formula definitions in a 5-slide executive deck
- detailed WBS
- full sample sizes / observation windows / numerator-denominator definitions in main body
- implementation-level procedures better suited to appendix

IMPORTANT RULES:
- Check the original source extract when available.
- If the requested fact already exists in the source, do NOT classify it as FACT_DECISION merely because it is absent from the generated deck. It may be AUTO_FIX because the fact can be restored from source.
- If only part of the fact exists, classify FACT_DECISION only for the unsupported remainder.
- Preserve executive-level brevity.
- Never invent missing facts.
- Classify ALL issues from the QA JSON.
- Keep the original severity/category/location so later scripts can trace each issue.

Return valid UTF-8 JSON only with this schema:

{
  "triage_summary": {
    "auto_fix": 0,
    "fact_decision": 0,
    "detail_overreach": 0
  },
  "issues": [
    {
      "issue_index": 1,
      "original_severity": "major",
      "original_category": "logic",
      "location": "slide/section",
      "classification": "AUTO_FIX | FACT_DECISION | DETAIL_OVERREACH",
      "reason": "concise Korean explanation",
      "safe_action": "specific action Jungle may take automatically, or NONE",
      "requires_user_decision": true,
      "appendix_candidate": false
    }
  ],
  "recommended_next_step": "AUTO_REVISION | ASK_USER | AUTO_REVISION_THEN_ASK | NO_CHANGE"
}

Decision guidance:
- AUTO_REVISION: only AUTO_FIX issues remain.
- ASK_USER: no safe auto-fixes, but FACT_DECISION issues remain.
- AUTO_REVISION_THEN_ASK: AUTO_FIX exists and FACT_DECISION also remains.
- NO_CHANGE: only DETAIL_OVERREACH remains or no meaningful issues remain.

Save the JSON to:
$outputPath

After saving, output only:
QA_TRIAGE_CREATED
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText(
    $instructionPath,
    $instruction,
    $utf8Bom
)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified there. Do not use web search, Google Drive, connectors, or external sources. When finished, output only QA_TRIAGE_CREATED."

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE QA TRIAGE v1"
Write-Host "========================================"
Write-Host ""
Write-Host "QA     :" $QaPath
if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
    Write-Host "Source :" $SourcePath
}
Write-Host ""
Write-Host "Claude triage started..."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "QA triage JSON이 생성되지 않았습니다."
}

$triage = Get-Content -LiteralPath $outputPath -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE QA TRIAGE RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "AUTO_FIX        :" $triage.triage_summary.auto_fix
Write-Host "FACT_DECISION   :" $triage.triage_summary.fact_decision
Write-Host "DETAIL_OVERREACH:" $triage.triage_summary.detail_overreach
Write-Host "Next            :" $triage.recommended_next_step
Write-Host "File            :" $outputPath
Write-Host ""
Write-Host "STATUS: QA_TRIAGE_COMPLETE"
