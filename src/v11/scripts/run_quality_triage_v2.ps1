param(
    [Parameter(Mandatory=$true)][string]$QaPath,
    [string]$SourcePath = "",
    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $QaPath -PathType Leaf)) { throw "QA JSON 누락: $QaPath" }

$tempDir = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$qaBase = [System.IO.Path]::GetFileNameWithoutExtension($QaPath)
$sourceExtractPath = ""
$instructionPath = Join-Path $tempDir "quality_triage_v2_${stamp}.txt"
$outputPath = Join-Path $outputDir "${qaBase}_triage_v2_${stamp}.json"

if (-not [string]::IsNullOrWhiteSpace($SourcePath) -and (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    $ext = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $extractor = $null

    if ($ext -in @(".pptx",".ppt")) { $extractor = "C:\AI\scripts\extract_ppt_for_qa.ps1" }
    elseif ($ext -eq ".docx") { $extractor = "C:\AI\scripts\extract_docx_for_qa.ps1" }
    elseif ($ext -in @(".xlsx",".xlsm",".xls")) { $extractor = "C:\AI\scripts\extract_xlsx_for_qa.ps1" }

    if ($null -ne $extractor -and (Test-Path -LiteralPath $extractor -PathType Leaf)) {
        $sourceExtractPath = Join-Path $tempDir "quality_triage_source_${stamp}.json"
        & $extractor -FilePath $SourcePath -OutputPath $sourceExtractPath
    }
}

$sourceSection = ""
if (-not [string]::IsNullOrWhiteSpace($sourceExtractPath)) {
    $sourceSection = "[ORIGINAL SOURCE EXTRACT]`r`n$sourceExtractPath"
}

$instruction = @"
You are Jungle QA TRIAGE v2.

Read QA:
$QaPath

Original request:
$UserRequest

$sourceSection

Use only listed local files. No web. Do not modify source artifacts.

Classify EVERY issue into exactly one:

AUTO_FIX
- safe without inventing facts/data
- wording, duplication, labels, terminology already supported
- content already present in source but omitted
- formula/value correction only if deterministically supported

FACT_DECISION
- requires unverified business decision, owner, date, target, approved basis, value, formula, or interpretation
- would be fabrication if guessed

DETAIL_OVERREACH
- reasonable in appendix/operating detail but excessive for requested main artifact

RULES:
- If source supports the missing item, AUTO_FIX is allowed.
- Never invent facts, numbers, formulas, owners, dates, or targets.
- Preserve concise executive style where applicable.
- Keep original severity/category/location.

Save UTF-8 JSON to:
$outputPath

Schema:
{
  "triage_summary":{"auto_fix":0,"fact_decision":0,"detail_overreach":0},
  "issues":[
    {
      "issue_index":1,
      "original_severity":"major",
      "original_category":"logic",
      "location":"section/sheet/cell",
      "classification":"AUTO_FIX | FACT_DECISION | DETAIL_OVERREACH",
      "reason":"concise Korean explanation",
      "safe_action":"specific action or NONE",
      "requires_user_decision":true,
      "appendix_candidate":false
    }
  ],
  "recommended_next_step":"AUTO_REVISION | ASK_USER | AUTO_REVISION_THEN_ASK | NO_CHANGE"
}

After saving output only QA_TRIAGE_V2_CREATED.
"@

$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($instructionPath,$instruction,$enc)

$req = "Read the instruction file at $instructionPath and execute it exactly. Output only QA_TRIAGE_V2_CREATED."
claude -p $req --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw "Triage v2 JSON 생성 실패" }

$t = Get-Content -LiteralPath $outputPath -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE QA TRIAGE v2 RESULT"
Write-Host "========================================"
Write-Host "AUTO_FIX        :" $t.triage_summary.auto_fix
Write-Host "FACT_DECISION   :" $t.triage_summary.fact_decision
Write-Host "DETAIL_OVERREACH:" $t.triage_summary.detail_overreach
Write-Host "Next            :" $t.recommended_next_step
Write-Host "File            :" $outputPath
Write-Host "STATUS: QA_TRIAGE_V2_COMPLETE"
