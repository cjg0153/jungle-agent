param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [string]$UserRequest = "",
    [ValidateSet("low","medium","high")][string]$Reasoning = "medium"
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "검수할 파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$artifact = $null
$extractor = $null
$artifactChecks = ""

if ($ext -eq ".docx") {
    $artifact = "DOCX"
    $extractor = "C:\AI\scripts\extract_docx_for_qa.ps1"
    $artifactChecks = @"
DOCX-SPECIFIC:
- heading/section flow
- duplicated or contradictory paragraphs
- table/body consistency
- executive readability
- residual template/source wording
"@
}
elseif ($ext -in @(".xlsx",".xlsm",".xls")) {
    $artifact = "XLSX"
    $extractor = "C:\AI\scripts\extract_xlsx_for_qa.ps1"
    $artifactChecks = @"
XLSX-SPECIFIC:
- formula errors (#REF!, #DIV/0!, #VALUE!, etc.)
- inconsistent formulas in a logical row/column
- suspicious hard-coded values inside formula regions
- inconsistent units/number formats
- totals/subtotals that are ambiguous or unsupported
- blank/unknown values incorrectly presented as zero
- unclear sheet/column/metric labels
"@
}
else {
    throw "Office QA v1 지원 형식이 아닙니다: $ext"
}

if (!(Test-Path -LiteralPath $extractor -PathType Leaf)) { throw "추출기 누락: $extractor" }

$codexPath = "C:\Users\Zephyrus G14\AppData\Local\Programs\OpenAI\Codex\bin\codex.exe"
if (!(Test-Path -LiteralPath $codexPath -PathType Leaf)) { throw "Codex CLI 누락: $codexPath" }

$tempDir = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$extractPath = Join-Path $tempDir "office_qa_extract_${stamp}.json"
$sourcePath = Join-Path $tempDir "office_qa_source_${stamp}.txt"
$rawPath = Join-Path $tempDir "office_qa_raw_${stamp}.txt"
$stderrPath = Join-Path $tempDir "office_qa_stderr_${stamp}.txt"
$qaPath = Join-Path $outputDir "${baseName}_QA_office_v1_${stamp}.json"

& $extractor -FilePath $FilePath -OutputPath $extractPath
$extract = Get-Content -LiteralPath $extractPath -Raw -Encoding UTF8 | ConvertFrom-Json

$source = @"
[OFFICE QA SOURCE]
Artifact: $artifact
Source: $FilePath

[COVERAGE]
$($extract.coverage | ConvertTo-Json -Compress)

[CONTENT]
$($extract.all_qa_text)
"@

$encBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($sourcePath,$source,$encBom)

$instruction = @"
You are Jungle's final CONTENT/DATA QA gate.

Artifact: $artifact
The original artifact path is metadata only. DO NOT open it.
Read ONLY:
$sourcePath

Original user request:
$UserRequest

MANDATORY SKILLS:
1. Apply installed skill "ax-deliverables".
2. Apply installed skill "humanizer-kr".
3. If unavailable, mark MISSING.

STRICT SCOPE:
- Do not open the original Office file.
- Do not use web search/external sources.
- Do not modify files.
- At most 12 issues.
- Distinguish "not visible in extracted source" from "definitely absent".
- Never invent business facts, dates, owners, values, formulas, or targets.
- No chain-of-thought.

GENERAL CHECKS:
- logical gaps / contradictions
- duplicated messages
- inconsistent terminology, dates, units, labels, numbers
- grounding
- executive readability
- awkward/formulaic AI-style Korean
- unclear responsibility/action/decision
- template/source residue

$artifactChecks

Return ONLY one valid JSON object:
{
  "overall_status": "PASS | REVISE | FAIL",
  "executive_readiness_score": 0,
  "review_scope": "content_data_only_$($artifact.ToLowerInvariant())",
  "visual_review_status": "NOT_PERFORMED",
  "skills": {
    "ax-deliverables": {"status":"USED | MISSING","note":""},
    "humanizer-kr": {"status":"USED | MISSING","note":""}
  },
  "issue_counts": {"critical":0,"major":0,"minor":0},
  "issues": [
    {
      "severity":"critical | major | minor",
      "category":"logic | duplication | consistency | grounding | executive_readability | humanizer | data_integrity | formula | template_leak | other",
      "location":"section/sheet/cell",
      "current_text":"short excerpt",
      "problem":"concise explanation",
      "recommendation":"specific safe revision direction"
    }
  ],
  "humanizer_summary":{"ai_style_risk":"low | medium | high","patterns":[]},
  "revision_required":true,
  "final_summary":"2-4 concise Korean sentences"
}

SCORING:
90-100 delivery-ready except minor polish
80-89 good but revise before important delivery
70-79 substantial revision needed
below 70 not delivery-ready

PASS requires no critical/major and score >= 90.
"@

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $codexPath
$psi.Arguments = "exec --skip-git-repo-check --sandbox read-only --config model_reasoning_effort=$Reasoning -"
$psi.UseShellExecute = $false
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.CreateNoWindow = $true

try {
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $psi.StandardErrorEncoding = New-Object System.Text.UTF8Encoding($false)
} catch {}

$p = New-Object System.Diagnostics.Process
$p.StartInfo = $psi

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE OFFICE CODEX QA v1"
Write-Host "========================================"
Write-Host "Artifact  :" $artifact
Write-Host "Reasoning :" $Reasoning
Write-Host "Codex QA started..."

[void]$p.Start()
$stdoutTask = $p.StandardOutput.ReadToEndAsync()
$stderrTask = $p.StandardError.ReadToEndAsync()

$enc = New-Object System.Text.UTF8Encoding($false)
$bytes = $enc.GetBytes($instruction)
$p.StandardInput.BaseStream.Write($bytes,0,$bytes.Length)
$p.StandardInput.BaseStream.Flush()
$p.StandardInput.Close()

$started = Get-Date
$nextProgress = 15
while (-not $p.HasExited) {
    Start-Sleep -Seconds 1
    $elapsed = [int]((Get-Date)-$started).TotalSeconds
    if ($elapsed -ge $nextProgress) {
        Write-Host ("Codex QA running... {0}s" -f $elapsed)
        $nextProgress += 15
    }
}

$p.WaitForExit()
$raw = $stdoutTask.Result
$stderr = $stderrTask.Result

[System.IO.File]::WriteAllText($rawPath,$raw,$encBom)
[System.IO.File]::WriteAllText($stderrPath,$stderr,$encBom)

if ($p.ExitCode -ne 0) { throw "Codex QA 실패. stderr: $stderrPath" }

$first = $raw.IndexOf("{")
$last = $raw.LastIndexOf("}")
if ($first -lt 0 -or $last -le $first) { throw "Codex QA JSON을 찾지 못했습니다. Raw: $rawPath" }

$qa = $raw.Substring($first,$last-$first+1) | ConvertFrom-Json
[System.IO.File]::WriteAllText($qaPath,($qa | ConvertTo-Json -Depth 12),$encBom)

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE OFFICE QA RESULT"
Write-Host "========================================"
Write-Host "Status   :" $qa.overall_status
Write-Host "Score    :" $qa.executive_readiness_score
Write-Host "Critical :" $qa.issue_counts.critical
Write-Host "Major    :" $qa.issue_counts.major
Write-Host "Minor    :" $qa.issue_counts.minor
Write-Host "QA File  :" $qaPath
Write-Host ""
Write-Host "Summary:"
Write-Host $qa.final_summary
Write-Host "STATUS: OFFICE_CODEX_QA_V1_COMPLETE"
