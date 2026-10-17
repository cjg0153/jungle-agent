param(
    [Parameter(Mandatory=$true)]
    [string]$Request
)

$ErrorActionPreference = 'Stop'

$GeneralPath = "C:\AI\jungle_v11.ps1"
$FastPath    = "C:\AI\run_verified_text_v3_fast.ps1"

$ExpectedGeneralSha = "694E1B8DC498E15D3563216F083AA9178F4FB7B5880473F9DDF4503604F819EF"
$ExpectedFastSha    = "A74D160CA96249143620970ED25014C1B302034C1248759266A50742431C36A0"

$TempRoot   = "C:\AI\work\temp"
$OutputRoot = "C:\AI\work\output"

function Get-JungleSha256 {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Dependency missing: $Path"
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Assert-JungleDependency {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Expected
    )

    $actual = Get-JungleSha256 $Path

    if ($actual -ne $Expected) {
        throw "FAIL-CLOSED: $Name dependency hash changed. Expected=$Expected Actual=$actual"
    }
}

Assert-JungleDependency `
    -Name "GENERAL_V11" `
    -Path $GeneralPath `
    -Expected $ExpectedGeneralSha

Assert-JungleDependency `
    -Name "FAST" `
    -Path $FastPath `
    -Expected $ExpectedFastSha

$r = $Request.Trim()

if ([string]::IsNullOrWhiteSpace($r)) {
    throw "Planner request is empty."
}

# ============================================================
# ARTIFACT FORMAT PRIORITY
#
# 1. Explicit user format ALWAYS wins.
# 2. If format is omitted, Jungle recommends a sensible format.
# 3. Unsupported creation types FAIL-CLOSED instead of pretending
#    that a TXT response is the requested artifact.
# ============================================================

$artifactType = "TEXT_RESPONSE"
$reason       = "No durable file artifact is clearly required."
$extension    = $null
$mermaidMode  = $false

# ------------------------------------------------------------
# Explicit formats: user choice has highest priority.
# ------------------------------------------------------------

if ($r -match '(?i)\bmermaid\b|머메이드') {

    $artifactType = "MARKDOWN_MERMAID"
    $extension    = ".md"
    $mermaidMode  = $true
    $reason       = "The user explicitly requested Mermaid; Markdown is the native maintainable container."

}
elseif ($r -match '(?i)\bhtml\b|HTML\s*파일|웹\s*페이지|웹페이지') {

    $artifactType = "HTML"
    $extension    = ".html"
    $reason       = "The user explicitly requested HTML/web output."

}
elseif ($r -match '(?i)\bmarkdown\b|마크다운|\bmd\s*파일\b') {

    $artifactType = "MARKDOWN"
    $extension    = ".md"
    $reason       = "The user explicitly requested Markdown."

}
elseif (
    $r -match '(?i)\bpptx?\b|프레젠테이션|파워포인트'
) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Intent   : CREATE_ARTIFACT"
    Write-Host "Format   : PPTX"
    Write-Host "Decision : NOT_ENABLED_IN_V11"
    Write-Host ""

    throw "FAIL-CLOSED: Fileless PPTX creation is not yet enabled in Planner v1.1. Jungle will not substitute a TXT answer."

}
elseif (
    $r -match '(?i)\bdocx\b|워드\s*문서|word\s*document'
) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Intent   : CREATE_ARTIFACT"
    Write-Host "Format   : DOCX"
    Write-Host "Decision : NOT_ENABLED_IN_V11"
    Write-Host ""

    throw "FAIL-CLOSED: Fileless DOCX creation is not yet enabled in Planner v1.1. Jungle will not substitute a TXT answer."

}
elseif (
    $r -match '(?i)\bxlsx?\b|엑셀|스프레드시트|spreadsheet'
) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Intent   : CREATE_ARTIFACT"
    Write-Host "Format   : XLSX"
    Write-Host "Decision : NOT_ENABLED_IN_V11"
    Write-Host ""

    throw "FAIL-CLOSED: Fileless XLSX creation is not yet enabled in Planner v1.1. Jungle will not substitute a TXT answer."

}
elseif (
    $r -match '(?i)\bpdf\b|PDF\s*파일'
) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Intent   : CREATE_ARTIFACT"
    Write-Host "Format   : PDF"
    Write-Host "Decision : NOT_ENABLED_IN_V11"
    Write-Host ""

    throw "FAIL-CLOSED: Fileless PDF creation is not yet enabled in Planner v1.1. Jungle will not substitute a TXT answer."

}
elseif (
    $r -match '(?i)프로그램|애플리케이션|어플리케이션|\bapp\b|앱\s*(?:만들|제작|구현)|프로젝트\s*(?:만들|제작|구현)'
) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Intent   : BUILD_PROJECT"
    Write-Host "Format   : CODE_PROJECT"
    Write-Host "Decision : RESERVED_FOR_SMART_VNEXT"
    Write-Host ""

    throw "FAIL-CLOSED: Autonomous app/project creation is reserved for a later Jungle Smart capability. Jungle will not substitute a TXT answer."

}

# ------------------------------------------------------------
# Jungle recommended formats when user did NOT specify one.
# ------------------------------------------------------------

elseif (
    $r -match '(?i)구조도|아키텍처|아키텍쳐|다이어그램|diagram|flowchart|플로우차트|시각화|그림으로|도식(?:화)?'
) {

    $artifactType = "HTML"
    $extension    = ".html"
    $reason       = "Jungle recommends HTML because the request is primarily visual/diagrammatic and should be immediately viewable."

}
elseif (
    $r -match '(?i)위키|wiki|설계서|명세서|기록(?:해|으로)|문서화|가이드|매뉴얼|회의록|정리문서|보고서'
) {

    $artifactType = "MARKDOWN"
    $extension    = ".md"
    $reason       = "Jungle recommends Markdown because the result is durable documentation that should remain easy to edit and version."

}

# ------------------------------------------------------------
# Creation words alone do NOT automatically force a file.
# Example: "문장 만들어줘" can remain a normal text response.
# ------------------------------------------------------------

else {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " JUNGLE INTENT PLAN" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Intent   : TEXT_RESPONSE"
    Write-Host "Format   : Terminal / Verified TXT Log"
    Write-Host "Reason   : $reason"
    Write-Host "Engine   : FAST Verified"
    Write-Host ""

    & $FastPath -Request $r

    if (-not $?) {
        throw "FAST execution failed."
    }

    return
}

# ============================================================
# DURABLE ARTIFACT PLAN
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " JUNGLE ARTIFACT PLAN" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Intent   : CREATE_ARTIFACT"
Write-Host "Format   : $artifactType"
Write-Host "Reason   : $reason"
Write-Host "Creator  : Claude via General File v1.1"
Write-Host "QA       : Codex"
Write-Host "Contract : Actual file required"
Write-Host ""

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runId = [Guid]::NewGuid().ToString("N").Substring(0,8)

$workDir = Join-Path `
    $TempRoot `
    ("planner_v11_{0}_{1}" -f $stamp, $runId)

New-Item `
    -ItemType Directory `
    -Path $workDir `
    -Force | Out-Null

$baseName = "Jungle_Artifact_" + $stamp
$source   = Join-Path $workDir ($baseName + $extension)

$enc = New-Object System.Text.UTF8Encoding($true)

if ($extension -eq ".html") {

    $scaffold = @"
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Jungle Artifact</title>
</head>
<body>
JUNGLE_ARTIFACT_PLACEHOLDER
</body>
</html>
"@

}
else {

    $scaffold = @"
# Jungle Artifact

JUNGLE_ARTIFACT_PLACEHOLDER
"@

}

[System.IO.File]::WriteAllText(
    $source,
    $scaffold,
    $enc
)

$sourceShaBefore = Get-JungleSha256 $source
$started         = Get-Date

$formatGuidance = ""

if ($artifactType -eq "HTML") {

    $formatGuidance = @"
Jungle selected HTML as the primary artifact format.

Artifact guidance:
- Produce a polished, self-contained HTML file that opens directly in a browser.
- Do not return only an explanation or review.
- For architecture/diagram requests, visually express hierarchy, flow, responsibilities, and relationships.
- Use professional layout, readable typography, clear grouping, and restrained motion only when it improves comprehension.
- The deliverable itself must contain the requested information.
- Remove JUNGLE_ARTIFACT_PLACEHOLDER completely.
"@

}
elseif ($artifactType -eq "MARKDOWN_MERMAID") {

    $formatGuidance = @"
Jungle selected Markdown + Mermaid because the user explicitly requested Mermaid.

Artifact guidance:
- Produce a maintainable Markdown file.
- Include at least one valid fenced Mermaid block using ```mermaid.
- The Mermaid diagram must represent the user's requested structure/process.
- Add concise explanatory text only where it helps interpretation.
- Do not return only an explanation.
- Remove JUNGLE_ARTIFACT_PLACEHOLDER completely.
"@

}
else {

    $formatGuidance = @"
Jungle selected Markdown for durable documentation.

Artifact guidance:
- Produce a clean, maintainable Markdown document.
- Use a logical heading hierarchy.
- Prefer concise tables or Mermaid only when they genuinely improve understanding.
- Do not return only an explanation.
- Remove JUNGLE_ARTIFACT_PLACEHOLDER completely.
"@

}

$task = @"
$source

위 파일은 Jungle Intent / Artifact Planner가 생성한 임시 scaffold다.

사용자의 원래 요청:
[$r]

Jungle Artifact Plan:
- Intent: CREATE_ARTIFACT
- Format: $artifactType
- Reason: $reason
- Actual file creation is mandatory.

$formatGuidance

중요:
- 위 scaffold 파일을 실제 완성 산출물로 수정해줘.
- 설명문이나 검토 의견만 반환하면 실패다.
- 사용자 요청이 지정한 형식이 있다면 그것이 최우선이다.
- 사용자 요청에 없는 사실, 수치, 경로, 기능을 임의로 발명하지 마라.
- 사용자가 형식을 지정하지 않은 경우 위 Jungle Artifact Plan을 기본 가이드로 사용하되,
  내용 구성과 디자인/표현 방식은 Claude가 자율적으로 최적화해도 된다.
- 반드시 실제 파일을 완성해야 한다.
"@

Write-Host "Scaffold : $source"
Write-Host "Executing verified General File v1.1..."
Write-Host ""

& $GeneralPath $task

if (-not $?) {
    throw "General File v1.1 execution failed."
}

# ============================================================
# SOURCE PROTECTION
# ============================================================

if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "ARTIFACT CONTRACT FAIL: planner scaffold disappeared."
}

$sourceShaAfter = Get-JungleSha256 $source

if ($sourceShaBefore -ne $sourceShaAfter) {
    throw "ARTIFACT CONTRACT FAIL: planner scaffold was modified directly."
}

# ============================================================
# VERIFIED ARTIFACT DISCOVERY
# ============================================================

$artifacts = @(
    Get-ChildItem `
        -LiteralPath $OutputRoot `
        -File `
        -ErrorAction SilentlyContinue |
    Where-Object {
        $_.BaseName -like ($baseName + "_jungle_v11_*") `
        -and $_.Extension.ToLowerInvariant() -eq $extension `
        -and $_.LastWriteTime -ge $started.AddMinutes(-1)
    } |
    Sort-Object LastWriteTime -Descending
)

if ($artifacts.Count -lt 1) {
    throw "ARTIFACT CONTRACT FAIL: verified artifact was not created."
}

$artifact = $artifacts[0]

$content = [System.IO.File]::ReadAllText(
    $artifact.FullName
)

if ($content.Contains("JUNGLE_ARTIFACT_PLACEHOLDER")) {
    throw "ARTIFACT CONTRACT FAIL: placeholder remains."
}

if ($artifact.Length -lt 150) {
    throw "ARTIFACT CONTRACT FAIL: artifact is suspiciously small."
}

if ($extension -eq ".html") {

    if (
        $content -notmatch '(?i)<!DOCTYPE\s+html' `
        -or $content -notmatch '(?i)<html'
    ) {
        throw "ARTIFACT CONTRACT FAIL: HTML structure missing."
    }

}

if ($mermaidMode) {

    if ($content -notmatch '(?i)```mermaid') {
        throw "ARTIFACT CONTRACT FAIL: Mermaid block missing."
    }

}

$artifactSha = Get-JungleSha256 $artifact.FullName

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " JUNGLE VERIFIED ARTIFACT" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Source unchanged : TRUE" -ForegroundColor Green
Write-Host "Artifact         : $($artifact.FullName)"
Write-Host "Artifact SHA256  : $artifactSha"
Write-Host "Artifact Size    : $($artifact.Length) bytes"
Write-Host "Artifact Type    : $artifactType"
Write-Host ""
Write-Host "STATUS : JUNGLE_INTENT_ARTIFACT_V11_COMPLETE" -ForegroundColor Cyan