param(
    [Parameter(Mandatory=$true)][string]$DraftPath,
    [Parameter(Mandatory=$true)][string]$TriagePath,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$UserRequest = ""
)

$ErrorActionPreference = "Stop"
$triage = Get-Content -LiteralPath $TriagePath -Raw -Encoding UTF8 | ConvertFrom-Json
$auto = @($triage.issues | Where-Object { $_.classification -eq "AUTO_FIX" })

if ($auto.Count -eq 0) {
    Write-Host "STATUS: DOCX_AUTO_REVISION_SKIPPED"
    return
}

$actions = (@($auto | ForEach-Object {
    "- 위치: $($_.location)`r`n  안전한 수정: $($_.safe_action)"
}) -join "`r`n")

$request = @"
$UserRequest

[STRICT DOCX AUTO_FIX ONLY]
아래 AUTO_FIX만 최소 수정한다.
$actions

금지:
- FACT_DECISION / DETAIL_OVERREACH 수정
- 원본에 없는 숫자, 기한, 책임자, 승인범위, 사실 생성
- 문서 전체 재작성
- 불필요한 표/문단 구조 대수술

기존 형식과 구조를 최대한 보존한다.
"@

& "C:\AI\scripts\run_docx_modify.ps1" -FilePath $DraftPath -UserRequest $request
if (-not $?) { throw "DOCX AUTO_FIX 수정 실패" }

Write-Host "STATUS: DOCX_AUTO_REVISION_V1_COMPLETE"
