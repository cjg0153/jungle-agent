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
    Write-Host "STATUS: XLSX_AUTO_REVISION_SKIPPED"
    return
}

$actions = (@($auto | ForEach-Object {
    "- 위치: $($_.location)`r`n  안전한 수정: $($_.safe_action)"
}) -join "`r`n")

$request = @"
$UserRequest

[STRICT XLSX AUTO_FIX ONLY]
아래 항목만 최소 수정:
$actions

금지:
- FACT_DECISION / DETAIL_OVERREACH 수정
- 원본에 없는 숫자/목표/사실 생성
- 근거 없는 수식 생성 또는 숫자 덮어쓰기
- 시트/행/열 구조 변경

정확한 값/수식이 확정되지 않으면 건드리지 않는다.
"@

& "C:\AI\scripts\run_xlsx_modify_v1.ps1" -FilePath $DraftPath -UserRequest $request
if (-not $?) { throw "XLSX AUTO_FIX 수정 실패" }

Write-Host "STATUS: XLSX_AUTO_REVISION_V1_COMPLETE"
