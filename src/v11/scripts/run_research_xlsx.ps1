param(
    [Parameter(Mandatory=$true)]
    [string]$UserRequest,

    [switch]$XHigh
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$TemplatePath = "C:\AI\prompts\research_xlsx.md"
$TempDir      = "C:\AI\work\temp"

if (!(Test-Path $TemplatePath)) {
    throw "RESEARCH + XLSX 프롬프트 템플릿이 없습니다: $TemplatePath"
}

if (!(Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$asOfDate  = Get-Date -Format "yyyy-MM-dd"

$outputJson = Join-Path $TempDir "research_xlsx_$timestamp.json"

$template = Get-Content `
    -LiteralPath $TemplatePath `
    -Raw `
    -Encoding UTF8

$codexPrompt = $template
$codexPrompt = $codexPrompt.Replace("{{USER_REQUEST}}", $UserRequest)
$codexPrompt = $codexPrompt.Replace("{{AS_OF_DATE}}", $asOfDate)
$codexPrompt = $codexPrompt.Replace("{{OUTPUT_JSON}}", $outputJson)

if ($XHigh) {
    $reasoning = "xhigh"
}
else {
    $reasoning = "medium"
}

Write-Host ""
Write-Host "=== CODEX LIVE RESEARCH ==="
Write-Host "Reasoning :" $reasoning
Write-Host "Output    :" $outputJson
Write-Host ""

Push-Location "C:\AI"

try {

    $codexPrompt | codex --search exec `
        --skip-git-repo-check `
        --sandbox workspace-write `
        --config "model_reasoning_effort=$reasoning" `
        -

}
finally {

    Pop-Location
}

if (!(Test-Path $outputJson)) {
    throw "Codex가 조사 JSON을 생성하지 못했습니다."
}

# JSON 형식 검증
$jsonText = Get-Content `
    -LiteralPath $outputJson `
    -Raw `
    -Encoding UTF8

try {
    $jsonData = $jsonText | ConvertFrom-Json
}
catch {
    throw "생성된 파일이 정상 JSON이 아닙니다: $outputJson"
}

Write-Host ""
Write-Host "=== RESEARCH RESULT ==="
Write-Host "JSON     :" $outputJson

if ($null -ne $jsonData.subjects) {
    Write-Host "Subjects :" $jsonData.subjects.Count
}

Write-Host ""
Write-Host "STATUS: RESEARCH_JSON_READY"

# ai.ps1이 경로를 받을 수 있도록 마지막에 반환
$outputJson
