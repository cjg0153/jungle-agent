param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$UserRequest
)

$ErrorActionPreference = "Stop"
$outputDir = "C:\AI\work\output"
$started = Get-Date

& "C:\AI\scripts\run_docx_create.ps1" -SourcePath $SourcePath -UserRequest $UserRequest
if (-not $?) { throw "DOCX create 실패" }

$draft = @(Get-ChildItem -LiteralPath $outputDir -Filter "*.docx" -File |
    Where-Object {
        $_.LastWriteTime -ge $started.AddSeconds(-2) -and
        $_.Name -notmatch "_FINAL_"
    } |
    Sort-Object LastWriteTime -Descending)[0]

if ($null -eq $draft) { throw "생성된 DOCX 누락" }

& "C:\AI\scripts\run_office_quality_pipeline_v1.ps1" `
    -DraftPath $draft.FullName `
    -SourcePath $SourcePath `
    -UserRequest $UserRequest

if (-not $?) { throw "DOCX create quality pipeline 실패" }
Write-Host "STATUS: DOCX_CREATE_QUALITY_COMPLETE"
