param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string]$UserRequest
)

$ErrorActionPreference = "Stop"
$outputDir = "C:\AI\work\output"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$started = Get-Date

& "C:\AI\scripts\run_pptx_modify.ps1" -FilePath $FilePath -UserRequest $UserRequest
if (-not $?) { throw "PPTX modify 실패" }

$draft = @(Get-ChildItem -LiteralPath $outputDir -Filter "${baseName}_modified_*.pptx" -File |
    Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending)[0]

if ($null -eq $draft) { throw "수정된 PPTX 누락" }

& "C:\AI\scripts\run_ppt_quality_pipeline_v1.ps1" `
    -DraftPath $draft.FullName `
    -SourcePath $FilePath `
    -UserRequest $UserRequest

if (-not $?) { throw "PPT MODIFY quality pipeline 실패" }
Write-Host "STATUS: PPTX_MODIFY_QUALITY_COMPLETE"
