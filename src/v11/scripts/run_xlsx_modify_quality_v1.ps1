param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string]$UserRequest
)

$ErrorActionPreference = "Stop"
$outputDir = "C:\AI\work\output"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$ext = [System.IO.Path]::GetExtension($FilePath)
$started = Get-Date

& "C:\AI\scripts\run_xlsx_modify_v1.ps1" -FilePath $FilePath -UserRequest $UserRequest
if (-not $?) { throw "XLSX modify 실패" }

$draft = @(Get-ChildItem -LiteralPath $outputDir -Filter "${baseName}_modified_*${ext}" -File |
    Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending)[0]

if ($null -eq $draft) { throw "수정된 Excel 파일 누락" }

& "C:\AI\scripts\run_office_quality_pipeline_v1.ps1" `
    -DraftPath $draft.FullName `
    -SourcePath $FilePath `
    -UserRequest $UserRequest

if (-not $?) { throw "XLSX quality pipeline 실패" }
Write-Host "STATUS: XLSX_MODIFY_QUALITY_COMPLETE"
