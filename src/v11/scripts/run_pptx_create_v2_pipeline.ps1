param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    throw "원본 PPTX를 찾을 수 없습니다: $SourcePath"
}

$cacheDir = "C:\AI\work\temp\ppt_layout_cache"
New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null

$hash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
$designPath = Join-Path $cacheDir "${hash}_design.json"
$layoutPath = Join-Path $cacheDir "${hash}_layout.json"

$analyzer = "C:\AI\scripts\analyze_ppt_design.ps1"
$builder  = "C:\AI\scripts\build_ppt_layout_library.ps1"
$runner   = "C:\AI\scripts\run_pptx_create_v2.ps1"

if (!(Test-Path -LiteralPath $designPath -PathType Leaf)) {
    Write-Host ""
    Write-Host "PPT Design Cache : MISS"
    Write-Host "Step 1/3 : Analyze source design"
    & $analyzer -FilePath $SourcePath -OutputPath $designPath
    if (-not $?) { throw "PPT design analysis failed." }
}
else {
    Write-Host ""
    Write-Host "PPT Design Cache : HIT (analysis)"
}

if (!(Test-Path -LiteralPath $layoutPath -PathType Leaf)) {
    Write-Host "Step 2/3 : Build reusable layout library"
    & $builder -DesignJsonPath $designPath -OutputPath $layoutPath
    if (-not $?) { throw "PPT layout library build failed." }
}
else {
    Write-Host "PPT Design Cache : HIT (layout library)"
}

Write-Host "Step 3/3 : Create design-aware PPT"
& $runner -SourcePath $SourcePath -LayoutLibraryPath $layoutPath -UserRequest $UserRequest

if (-not $?) {
    throw "PPT Design v2 creation failed."
}

Write-Host ""
Write-Host "STATUS: PPT_DESIGN_V2_PIPELINE_COMPLETE"
