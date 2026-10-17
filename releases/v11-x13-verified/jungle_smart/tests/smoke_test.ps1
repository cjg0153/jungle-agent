$ErrorActionPreference =
    "Stop"


$root =
    Split-Path `
        -Parent `
        (Split-Path `
            -Parent `
            $MyInvocation.MyCommand.Path)


$stableManifest =
    "C:\AI\stable_candidates\Jungle_v32_SAFE_CANDIDATE_R2_20261011_185738\FREEZE_MANIFEST.txt"


if (-not (
    Test-Path `
        -LiteralPath $stableManifest `
        -PathType Leaf
)) {

    throw "Stable manifest missing."
}


$stableHashBefore =
    (
        Get-FileHash `
            -LiteralPath $stableManifest `
            -Algorithm SHA256
    ).Hash


$required =
    @(
        "jungle_smart.ps1",
        "config\jungle.smart.json",
        "modules\router.ps1",
        "modules\state.ps1",
        "modules\artifact.ps1",
        "modules\verifier.ps1",
        "modules\watchdog.ps1",
        "modules\worker.ps1",
        "providers\claude.ps1",
        "providers\codex.ps1",
        "executors\ppt_safe_bridge.ps1",
        "PHASE_B_ARCHITECTURE.md"
    )


foreach ($relative in $required) {

    $path =
        Join-Path `
            $root `
            $relative


    if (-not (
        Test-Path `
            -LiteralPath $path `
            -PathType Leaf
    )) {

        throw "Required Jungle Smart file missing: $relative"
    }
}


$psFiles =
    @(
        Get-ChildItem `
            -LiteralPath $root `
            -Filter "*.ps1" `
            -File `
            -Recurse
    )


foreach ($file in $psFiles) {

    $text =
        [System.IO.File]::ReadAllText(
            $file.FullName
        )


    [scriptblock]::Create(
        $text
    ) | Out-Null
}


$config =
    Get-Content `
        -LiteralPath (
            Join-Path `
                $root `
                "config\jungle.smart.json"
        ) `
        -Raw |
    ConvertFrom-Json


if ($config.feature_flags.execution_enabled) {

    throw "Execution must be disabled in Phase B v0.1 baseline."
}


$request =
    'C:\AI\sample\board_report.pptx 디자인과 레이아웃을 안전하게 정돈해줘'


$output =
    (
        & (
            Join-Path `
                $root `
                "jungle_smart.ps1"
        ) `
            -Request $request `
            -DryRun *>&1 |
        Out-String
    )


if (-not $output.Contains(
    "ROUTE     : PPT_SAFE"
)) {

    throw "Router smoke test failed."
}


if (-not $output.Contains(
    "PROVIDER  : CLAUDE"
)) {

    throw "Claude routing smoke test failed."
}


if (-not $output.Contains(
    "AI CALLS  : 0"
)) {

    throw "Unexpected AI call behavior."
}


if (-not $output.Contains(
    "STATUS : JUNGLE_SMART_DRY_RUN_COMPLETE"
)) {

    throw "Dry-run completion marker missing."
}


$stableHashAfter =
    (
        Get-FileHash `
            -LiteralPath $stableManifest `
            -Algorithm SHA256
    ).Hash


if ($stableHashBefore -ne $stableHashAfter) {

    throw "v3.2 Stable manifest changed."
}


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE SMART v0.1 SMOKE TEST"
Write-Host "========================================"
Write-Host ""

Write-Host "PROGRAM FILES      : PASS"
Write-Host "POWERSHELL SYNTAX  : PASS"
Write-Host "CONFIG             : PASS"
Write-Host "ROUTER             : PASS"
Write-Host "PPT SAFE ROUTE     : PASS"
Write-Host "CLAUDE ROUTE       : PASS"
Write-Host "EXECUTION          : OFF"
Write-Host "AI CALLS           : 0"
Write-Host "v3.2 STABLE        : UNCHANGED"
Write-Host ""
Write-Host "STATUS : JUNGLE_SMART_V01_BASELINE_PASS"
Write-Host "========================================"