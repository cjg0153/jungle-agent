param(
    [Parameter(Mandatory=$true)]
    [string]$Request,

    [Parameter(Mandatory=$true)]
    [string]$ConfigPath
)

$ErrorActionPreference =
    "Stop"


$config =
    Get-Content `
        -LiteralPath $ConfigPath `
        -Raw |
    ConvertFrom-Json


if (-not $config.feature_flags.execution_enabled) {

    throw "Jungle Smart execution is disabled."
}


$stable =
    [string]$config.stable_candidate


$runner =
    [string]$config.ppt_safe_runner


$pairs =
    @(
        @{
            Working =
                "C:\AI\jungle_v32.ps1"

            Frozen =
                Join-Path `
                    $stable `
                    "jungle_v32.ps1"
        },

        @{
            Working =
                "C:\AI\run_verified_text_v3_2_hybrid.ps1"

            Frozen =
                Join-Path `
                    $stable `
                    "run_verified_text_v3_2_hybrid.ps1"
        },

        @{
            Working =
                "C:\AI\scripts\run_pptx_visual_modify_v1.ps1"

            Frozen =
                Join-Path `
                    $stable `
                    "scripts\run_pptx_visual_modify_v1.ps1"
        },

        @{
            Working =
                "C:\AI\scripts\run_pptx_visual_revision_v1.ps1"

            Frozen =
                Join-Path `
                    $stable `
                    "scripts\run_pptx_visual_revision_v1.ps1"
        },

        @{
            Working =
                "C:\AI\scripts\run_pptx_visual_reqa_v1.ps1"

            Frozen =
                Join-Path `
                    $stable `
                    "scripts\run_pptx_visual_reqa_v1.ps1"
        }
    )


foreach ($pair in $pairs) {

    if (-not (
        Test-Path `
            -LiteralPath $pair.Working `
            -PathType Leaf
    )) {

        throw "Working file missing: $($pair.Working)"
    }


    if (-not (
        Test-Path `
            -LiteralPath $pair.Frozen `
            -PathType Leaf
    )) {

        throw "Frozen file missing: $($pair.Frozen)"
    }


    $workingHash =
        (
            Get-FileHash `
                -LiteralPath $pair.Working `
                -Algorithm SHA256
        ).Hash


    $frozenHash =
        (
            Get-FileHash `
                -LiteralPath $pair.Frozen `
                -Algorithm SHA256
        ).Hash


    if ($workingHash -ne $frozenHash) {

        throw (
            "Stable integrity mismatch: " +
            $pair.Working
        )
    }
}


if (-not (
    Test-Path `
        -LiteralPath $runner `
        -PathType Leaf
)) {

    throw "PPT SAFE runner missing: $runner"
}


& $runner `
    -Request $Request