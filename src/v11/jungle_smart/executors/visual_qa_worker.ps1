# JUNGLE_SMART_VISUAL_QA_WORKER_V1

param(
    [Parameter(Mandatory=$true)]
    [string]$PromptPath,

    [Parameter(Mandatory=$true)]
    [string]$ResultPath
)

$ErrorActionPreference =
    "Stop"


try {

    if (-not (
        Test-Path `
            -LiteralPath $PromptPath `
            -PathType Leaf
    )) {

        throw "Visual QA prompt file missing: $PromptPath"
    }


    $claudePath =
        Join-Path `
            $env:USERPROFILE `
            ".local\bin\claude.exe"


    if (-not (
        Test-Path `
            -LiteralPath $claudePath `
            -PathType Leaf
    )) {

        throw "Claude CLI missing: $claudePath"
    }


    $prompt =
        [System.IO.File]::ReadAllText(
            $PromptPath
        )


    $workDirectory =
        Split-Path `
            -Parent `
            $PromptPath


    Push-Location $workDirectory


    try {

        $output =
            & $claudePath `
                -p $prompt 2>&1


        $exitCode =
            $LASTEXITCODE
    }
    finally {

        Pop-Location
    }


    $raw =
        (
            $output |
            Out-String
        ).Trim()


    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )


    [System.IO.File]::WriteAllText(
        $ResultPath,
        $raw,
        $encoding
    )


    if ($exitCode -ne 0) {

        throw (
            "Claude Visual QA exited with code " +
            $exitCode
        )
    }


    Write-Output "STATUS : JUNGLE_VISUAL_QA_WORKER_COMPLETE"

    exit 0
}
catch {

    Write-Error $_.Exception.Message

    exit 1
}