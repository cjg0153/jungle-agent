param(
    [Parameter(Mandatory=$true)]
    [string]$Request,

    [Parameter(Mandatory=$true)]
    [string]$ConfigPath,

    [switch]$DryRun,

    [int]$TimeoutSeconds = 1800
)

$ErrorActionPreference =
    "Stop"


# ================================================================
# JUNGLE_SMART_PPT_REFLOW_BRIDGE_V3
#
# JUNGLE_SMART_REFLOW_WORKER_PROTOCOL_V2
# JUNGLE_SMART_REFLOW_HARD_GATE_WIRING_V1
#
# Bridge -> isolated REFLOW Worker V3
# ================================================================


function Get-JungleBridgeStringSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )


    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )


    $bytes =
        $encoding.GetBytes(
            $Text
        )


    $sha =
        [System.Security.Cryptography.SHA256]::Create()


    try {

        $hashBytes =
            $sha.ComputeHash(
                $bytes
            )
    }
    finally {

        $sha.Dispose()
    }


    return (
        -join (
            $hashBytes |
            ForEach-Object {

                $_.ToString("x2")
            }
        )
    ).ToUpperInvariant()
}


$root =
    Split-Path `
        -Parent (
            Split-Path `
                -Parent $MyInvocation.MyCommand.Path
        )


$config =
    Get-Content `
        -LiteralPath $ConfigPath `
        -Raw |
    ConvertFrom-Json


if (-not $config.feature_flags.reflow_v1) {

    throw (
        "REFLOW V1 execution is disabled by feature flag."
    )
}


$workerPath =
    Join-Path `
        $root `
        "executors\ppt_reflow_worker.ps1"


$executorPath =
    Join-Path `
        $root `
        "executors\run_pptx_reflow_v1.ps1"


$guardPath =
    Join-Path `
        $root `
        "modules\reflow_guard.ps1"


foreach ($file in @(
    $workerPath,
    $executorPath,
    $guardPath
)) {

    if (-not (
        Test-Path `
            -LiteralPath $file `
            -PathType Leaf
    )) {

        throw (
            "REFLOW dependency missing: " +
            $file
        )
    }
}


$stamp =
    Get-Date -Format "yyyyMMdd_HHmmss_fff"


$runId =
    (
        "reflow_worker_" +
        $stamp +
        "_" +
        [guid]::NewGuid().ToString("N").Substring(0,8)
    )


$runDirectory =
    Join-Path `
        (
            Join-Path `
                $root `
                "runs"
        ) `
        $runId


New-Item `
    -ItemType Directory `
    -Path $runDirectory `
    -Force |
Out-Null


$inputPath =
    Join-Path `
        $runDirectory `
        "worker_input.json"


$resultPath =
    Join-Path `
        $runDirectory `
        "worker_result.json"


$stdoutPath =
    Join-Path `
        $runDirectory `
        "worker_stdout.log"


$stderrPath =
    Join-Path `
        $runDirectory `
        "worker_stderr.log"


$requestHash =
    Get-JungleBridgeStringSha256 `
        -Text $Request


$inputObject =
    [ordered]@{

        protocol =
            "JUNGLE_SMART_REFLOW_WORKER_PROTOCOL_V2"

        request =
            $Request

        request_sha256 =
            $requestHash

        config_path =
            $ConfigPath

        executor_path =
            $executorPath

        guard_path =
            $guardPath

        result_path =
            $resultPath

        dry_run =
            [bool]$DryRun
    }


$inputJson =
    $inputObject |
    ConvertTo-Json -Depth 10


# Worker transport explicitly UTF-8 BOM.

$utf8Bom =
    New-Object System.Text.UTF8Encoding(
        $true
    )


[System.IO.File]::WriteAllText(
    $inputPath,
    $inputJson,
    $utf8Bom
)


$powerShellExe =
    (
        Get-Command `
            "powershell.exe" `
            -ErrorAction Stop
    ).Source


$process =
    Start-Process `
        -FilePath $powerShellExe `
        -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            $workerPath,
            "-InputJsonPath",
            $inputPath
        ) `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -PassThru


$deadline =
    (Get-Date).AddSeconds(
        $TimeoutSeconds
    )


while (-not $process.HasExited) {

    if ((Get-Date) -gt $deadline) {

        # --------------------------------------------------------
        # Try to kill worker process tree first.
        # --------------------------------------------------------

        try {

            $taskKill =
                Get-Command `
                    "taskkill.exe" `
                    -ErrorAction SilentlyContinue


            if ($null -ne $taskKill) {

                Start-Process `
                    -FilePath $taskKill.Source `
                    -ArgumentList @(
                        "/PID",
                        [string]$process.Id,
                        "/T",
                        "/F"
                    ) `
                    -WindowStyle Hidden `
                    -Wait
            }
        }
        catch {}


        try {

            if (-not $process.HasExited) {

                $process.Kill()
            }
        }
        catch {}


        throw (
            "REFLOW worker timeout after " +
            $TimeoutSeconds +
            " seconds."
        )
    }


    Start-Sleep `
        -Milliseconds 200


    $process.Refresh()
}


$exitCode =
    [int]$process.ExitCode


if (-not (
    Test-Path `
        -LiteralPath $resultPath `
        -PathType Leaf
)) {

    $stderrText =
        ""


    if (
        Test-Path `
            -LiteralPath $stderrPath `
            -PathType Leaf
    ) {

        $stderrText =
            [System.IO.File]::ReadAllText(
                $stderrPath
            )
    }


    throw (
        "REFLOW worker did not create result JSON. ExitCode=" +
        $exitCode +
        " STDERR=" +
        $stderrText
    )
}


$result =
    Get-Content `
        -LiteralPath $resultPath `
        -Raw |
    ConvertFrom-Json


if (
    [string]$result.protocol -ne
    "JUNGLE_SMART_REFLOW_WORKER_RESULT_V2"
) {

    throw (
        "Invalid REFLOW Worker Result Protocol."
    )
}


# ================================================================
# DRY RUN
# ================================================================

if ($DryRun) {

    if (
        [string]$result.status -ne
        "REFLOW_DRY_RUN_READY"
    ) {

        throw (
            "REFLOW dry-run failed. Status=" +
            [string]$result.status
        )
    }


    if (-not [bool]$result.accepted) {

        throw (
            "REFLOW dry-run not accepted."
        )
    }


    if ([bool]$result.executor_called) {

        throw (
            "REFLOW executor unexpectedly ran in DryRun."
        )
    }


    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE REFLOW V3 DRY RUN"
    Write-Host "========================================"
    Write-Host ""

    Write-Host "Worker status       : $($result.status)"
    Write-Host "Executor called     : $($result.executor_called)"
    Write-Host "Mutation Boundary   : $($result.mutation_boundary)"
    Write-Host "Changed Scope       : $($result.changed_scope)"
    Write-Host "Blind Pairwise      : $($result.blind_pairwise)"
    Write-Host "Visual Re-QA        : $($result.visual_reqa)"
    Write-Host "Hard Gate           : $($result.hard_gate)"
    Write-Host "AI calls            : $($result.ai_calls)"
    Write-Host "COM calls           : $($result.com_calls)"
    Write-Host "Run directory       : $runDirectory"
    Write-Host ""
    Write-Host "STATUS : JUNGLE_REFLOW_V3_DRY_RUN_PASS"
    Write-Host ""

    return
}


# ================================================================
# LIVE RESULT
# ================================================================

if (
    [string]$result.status -eq
    "REFLOW_VERIFIED"
) {

    if (
        -not [bool]$result.accepted `
        -or
        -not [bool]$result.pipeline_complete `
        -or
        -not [bool]$result.hard_gate_passed
    ) {

        throw (
            "REFLOW VERIFIED result has inconsistent acceptance fields."
        )
    }


    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE REFLOW VERIFIED"
    Write-Host "========================================"
    Write-Host ""

    Write-Host "Artifact          : $($result.artifact_path)"
    Write-Host "Changed slides    : $($result.changed_count)"
    Write-Host "Score             : $($result.score)"
    Write-Host "Critical          : $($result.critical)"
    Write-Host "Major             : $($result.major)"
    Write-Host "Minor             : $($result.minor)"
    Write-Host "Hard Gate         : PASS"
    Write-Host "Claude calls      : $($result.claude_calls)"
    Write-Host ""
    Write-Host "STATUS : JUNGLE_REFLOW_VERIFIED"
    Write-Host ""

    return
}


if (
    [string]$result.status -eq
    "REFLOW_NOOP"
) {

    if (-not [bool]$result.accepted) {

        throw (
            "REFLOW NOOP result not accepted."
        )
    }


    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE REFLOW NO-OP"
    Write-Host "========================================"
    Write-Host ""

    Write-Host "Artifact          : $($result.artifact_path)"
    Write-Host "Byte identical    : $($result.byte_identical)"
    Write-Host "Changed slides    : 0"
    Write-Host "Hard Gate         : SKIPPED_NO_VISUAL_CHANGE"
    Write-Host ""
    Write-Host "STATUS : JUNGLE_REFLOW_NOOP"
    Write-Host ""

    return
}


throw (
    "REFLOW pipeline rejected result. Status=" +
    [string]$result.status +
    " ExitCode=" +
    $exitCode
)