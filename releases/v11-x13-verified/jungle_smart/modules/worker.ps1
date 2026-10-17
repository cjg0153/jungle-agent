# JUNGLE_SMART_WORKER_PROTOCOL_V2
# JUNGLE_WORKER_UNICODE_TRANSPORT_V1
# JUNGLE_WORKER_EXPLICIT_RESULT_V1

function Get-JungleWorkerTextSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )


    $sha =
        [System.Security.Cryptography.SHA256]::Create()


    try {

        $bytes =
            [System.Text.Encoding]::UTF8.GetBytes(
                $Text
            )


        $hash =
            $sha.ComputeHash(
                $bytes
            )


        return (
            [System.BitConverter]::ToString(
                $hash
            )
        ).Replace(
            "-",
            ""
        )
    }
    finally {

        $sha.Dispose()
    }
}


function Write-JunglePptSafeWorkerInput {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Request,

        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )


    $requestHash =
        Get-JungleWorkerTextSha256 `
            -Text $Request


    $record =
        [ordered]@{

            protocol =
                "JUNGLE_SMART_PPT_SAFE_WORKER_V2"

            request =
                $Request

            request_sha256 =
                $requestHash

            config_path =
                $ConfigPath

            created_at =
                (Get-Date).ToString("s")
        }


    $json =
        $record |
        ConvertTo-Json -Depth 5


    # Windows PowerShell 5.1 protocol rule:
    # UTF-8 WITH BOM.
    $encoding =
        New-Object System.Text.UTF8Encoding(
            $true
        )


    [System.IO.File]::WriteAllText(
        $OutputPath,
        $json,
        $encoding
    )


    return $OutputPath
}


function Invoke-JunglePptSafeWorker {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Root,

        [Parameter(Mandatory=$true)]
        [string]$Request,

        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,

        [Parameter(Mandatory=$true)]
        [string]$RunDirectory,

        [Parameter(Mandatory=$true)]
        [int]$TimeoutSeconds,

        [int]$PollMilliseconds = 200
    )


    $workerScript =
        Join-Path `
            $Root `
            "executors\ppt_safe_worker.ps1"


    if (-not (
        Test-Path `
            -LiteralPath $workerScript `
            -PathType Leaf
    )) {

        throw "PPT SAFE worker script missing: $workerScript"
    }


    $powershellExe =
        Join-Path `
            $PSHOME `
            "powershell.exe"


    if (-not (
        Test-Path `
            -LiteralPath $powershellExe `
            -PathType Leaf
    )) {

        throw "Windows PowerShell missing: $powershellExe"
    }


    New-Item `
        -ItemType Directory `
        -Path $RunDirectory `
        -Force |
    Out-Null


    $inputPath =
        Join-Path `
            $RunDirectory `
            "worker_input.json"


    $resultPath =
        Join-Path `
            $RunDirectory `
            "worker_result.json"


    Write-JunglePptSafeWorkerInput `
        -Request $Request `
        -ConfigPath $ConfigPath `
        -OutputPath $inputPath |
    Out-Null


    $processDir =
        Join-Path `
            $RunDirectory `
            "process"


    $processResult =
        Invoke-JungleProcessWithWatchdog `
            -FilePath $powershellExe `
            -Arguments @(
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                $workerScript,
                "-InputPath",
                $inputPath,
                "-ResultPath",
                $resultPath
            ) `
            -TimeoutSeconds $TimeoutSeconds `
            -WorkDirectory $processDir `
            -PollMilliseconds $PollMilliseconds


    $stdout =
        ""


    if (
        Test-Path `
            -LiteralPath $processResult.stdout_path `
            -PathType Leaf
    ) {

        $stdout =
            [System.IO.File]::ReadAllText(
                $processResult.stdout_path
            )
    }


    $stderr =
        ""


    if (
        Test-Path `
            -LiteralPath $processResult.stderr_path `
            -PathType Leaf
    ) {

        $stderr =
            [System.IO.File]::ReadAllText(
                $processResult.stderr_path
            )
    }


    $workerResultRecord =
        $null


    if (
        Test-Path `
            -LiteralPath $resultPath `
            -PathType Leaf
    ) {

        $resultText =
            [System.IO.File]::ReadAllText(
                $resultPath
            )


        if (-not [string]::IsNullOrWhiteSpace(
            $resultText
        )) {

            try {

                $workerResultRecord =
                    $resultText |
                    ConvertFrom-Json
            }
            catch {

                throw (
                    "Worker result JSON parse failed. " +
                    $_.Exception.Message
                )
            }
        }
    }


    $pipelineComplete =
        $false


    $artifactPath =
        $null


    $protocol =
        $null


    if ($null -ne $workerResultRecord) {

        $protocol =
            [string]$workerResultRecord.protocol


        if (
            $protocol -ne
            "JUNGLE_SMART_PPT_SAFE_WORKER_RESULT_V2"
        ) {

            throw (
                "Worker result protocol mismatch: " +
                $protocol
            )
        }


        $pipelineComplete =
            [bool]$workerResultRecord.pipeline_complete


        if (-not [string]::IsNullOrWhiteSpace(
            [string]$workerResultRecord.artifact_path
        )) {

            $artifactPath =
                [string]$workerResultRecord.artifact_path
        }
    }


    return [pscustomobject]@{

        started =
            $processResult.started

        process_id =
            $processResult.process_id

        completed =
            $processResult.completed

        timed_out =
            $processResult.timed_out

        killed =
            $processResult.killed

        exit_code =
            $processResult.exit_code

        duration_ms =
            $processResult.duration_ms

        input_path =
            $inputPath

        result_path =
            $resultPath

        stdout_path =
            $processResult.stdout_path

        stderr_path =
            $processResult.stderr_path

        stdout =
            $stdout

        stderr =
            $stderr

        result_protocol =
            $protocol

        pipeline_complete =
            $pipelineComplete

        artifact_path =
            $artifactPath
    }
}


function Get-JungleWorkerStatus {

    return [pscustomobject]@{

        enabled =
            $true

        version =
            "V2"

        protocol =
            "JUNGLE_SMART_PPT_SAFE_WORKER_V2"

        result_protocol =
            "JUNGLE_SMART_PPT_SAFE_WORKER_RESULT_V2"

        unicode_transport =
            $true

        isolation =
            "PROCESS"

        watchdog_enforced =
            $true

        explicit_result =
            $true

        status =
            "READY"
    }
}

# ================================================================
# JUNGLE_WORKER_EXPLICIT_RESULT_GUARD_V1
#
# Process completion is NOT task success.
#
# worker_result.json is authoritative for Worker Protocol V2.
# A FAIL result must stop the Smart pipeline before artifact/G1/
# Changed Scope/Visual QA processing.
# ================================================================

function Assert-JungleWorkerExplicitResult {

    param(
        [Parameter(Mandatory=$true)]
        [string]$ResultPath
    )


    if (-not (
        Test-Path `
            -LiteralPath $ResultPath `
            -PathType Leaf
    )) {

        throw (
            "Worker explicit result file is missing: " +
            $ResultPath
        )
    }


    $text =
        [System.IO.File]::ReadAllText(
            $ResultPath
        )


    if ([string]::IsNullOrWhiteSpace(
        $text
    )) {

        throw "Worker explicit result file is empty."
    }


    try {

        $record =
            $text |
            ConvertFrom-Json
    }
    catch {

        throw (
            "Worker explicit result JSON parse failed. " +
            $_.Exception.Message
        )
    }


    if (
        [string]$record.protocol -ne
        "JUNGLE_SMART_PPT_SAFE_WORKER_RESULT_V2"
    ) {

        throw (
            "Worker explicit result protocol mismatch: " +
            [string]$record.protocol
        )
    }


    $status =
        (
            [string]$record.status
        ).ToUpperInvariant()


    $pipelineComplete =
        [bool]$record.pipeline_complete


    if (
        $status -ne "PASS" `
        -or
        -not $pipelineComplete
    ) {

        $message =
            [string]$record.message


        if ([string]::IsNullOrWhiteSpace(
            $message
        )) {

            $message =
                "No Worker failure message was provided."
        }


        throw (
            "PPT SAFE Worker explicit result FAILED. " +
            "status=" +
            $status +
            " pipeline_complete=" +
            $pipelineComplete +
            " message=" +
            $message
        )
    }


    $artifactPath =
        [string]$record.artifact_path


    if ([string]::IsNullOrWhiteSpace(
        $artifactPath
    )) {

        throw (
            "Worker explicit PASS result has no artifact path."
        )
    }


    if (-not (
        Test-Path `
            -LiteralPath $artifactPath `
            -PathType Leaf
    )) {

        throw (
            "Worker explicit PASS artifact does not exist: " +
            $artifactPath
        )
    }


    return $record
}

# ================================================================
# JUNGLE_WORKER_SAFE_NOOP_GUARD_V1
#
# Accepted Worker outcomes:
#
# 1. PASS
#    - pipeline_complete == true
#    - artifact exists
#
# 2. SAFE_NOOP
#    - pipeline_complete == false
#    - accepted == true
#    - hard_gate_state == NEEDS_REVIEW
#    - byte_identical == true
#    - source and artifact hashes independently recomputed
#    - source SHA256 == artifact SHA256
#
# Hard Gate thresholds are NOT changed.
# ================================================================

function Assert-JungleWorkerAcceptedResult {

    param(
        [Parameter(Mandatory=$true)]
        [string]$ResultPath
    )


    if (-not (
        Test-Path `
            -LiteralPath $ResultPath `
            -PathType Leaf
    )) {

        throw (
            "Worker result file missing: " +
            $ResultPath
        )
    }


    $text =
        [System.IO.File]::ReadAllText(
            $ResultPath
        )


    if ([string]::IsNullOrWhiteSpace(
        $text
    )) {

        throw "Worker result file is empty."
    }


    try {

        $record =
            $text |
            ConvertFrom-Json
    }
    catch {

        throw (
            "Worker result JSON parse failed. " +
            $_.Exception.Message
        )
    }


    if (
        [string]$record.protocol -ne
        "JUNGLE_SMART_PPT_SAFE_WORKER_RESULT_V2"
    ) {

        throw (
            "Worker result protocol mismatch: " +
            [string]$record.protocol
        )
    }


    $status =
        (
            [string]$record.status
        ).ToUpperInvariant()


    # ------------------------------------------------------------
    # VERIFIED PASS
    # ------------------------------------------------------------

    if ($status -eq "PASS") {

        if (-not [bool]$record.pipeline_complete) {

            throw (
                "Worker PASS requires pipeline_complete=true."
            )
        }


        $artifact =
            [string]$record.artifact_path


        if ([string]::IsNullOrWhiteSpace(
            $artifact
        )) {

            throw "Worker PASS has no artifact path."
        }


        if (-not (
            Test-Path `
                -LiteralPath $artifact `
                -PathType Leaf
        )) {

            throw (
                "Worker PASS artifact missing: " +
                $artifact
            )
        }


        return $record
    }


    # ------------------------------------------------------------
    # SAFE_NOOP
    # ------------------------------------------------------------

    if ($status -eq "SAFE_NOOP") {

        if ([bool]$record.pipeline_complete) {

            throw (
                "SAFE_NOOP must not claim verified pipeline completion."
            )
        }


        if (-not [bool]$record.accepted) {

            throw "SAFE_NOOP accepted flag must be true."
        }


        if (
            [string]$record.hard_gate_state -ne
            "NEEDS_REVIEW"
        ) {

            throw (
                "SAFE_NOOP requires hard_gate_state=NEEDS_REVIEW."
            )
        }


        if (-not [bool]$record.byte_identical) {

            throw (
                "SAFE_NOOP requires byte_identical=true."
            )
        }


        $sourcePath =
            [string]$record.source_path


        $artifactPath =
            [string]$record.artifact_path


        foreach ($file in @(
            $sourcePath,
            $artifactPath
        )) {

            if ([string]::IsNullOrWhiteSpace(
                $file
            )) {

                throw "SAFE_NOOP contains an empty file path."
            }


            if (-not (
                Test-Path `
                    -LiteralPath $file `
                    -PathType Leaf
            )) {

                throw (
                    "SAFE_NOOP file missing: " +
                    $file
                )
            }
        }


        $actualSourceHash =
            (
                Get-FileHash `
                    -LiteralPath $sourcePath `
                    -Algorithm SHA256
            ).Hash


        $actualArtifactHash =
            (
                Get-FileHash `
                    -LiteralPath $artifactPath `
                    -Algorithm SHA256
            ).Hash


        if (
            $actualSourceHash -ne
            $actualArtifactHash
        ) {

            throw (
                "SAFE_NOOP rejected: source and artifact bytes differ."
            )
        }


        if (
            [string]$record.source_sha256 -ne
            $actualSourceHash
        ) {

            throw (
                "SAFE_NOOP source SHA256 record mismatch."
            )
        }


        if (
            [string]$record.artifact_sha256 -ne
            $actualArtifactHash
        ) {

            throw (
                "SAFE_NOOP artifact SHA256 record mismatch."
            )
        }


        return $record
    }


    $message =
        [string]$record.message


    if ([string]::IsNullOrWhiteSpace(
        $message
    )) {

        $message =
            "No Worker failure message."
    }


    throw (
        "Worker result rejected. status=" +
        $status +
        " message=" +
        $message
    )
}