# JUNGLE_SMART_WATCHDOG_V1

function Test-JungleProcessAlive {

    param(
        [Parameter(Mandatory=$true)]
        [int]$ProcessId
    )

    $process =
        Get-Process `
            -Id $ProcessId `
            -ErrorAction SilentlyContinue

    return (
        $null -ne $process
    )
}


function Stop-JungleProcessTree {

    param(
        [Parameter(Mandatory=$true)]
        [int]$ProcessId
    )

    $taskkill =
        Get-Command `
            "taskkill.exe" `
            -ErrorAction SilentlyContinue


    if ($null -ne $taskkill) {

        try {

            & $taskkill.Source `
                /PID $ProcessId `
                /T `
                /F *> $null
        }
        catch {}
    }


    Start-Sleep `
        -Milliseconds 250


    if (
        Test-JungleProcessAlive `
            -ProcessId $ProcessId
    ) {

        try {

            Stop-Process `
                -Id $ProcessId `
                -Force `
                -ErrorAction SilentlyContinue
        }
        catch {}
    }


    Start-Sleep `
        -Milliseconds 250


    return (
        -not (
            Test-JungleProcessAlive `
                -ProcessId $ProcessId
        )
    )
}


function Invoke-JungleProcessWithWatchdog {

    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [Parameter(Mandatory=$true)]
        [int]$TimeoutSeconds,

        [Parameter(Mandatory=$true)]
        [string]$WorkDirectory,

        [int]$PollMilliseconds = 200
    )


    if ($TimeoutSeconds -lt 1) {

        throw "TimeoutSeconds must be at least 1."
    }


    if ($PollMilliseconds -lt 50) {

        throw "PollMilliseconds must be at least 50."
    }


    $resolvedCommand =
        Get-Command `
            $FilePath `
            -ErrorAction SilentlyContinue


    if ($null -eq $resolvedCommand) {

        if (-not (
            Test-Path `
                -LiteralPath $FilePath `
                -PathType Leaf
        )) {

            throw "Watchdog executable not found: $FilePath"
        }


        $resolvedFile =
            $FilePath
    }
    else {

        $resolvedFile =
            $resolvedCommand.Source
    }


    New-Item `
        -ItemType Directory `
        -Path $WorkDirectory `
        -Force |
    Out-Null


    $stdoutPath =
        Join-Path `
            $WorkDirectory `
            "stdout.txt"


    $stderrPath =
        Join-Path `
            $WorkDirectory `
            "stderr.txt"


    if (Test-Path -LiteralPath $stdoutPath) {

        Remove-Item `
            -LiteralPath $stdoutPath `
            -Force
    }


    if (Test-Path -LiteralPath $stderrPath) {

        Remove-Item `
            -LiteralPath $stderrPath `
            -Force
    }


    $stopwatch =
        [System.Diagnostics.Stopwatch]::StartNew()


    $process =
        $null


    try {

        $process =
            Start-Process `
                -FilePath $resolvedFile `
                -ArgumentList $Arguments `
                -PassThru `
                -WindowStyle Hidden `
                -RedirectStandardOutput $stdoutPath `
                -RedirectStandardError $stderrPath


        $processId =
            [int]$process.Id


        $timedOut =
            $false


        while ($true) {

            $process.Refresh()


            if ($process.HasExited) {

                break
            }


            if (
                $stopwatch.Elapsed.TotalSeconds -ge
                $TimeoutSeconds
            ) {

                $timedOut =
                    $true

                break
            }


            Start-Sleep `
                -Milliseconds $PollMilliseconds
        }


        $killed =
            $false


        if ($timedOut) {

            $killed =
                Stop-JungleProcessTree `
                    -ProcessId $processId


            try {

                $process.WaitForExit(
                    5000
                )
            }
            catch {}
        }
        else {

            try {

                $process.WaitForExit()
            }
            catch {}
        }


        $stopwatch.Stop()


        $exitCode =
            $null


        if (-not $timedOut) {

            try {

                $exitCode =
                    [int]$process.ExitCode
            }
            catch {}
        }


        return [pscustomobject]@{

            started =
                $true

            process_id =
                $processId

            completed =
                (-not $timedOut)

            timed_out =
                $timedOut

            killed =
                $killed

            exit_code =
                $exitCode

            duration_ms =
                [int64]$stopwatch.ElapsedMilliseconds

            stdout_path =
                $stdoutPath

            stderr_path =
                $stderrPath
        }
    }
    finally {

        if ($stopwatch.IsRunning) {

            $stopwatch.Stop()
        }


        if ($null -ne $process) {

            try {

                $process.Dispose()
            }
            catch {}
        }
    }
}


function Get-JungleWatchdogStatus {

    return [pscustomobject]@{

        enabled =
            $true

        version =
            "V1"

        status =
            "READY"

        enforcement =
            "PROCESS_LEVEL"

        ppt_bridge_enforcement =
            "DEFERRED_TO_WORKER_PROTOCOL"
    }
}