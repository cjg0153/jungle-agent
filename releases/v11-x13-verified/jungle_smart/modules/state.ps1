function New-JungleRunContext {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Root,

        [Parameter(Mandatory=$true)]
        [string]$Request
    )

    # JUNGLE_SMART_RUN_ID_UNIQUENESS_V1
    #
    # Millisecond timestamp + random suffix prevents same-second
    # and same-millisecond run-directory collisions.
    $stamp =
        Get-Date -Format "yyyyMMdd_HHmmss_fff"

    $suffix =
        [guid]::NewGuid().ToString("N").Substring(0, 8)

    $runId =
        "run_" + $stamp + "_" + $suffix

    $runDir =
        Join-Path `
            (Join-Path $Root "runs") `
            $runId

    New-Item `
        -ItemType Directory `
        -Path $runDir `
        -Force |
    Out-Null

    return [pscustomobject]@{

        run_id =
            $runId

        run_dir =
            $runDir

        request =
            $Request

        created_at =
            (Get-Date).ToString("s")
    }
}