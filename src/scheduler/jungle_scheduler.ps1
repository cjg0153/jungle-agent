param(
    [int]$PollSeconds = 10,
    [int]$TestSeconds = 0
)

$ErrorActionPreference = "Stop"

$Version     = "0.1"
$ConfigPath  = "C:\AI\Jungle\config\scheduler_jobs.json"
$InboxDir    = "C:\AI\Jungle\tasks\inbox"
$RuntimeRoot = "C:\AI\Jungle\runtime\scheduler"
$LogRoot     = "C:\AI\Jungle\logs\scheduler"

$StateFile   = Join-Path $RuntimeRoot "state.json"
$StatusFile  = Join-Path $RuntimeRoot "status.json"
$LogFile     = Join-Path $LogRoot "jungle_scheduler.log"

@(
    $InboxDir,
    $RuntimeRoot,
    $LogRoot
) | ForEach-Object {
    New-Item -ItemType Directory -Path $_ -Force | Out-Null
}

function Write-SchedulerLog {
    param([string]$Message)

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content $LogFile "[$ts] $Message" -Encoding UTF8
}

function Load-State {

    if (!(Test-Path $StateFile)) {
        return @{}
    }

    try {
        $obj = Get-Content $StateFile -Raw | ConvertFrom-Json
        $state = @{}

        foreach ($p in $obj.PSObject.Properties) {
            $state[$p.Name] = $p.Value
        }

        return $state
    }
    catch {
        return @{}
    }
}

function Save-State {
    param([hashtable]$State)

    $State |
        ConvertTo-Json |
        Set-Content $StateFile -Encoding UTF8
}

function Enqueue-Job {
    param($Job)

    $id = [guid]::NewGuid().ToString()

    $task = [ordered]@{
        id           = $id
        type         = $Job.type
        priority     = $Job.priority
        request      = $Job.request
        status       = "NEW"
        source       = "Jungle Scheduler"
        schedule_job = $Job.id
        created_at   = (Get-Date).ToString("o")
    }

    $target = Join-Path $InboxDir "$id.json"
    $tmp    = "$target.tmp"

    $task |
        ConvertTo-Json -Depth 10 |
        Set-Content $tmp -Encoding UTF8

    Move-Item $tmp $target -Force

    Write-SchedulerLog "ENQUEUE job=$($Job.id) task=$id"
}

$startedAt = Get-Date

Write-SchedulerLog "START version=$Version pid=$PID"

try {

    while ($true) {

        if (Test-Path $ConfigPath) {

            $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            $state  = Load-State
            $now    = Get-Date
            $today  = $now.ToString("yyyy-MM-dd")

            foreach ($job in @($config.jobs)) {

                if (!$job.enabled) {
                    continue
                }

                if ($job.cadence -eq "DAILY") {

                    $scheduled = [datetime]::ParseExact(
                        "$today $($job.time)",
                        "yyyy-MM-dd HH:mm",
                        $null
                    )

                    $lastRun = $null

                    if ($state.ContainsKey($job.id)) {
                        $lastRun = $state[$job.id]
                    }

                    if (
                        $now -ge $scheduled -and
                        $lastRun -ne $today
                    ) {
                        Enqueue-Job $job

                        $state[$job.id] = $today
                        Save-State $state
                    }
                }
            }
        }

        $status = [ordered]@{
            service      = "Jungle Scheduler"
            version      = $Version
            status       = "RUNNING"
            host         = $env:COMPUTERNAME
            pid          = $PID
            heartbeat_at = (Get-Date).ToString("o")
        }

        $status |
            ConvertTo-Json |
            Set-Content $StatusFile -Encoding UTF8

        if (
            $TestSeconds -gt 0 -and
            ((Get-Date) - $startedAt).TotalSeconds -ge $TestSeconds
        ) {
            break
        }

        Start-Sleep -Seconds $PollSeconds
    }

    Write-SchedulerLog "STOP test-complete pid=$PID"

    Write-Host "STATUS : JUNGLE_SCHEDULER_TEST_COMPLETE" -ForegroundColor Green
}
finally {

    $final = [ordered]@{
        service      = "Jungle Scheduler"
        version      = $Version
        status       = "STOPPED"
        host         = $env:COMPUTERNAME
        pid          = $PID
        heartbeat_at = (Get-Date).ToString("o")
    }

    $final |
        ConvertTo-Json |
        Set-Content $StatusFile -Encoding UTF8
}
