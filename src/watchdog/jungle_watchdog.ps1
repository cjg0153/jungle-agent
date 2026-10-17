param(
    [int]$PollSeconds = 10,
    [int]$HeartbeatTimeoutSeconds = 45,
    [int]$TestSeconds = 0
)

$ErrorActionPreference = "Stop"

$Version = "0.1"

$RuntimeRoot = "C:\AI\Jungle\runtime\watchdog"
$LogRoot     = "C:\AI\Jungle\logs\watchdog"

$StatusFile = Join-Path $RuntimeRoot "status.json"
$LogFile    = Join-Path $LogRoot "jungle_watchdog.log"

$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

New-Item -ItemType Directory -Path $RuntimeRoot -Force | Out-Null
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

$Services = @(
    @{
        Name       = "Queue"
        Script     = "C:\AI\Jungle\runtime\queue\bin\jungle_queue.ps1"
        StatusFile = "C:\AI\Jungle\runtime\queue\status.json"
        Arguments  = "-PollSeconds 2"
    },
    @{
        Name       = "Scheduler"
        Script     = "C:\AI\Jungle\runtime\scheduler\bin\jungle_scheduler.ps1"
        StatusFile = "C:\AI\Jungle\runtime\scheduler\status.json"
        Arguments  = "-PollSeconds 10"
    }
)

function Write-WatchdogLog {
    param([string]$Message)

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content $LogFile "[$ts] $Message" -Encoding UTF8
}

function Get-ServiceHealth {
    param($Service)

    if (!(Test-Path $Service.StatusFile -PathType Leaf)) {
        return @{
            Healthy = $false
            Reason  = "STATUS_MISSING"
            PID     = $null
        }
    }

    try {
        $status = Get-Content $Service.StatusFile -Raw | ConvertFrom-Json
    }
    catch {
        return @{
            Healthy = $false
            Reason  = "STATUS_INVALID"
            PID     = $null
        }
    }

    $pidValue = $status.pid

    if ($status.status -ne "RUNNING") {
        return @{
            Healthy = $false
            Reason  = "STATUS_$($status.status)"
            PID     = $pidValue
        }
    }

    $proc = Get-Process -Id $pidValue -ErrorAction SilentlyContinue

    if ($null -eq $proc) {
        return @{
            Healthy = $false
            Reason  = "PROCESS_MISSING"
            PID     = $pidValue
        }
    }

    try {
        $heartbeat = [datetimeoffset]::Parse($status.heartbeat_at)
        $age = ((Get-Date) - $heartbeat.LocalDateTime).TotalSeconds
    }
    catch {
        return @{
            Healthy = $false
            Reason  = "HEARTBEAT_INVALID"
            PID     = $pidValue
        }
    }

    if ($age -gt $HeartbeatTimeoutSeconds) {
        return @{
            Healthy = $false
            Reason  = "HEARTBEAT_STALE"
            PID     = $pidValue
        }
    }

    return @{
        Healthy = $true
        Reason  = "HEALTHY"
        PID     = $pidValue
    }
}

function Start-JungleService {
    param($Service)

    if (!(Test-Path $Service.Script -PathType Leaf)) {
        Write-WatchdogLog "START_FAIL service=$($Service.Name) reason=SCRIPT_MISSING"
        return $false
    }

    $args = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$($Service.Script)`" $($Service.Arguments)"

    $proc = Start-Process `
        -FilePath $PowerShellExe `
        -ArgumentList $args `
        -WorkingDirectory "C:\AI\Jungle" `
        -WindowStyle Hidden `
        -PassThru

    Write-WatchdogLog "START service=$($Service.Name) pid=$($proc.Id)"

    Start-Sleep -Seconds 2

    return $true
}

$mutexName = "Global\JungleWatchdog_v01"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

try {
    $acquired = $mutex.WaitOne(0, $false)
}
catch [System.Threading.AbandonedMutexException] {
    $acquired = $true
}

if (!$acquired) {
    Write-Host "Jungle Watchdog is already running." -ForegroundColor Yellow
    exit 10
}

$startedAt = Get-Date

try {

    Write-WatchdogLog "START version=$Version pid=$PID"

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " JUNGLE WATCHDOG v$Version" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan

    if ($TestSeconds -gt 0) {
        Write-Host "MODE : TEST ($TestSeconds sec)" -ForegroundColor Yellow
    }
    else {
        Write-Host "MODE : ALWAYS-ON" -ForegroundColor Green
    }

    while ($true) {

        $healthSummary = @()

        foreach ($service in $Services) {

            $health = Get-ServiceHealth $service

            if (!$health.Healthy) {

                Write-WatchdogLog "UNHEALTHY service=$($service.Name) reason=$($health.Reason) pid=$($health.PID)"

                if (
                    $null -ne $health.PID -and
                    $health.Reason -eq "HEARTBEAT_STALE"
                ) {
                    Stop-Process -Id $health.PID -Force -ErrorAction SilentlyContinue
                    Write-WatchdogLog "KILL service=$($service.Name) pid=$($health.PID)"
                }

                [void](Start-JungleService $service)

                $health = Get-ServiceHealth $service
            }

            $healthSummary += [ordered]@{
                name    = $service.Name
                healthy = $health.Healthy
                reason  = $health.Reason
                pid     = $health.PID
            }
        }

        $status = [ordered]@{
            service      = "Jungle Watchdog"
            version      = $Version
            status       = "RUNNING"
            host         = $env:COMPUTERNAME
            pid          = $PID
            heartbeat_at = (Get-Date).ToString("o")
            services     = $healthSummary
        }

        $tmp = "$StatusFile.tmp"

        $status |
            ConvertTo-Json -Depth 10 |
            Set-Content $tmp -Encoding UTF8

        Move-Item $tmp $StatusFile -Force

        if (
            $TestSeconds -gt 0 -and
            ((Get-Date) - $startedAt).TotalSeconds -ge $TestSeconds
        ) {
            break
        }

        Start-Sleep -Seconds $PollSeconds
    }

    Write-WatchdogLog "STOP test-complete pid=$PID"

    Write-Host ""
    Write-Host "STATUS : JUNGLE_WATCHDOG_TEST_COMPLETE" -ForegroundColor Green
}
finally {

    try {
        $final = [ordered]@{
            service      = "Jungle Watchdog"
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
    catch {}

    if ($mutex) {
        try { $mutex.ReleaseMutex() } catch {}
        $mutex.Dispose()
    }
}
