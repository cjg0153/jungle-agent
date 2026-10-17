param(
    [int]$HeartbeatSeconds = 30,
    [int]$WatchdogTimeoutSeconds = 90,
    [int]$TestSeconds = 0
)

$ErrorActionPreference = "Stop"

$Version = "0.3"

$RuntimeRoot = "C:\AI\Jungle\runtime\server"
$LogRoot     = "C:\AI\Jungle\logs\server"

$StatusFile = Join-Path $RuntimeRoot "status.json"
$LogFile    = Join-Path $LogRoot "jungle_server.log"

$WatchdogScript = "C:\AI\Jungle\runtime\watchdog\bin\jungle_watchdog.ps1"
$WatchdogStatus = "C:\AI\Jungle\runtime\watchdog\status.json"

$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

New-Item -ItemType Directory -Path $RuntimeRoot -Force | Out-Null
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

function Write-ServerLog {
    param([string]$Message)

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content $LogFile "[$ts] $Message" -Encoding UTF8
}

function Get-WatchdogHealth {

    if (!(Test-Path $WatchdogStatus -PathType Leaf)) {
        return @{
            Healthy = $false
            Reason  = "STATUS_MISSING"
            PID     = $null
        }
    }

    try {
        $status = Get-Content $WatchdogStatus -Raw | ConvertFrom-Json
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

    if ($age -gt $WatchdogTimeoutSeconds) {
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

function Test-JungleProcessIdentity {
    param(
        [int]$ProcessId,
        [string]$ExpectedScript
    )

    try {
        $p = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction Stop

        if ($null -eq $p) {
            return $false
        }

        $commandLine = [string]$p.CommandLine

        if ([string]::IsNullOrWhiteSpace($commandLine)) {
            return $false
        }

        return $commandLine.IndexOf(
            $ExpectedScript,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -ge 0
    }
    catch {
        return $false
    }
}

function Start-JungleWatchdog {

    if (!(Test-Path $WatchdogScript -PathType Leaf)) {
        Write-ServerLog "WATCHDOG_START_FAIL reason=SCRIPT_MISSING"
        return $false
    }

    $args = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$WatchdogScript`" -PollSeconds 10 -HeartbeatTimeoutSeconds 45"

    $proc = Start-Process `
        -FilePath $PowerShellExe `
        -ArgumentList $args `
        -WorkingDirectory "C:\AI\Jungle" `
        -WindowStyle Hidden `
        -PassThru

    Write-ServerLog "WATCHDOG_START pid=$($proc.Id)"

    Start-Sleep -Seconds 3

    return $true
}

$mutexName = "Global\JungleServer_v03"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

try {
    $acquired = $mutex.WaitOne(0, $false)
}
catch [System.Threading.AbandonedMutexException] {
    $acquired = $true
}

if (!$acquired) {
    Write-Host "Jungle Server v0.3 is already running." -ForegroundColor Yellow
    exit 10
}

$startedAt = Get-Date

try {

    Write-ServerLog "START version=$Version pid=$PID"

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " JUNGLE SERVER v$Version" -ForegroundColor Green
    Write-Host " SUPERVISOR MODE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan

    if ($TestSeconds -gt 0) {
        Write-Host "MODE : TEST ($TestSeconds sec)" -ForegroundColor Yellow
    }
    else {
        Write-Host "MODE : ALWAYS-ON" -ForegroundColor Green
    }

    while ($true) {

        $health = Get-WatchdogHealth

        if (!$health.Healthy) {

            Write-ServerLog "WATCHDOG_UNHEALTHY reason=$($health.Reason) pid=$($health.PID)"

            if (
                $null -ne $health.PID -and
                $health.Reason -eq "HEARTBEAT_STALE"
            ) {
                $identityMatch = Test-JungleProcessIdentity `
                    -ProcessId $health.PID `
                    -ExpectedScript $WatchdogScript

                if ($identityMatch) {
                    Stop-Process -Id $health.PID -Force -ErrorAction SilentlyContinue
                    Write-ServerLog "WATCHDOG_KILL pid=$($health.PID) identity=VERIFIED"
                }
                else {
                    Write-ServerLog "WATCHDOG_KILL_SKIPPED pid=$($health.PID) reason=IDENTITY_MISMATCH"
                }
            }

            [void](Start-JungleWatchdog)

            $health = Get-WatchdogHealth
        }

        $status = [ordered]@{
            service      = "Jungle Server"
            version      = $Version
            role         = "Supervisor"
            status       = "RUNNING"
            host         = $env:COMPUTERNAME
            pid          = $PID
            heartbeat_at = (Get-Date).ToString("o")
            watchdog     = [ordered]@{
                healthy = $health.Healthy
                reason  = $health.Reason
                pid     = $health.PID
            }
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

        Start-Sleep -Seconds $HeartbeatSeconds
    }

    Write-ServerLog "STOP test-complete pid=$PID"

    Write-Host ""
    Write-Host "STATUS : JUNGLE_SERVER_V03_TEST_COMPLETE" -ForegroundColor Green
}
finally {

    try {
        $final = [ordered]@{
            service      = "Jungle Server"
            version      = $Version
            role         = "Supervisor"
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

