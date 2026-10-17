param(
    [int]$HeartbeatSeconds = 30,
    [int]$TestSeconds = 0
)

$ErrorActionPreference = "Stop"

$Version = "0.2"
$RuntimeRoot = "C:\AI\Jungle\runtime\server"
$LogRoot     = "C:\AI\Jungle\logs\server"

$StatusFile = Join-Path $RuntimeRoot "status.json"
$LogFile    = Join-Path $LogRoot "jungle_server.log"

New-Item -ItemType Directory -Path $RuntimeRoot -Force | Out-Null
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

function Write-JungleLog {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$timestamp] $Message" -Encoding UTF8
}

# Prevent multiple Jungle Server instances.
$mutexName = "Global\JungleServer_v01"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

try {
    $acquired = $mutex.WaitOne(0, $false)
}
catch [System.Threading.AbandonedMutexException] {
    $acquired = $true
}

if (!$acquired) {
    Write-Host "Jungle Server is already running." -ForegroundColor Yellow
    exit 10
}

$startedAt = Get-Date
$pidValue  = $PID

try {

    Write-JungleLog "START version=$Version pid=$pidValue host=$env:COMPUTERNAME"

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " JUNGLE SERVER v$Version" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "HOST      : $env:COMPUTERNAME"
    Write-Host "PID       : $pidValue"
    Write-Host "HEARTBEAT : $HeartbeatSeconds sec"

    if ($TestSeconds -gt 0) {
        Write-Host "TEST MODE : $TestSeconds sec" -ForegroundColor Yellow
    }
    else {
        Write-Host "MODE      : ALWAYS-ON" -ForegroundColor Green
    }

    while ($true) {

        $now = Get-Date

        $status = [ordered]@{
            server       = "Jungle Server"
            version      = $Version
            status       = "RUNNING"
            host         = $env:COMPUTERNAME
            pid          = $pidValue
            started_at   = $startedAt.ToString("o")
            heartbeat_at = $now.ToString("o")
            uptime_sec   = [int](($now - $startedAt).TotalSeconds)
        }

        $tmp = "$StatusFile.tmp"

        $status |
            ConvertTo-Json |
            Set-Content $tmp -Encoding UTF8

        Move-Item $tmp $StatusFile -Force

        if (
            $TestSeconds -gt 0 -and
            ($now - $startedAt).TotalSeconds -ge $TestSeconds
        ) {
            break
        }

        Start-Sleep -Seconds $HeartbeatSeconds
    }

    Write-JungleLog "STOP test-complete pid=$pidValue"

    Write-Host ""
    Write-Host "STATUS : JUNGLE_SERVER_TEST_COMPLETE" -ForegroundColor Green
}
catch {

    Write-JungleLog "ERROR pid=$pidValue message=$($_.Exception.Message)"
    throw
}
finally {

    $stoppedAt = Get-Date

    try {
        $finalStatus = [ordered]@{
            server       = "Jungle Server"
            version      = $Version
            status       = "STOPPED"
            host         = $env:COMPUTERNAME
            pid          = $pidValue
            started_at   = $startedAt.ToString("o")
            heartbeat_at = $stoppedAt.ToString("o")
            uptime_sec   = [int](($stoppedAt - $startedAt).TotalSeconds)
        }

        $tmp = "$StatusFile.tmp"

        $finalStatus |
            ConvertTo-Json |
            Set-Content $tmp -Encoding UTF8

        Move-Item $tmp $StatusFile -Force
    }
    catch {}

    if ($mutex) {
        try { $mutex.ReleaseMutex() } catch {}
        $mutex.Dispose()
    }
}

