param(
    [int]$PollSeconds = 2,
    [int]$TestSeconds = 0
)

$ErrorActionPreference = "Stop"

$Version = "0.1"

$TaskRoot    = "C:\AI\Jungle\tasks"
$InboxDir    = Join-Path $TaskRoot "inbox"
$ReadyDir    = Join-Path $TaskRoot "ready"
$FailedDir   = Join-Path $TaskRoot "failed"

$RuntimeRoot = "C:\AI\Jungle\runtime\queue"
$LogRoot     = "C:\AI\Jungle\logs\queue"

$StatusFile  = Join-Path $RuntimeRoot "status.json"
$LogFile     = Join-Path $LogRoot "jungle_queue.log"

@(
    $InboxDir,
    $ReadyDir,
    $FailedDir,
    $RuntimeRoot,
    $LogRoot
) | ForEach-Object {
    New-Item -ItemType Directory -Path $_ -Force | Out-Null
}

function Write-QueueLog {
    param([string]$Message)

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content $LogFile "[$ts] $Message" -Encoding UTF8
}

function Write-QueueStatus {

    $status = [ordered]@{
        service       = "Jungle Queue"
        version       = $Version
        status        = "RUNNING"
        host          = $env:COMPUTERNAME
        pid           = $PID
        heartbeat_at  = (Get-Date).ToString("o")
        inbox_count   = @(Get-ChildItem $InboxDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
        ready_count   = @(Get-ChildItem $ReadyDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
        failed_count  = @(Get-ChildItem $FailedDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
    }

    $tmp = "$StatusFile.tmp"

    $status |
        ConvertTo-Json |
        Set-Content $tmp -Encoding UTF8

    Move-Item $tmp $StatusFile -Force
}

$mutexName = "Global\JungleQueue_v01"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

try {
    $acquired = $mutex.WaitOne(0, $false)
}
catch [System.Threading.AbandonedMutexException] {
    $acquired = $true
}

if (!$acquired) {
    Write-Host "Jungle Queue is already running." -ForegroundColor Yellow
    exit 10
}

$startedAt = Get-Date

try {

    Write-QueueLog "START version=$Version pid=$PID"

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " JUNGLE QUEUE v$Version" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan

    if ($TestSeconds -gt 0) {
        Write-Host "MODE : TEST ($TestSeconds sec)" -ForegroundColor Yellow
    }
    else {
        Write-Host "MODE : ALWAYS-ON" -ForegroundColor Green
    }

    while ($true) {

        $items = Get-ChildItem $InboxDir -Filter *.json -File |
            Sort-Object CreationTime

        foreach ($item in $items) {

            try {

                $task = Get-Content $item.FullName -Raw |
                    ConvertFrom-Json

                if (
                    [string]::IsNullOrWhiteSpace($task.id) -or
                    [string]::IsNullOrWhiteSpace($task.type) -or
                    [string]::IsNullOrWhiteSpace($task.request)
                ) {
                    throw "Required field missing: id/type/request"
                }

                $task.status = "READY"

                if ($task.PSObject.Properties.Name -contains "accepted_at") {
                    $task.accepted_at = (Get-Date).ToString("o")
                }
                else {
                    $task |
                        Add-Member `
                            -NotePropertyName accepted_at `
                            -NotePropertyValue (Get-Date).ToString("o")
                }

                $readyPath = Join-Path $ReadyDir $item.Name
                $tmpPath   = "$readyPath.tmp"

                $task |
                    ConvertTo-Json -Depth 10 |
                    Set-Content $tmpPath -Encoding UTF8

                Move-Item $tmpPath $readyPath -Force
                Remove-Item $item.FullName -Force

                Write-QueueLog "ACCEPT id=$($task.id) type=$($task.type)"
            }
            catch {

                $failedPath = Join-Path $FailedDir $item.Name

                Move-Item $item.FullName $failedPath -Force

                Write-QueueLog "FAIL file=$($item.Name) error=$($_.Exception.Message)"
            }
        }

        Write-QueueStatus

        if (
            $TestSeconds -gt 0 -and
            ((Get-Date) - $startedAt).TotalSeconds -ge $TestSeconds
        ) {
            break
        }

        Start-Sleep -Seconds $PollSeconds
    }

    Write-QueueLog "STOP test-complete pid=$PID"

    Write-Host ""
    Write-Host "STATUS : JUNGLE_QUEUE_TEST_COMPLETE" -ForegroundColor Green
}
finally {

    try {
        $final = [ordered]@{
            service      = "Jungle Queue"
            version      = $Version
            status       = "STOPPED"
            host         = $env:COMPUTERNAME
            pid          = $PID
            heartbeat_at = (Get-Date).ToString("o")
            inbox_count  = @(Get-ChildItem $InboxDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
            ready_count  = @(Get-ChildItem $ReadyDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
            failed_count = @(Get-ChildItem $FailedDir -Filter *.json -File -ErrorAction SilentlyContinue).Count
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
