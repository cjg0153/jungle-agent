param(
    [Parameter(Mandatory=$true)]
    [string]$Request,

    [string]$Type = "TEXT",

    [ValidateSet("LOW","NORMAL","HIGH")]
    [string]$Priority = "NORMAL"
)

$ErrorActionPreference = "Stop"

$InboxDir = "C:\AI\Jungle\tasks\inbox"

New-Item -ItemType Directory -Path $InboxDir -Force | Out-Null

$id = [guid]::NewGuid().ToString()

$task = [ordered]@{
    id         = $id
    type       = $Type
    priority   = $Priority
    request    = $Request
    status     = "NEW"
    source     = $env:COMPUTERNAME
    created_at = (Get-Date).ToString("o")
}

$target = Join-Path $InboxDir "$id.json"
$tmp    = "$target.tmp"

$task |
    ConvertTo-Json |
    Set-Content $tmp -Encoding UTF8

Move-Item $tmp $target -Force

Write-Host "TASK ENQUEUED" -ForegroundColor Green
Write-Host "ID      : $id"
Write-Host "TYPE    : $Type"
Write-Host "PRIORITY: $Priority"
Write-Host "FILE    : $target"
