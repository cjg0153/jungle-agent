param(
    [switch]$Apply
)

$ErrorActionPreference = "Stop"

$SourceRoot = "C:\AI\Jungle\src\v11"
$TargetRoot = "C:\AI"
$BackupRoot = "C:\AI\Jungle\backups\deploy"
$LogRoot    = "C:\AI\Jungle\logs"

if (!(Test-Path $SourceRoot)) {
    throw "SOURCE NOT FOUND: $SourceRoot"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " JUNGLE SAFE DEPLOY v1" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

#
# 1. PowerShell syntax preflight
#
$syntaxFailures = @()

Get-ChildItem $SourceRoot -Filter *.ps1 -File -Recurse |
ForEach-Object {

    $tokens = $null
    $errors = $null

    [System.Management.Automation.Language.Parser]::ParseFile(
        $_.FullName,
        [ref]$tokens,
        [ref]$errors
    ) | Out-Null

    if (@($errors).Count -gt 0) {
        $syntaxFailures += $_.FullName
    }
}

if ($syntaxFailures.Count -gt 0) {
    Write-Host "SYNTAX PREFLIGHT : FAIL" -ForegroundColor Red
    $syntaxFailures
    throw "DEPLOY BLOCKED: PowerShell syntax failure."
}

Write-Host "SYNTAX PREFLIGHT : PASS" -ForegroundColor Green


#
# 2. Build deployment plan
#
$sourceFiles = Get-ChildItem $SourceRoot -File -Recurse

$plan = foreach ($file in $sourceFiles) {

    $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart("\")
    $target   = Join-Path $TargetRoot $relative

    $sourceHash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash

    if (!(Test-Path $target -PathType Leaf)) {
        $status = "NEW"
        $targetHash = ""
    }
    else {
        $targetHash = (Get-FileHash $target -Algorithm SHA256).Hash

        if ($sourceHash -eq $targetHash) {
            $status = "SAME"
        }
        else {
            $status = "CHANGED"
        }
    }

    [PSCustomObject]@{
        Status     = $status
        Path       = $relative
        SourceHash = $sourceHash
        TargetHash = $targetHash
    }
}

$newCount     = @($plan | Where-Object Status -eq "NEW").Count
$changedCount = @($plan | Where-Object Status -eq "CHANGED").Count
$sameCount    = @($plan | Where-Object Status -eq "SAME").Count

Write-Host ""
Write-Host "SOURCE FILES : $($sourceFiles.Count)"
Write-Host "NEW          : $newCount"
Write-Host "CHANGED      : $changedCount"
Write-Host "SAME         : $sameCount"


#
# 3. Dry-run is the default
#
if (!$Apply) {

    Write-Host ""
    Write-Host "MODE : DRY RUN" -ForegroundColor Yellow
    Write-Host "No production files were modified."

    $plan |
    Where-Object Status -ne "SAME" |
    Format-Table Status, Path -AutoSize

    Write-Host ""
    Write-Host "STATUS : JUNGLE_DEPLOY_DRYRUN_PASS" -ForegroundColor Green

    return
}


#
# 4. Apply deployment
#
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

$backupPath = Join-Path $BackupRoot $stamp

New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

foreach ($item in ($plan | Where-Object Status -ne "SAME")) {

    $source = Join-Path $SourceRoot $item.Path
    $target = Join-Path $TargetRoot $item.Path

    $targetDir = Split-Path $target -Parent

    if (!(Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    if (Test-Path $target -PathType Leaf) {

        $backupFile = Join-Path $backupPath $item.Path
        $backupDir  = Split-Path $backupFile -Parent

        if (!(Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }

        Copy-Item $target $backupFile -Force
    }

    Copy-Item $source $target -Force
}


#
# 5. Post-deployment SHA256 verification
#
$verifyFailures = @()

foreach ($file in $sourceFiles) {

    $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart("\")
    $target   = Join-Path $TargetRoot $relative

    if (!(Test-Path $target -PathType Leaf)) {
        $verifyFailures += "MISSING : $relative"
        continue
    }

    $sourceHash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash
    $targetHash = (Get-FileHash $target -Algorithm SHA256).Hash

    if ($sourceHash -ne $targetHash) {
        $verifyFailures += "HASH_FAIL : $relative"
    }
}

if ($verifyFailures.Count -gt 0) {

    Write-Host ""
    Write-Host "DEPLOY VERIFY : FAIL" -ForegroundColor Red

    $verifyFailures

    throw "DEPLOY VERIFICATION FAILED."
}

$logPath = Join-Path $LogRoot "deploy_$stamp.csv"
$plan | Export-Csv $logPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "DEPLOY VERIFY : PASS" -ForegroundColor Green
Write-Host "BACKUP        : $backupPath"
Write-Host "LOG           : $logPath"
Write-Host "STATUS        : JUNGLE_DEPLOY_COMPLETE" -ForegroundColor Green
