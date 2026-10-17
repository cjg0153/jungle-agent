param(
    [Parameter(Position=0, ValueFromRemainingArguments=$true)]
    [string[]]$RequestParts,

    [switch]$RouteOnly
)

$ErrorActionPreference = 'Stop'

# [ROUTER-V1-BEGIN]

$LegacyPath  = "C:\AI\jungle.ps1"
$GeneralPath = "C:\AI\jungle_v11.ps1"
$SmartPath   = "C:\AI\jungle_smart\jungle_smart.ps1"
$FastPath    = "C:\AI\run_verified_text_v3_fast.ps1"
$PlannerPath = "C:\AI\scripts\run_intent_artifact_planner_v11.ps1"

$ExpectedSha = @{
    LEGACY  = "ECF049108FB2EE1745C9FCA44D9084195F84BA3A20E4D97946C3C741179EE557"
    GENERAL = "694E1B8DC498E15D3563216F083AA9178F4FB7B5880473F9DDF4503604F819EF"
    SMART   = "1EC9D3B6EFB42ABFE3D9E41C984B38B4CDA29C284B7DBA2CB9EF048E1EAFC04C"
    FAST    = "E10315B4A02EC5E96FEF89601CF5D9C7F2DD0AF92440EE5E683E93E49A82BA17"
    PLANNER = "B56C7B90D97C0E0A077A7FEC674D92B8AE7FB39FB0F69696719BF35C13391A66"
}

function Get-CurrentSha {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Operational dependency missing: $Path"
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Assert-Dependency {
    param(
        [string]$Name,
        [string]$Path
    )

    $actual = Get-CurrentSha $Path
    $expect = $ExpectedSha[$Name]

    if ($actual -ne $expect) {
        throw "FAIL-CLOSED: $Name dependency hash changed. Expected=$expect Actual=$actual Path=$Path"
    }
}

function Invoke-JungleTarget {
    param(
        [string]$Path,
        [string]$Request
    )

    $cmd = Get-Command $Path -ErrorAction Stop

    if ($cmd.Parameters.ContainsKey('Request')) {
        & $Path -Request $Request
    }
    elseif ($cmd.Parameters.ContainsKey('Prompt')) {
        & $Path -Prompt $Request
    }
    elseif ($cmd.Parameters.ContainsKey('Text')) {
        & $Path -Text $Request
    }
    else {
        & $Path $Request
    }

    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
        throw "Target returned exit code $LASTEXITCODE : $Path"
    }
}

function Get-JungleRoute {
    param([string]$Request)

    $r = $Request.Trim()

    if ([string]::IsNullOrWhiteSpace($r)) {
        throw "Request is empty."
    }

    # Explicit Smart remains available, but now points to OUR Jungle Smart.
    if ($r -match '^(?i)smart(?:\s+|$)') {
        return [pscustomobject]@{
            Route = 'SMART_DIRECT'
            Target = $SmartPath
            Request = ($r -replace '^(?i)smart\s*','')
            File = $null
        }
    }

    $filePattern = '(?i)([A-Z]:\\.+?\.(?:pptx?|docx?|xlsx?|xlsm|xls|hwp|html?|md|txt|json|xml|ya?ml|css|js|ts|jsx|tsx|py|ps1|sql|csv))'
    $m = [regex]::Match($r, $filePattern)

    $file = $null
    $ext = $null

    if ($m.Success) {
        $file = $m.Groups[1].Value
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
    }

    # PPT reflow intent -> Jungle Smart
    if ($file -and $ext -eq '.pptx' -and
        $r -match '(?i)\breflow\b|재배치|레이아웃\s*(?:재구성|재배치)|구조\s*재배치') {

        return [pscustomobject]@{
            Route = 'PPT_REFLOW_SMART'
            Target = $SmartPath
            Request = $r
            File = $file
        }
    }

    # General text/code/web files -> v1.1 Source Protection
    $generalExt = @(
        '.html','.htm',
        '.md','.txt','.json','.xml','.yaml','.yml',
        '.css','.js','.ts','.jsx','.tsx',
        '.py','.ps1','.sql','.csv'
    )

    if ($file -and $generalExt -contains $ext) {
        return [pscustomobject]@{
            Route = 'GENERAL_FILE_V11'
            Target = $GeneralPath
            Request = $r
            File = $file
        }
    }

    # Existing office path is preserved for now.
    # We do NOT rewrite or patch its stable behavior here.
    $officeExt = @(
        '.ppt','.pptx',
        '.doc','.docx',
        '.xls','.xlsx','.xlsm',
        '.hwp'
    )

    if ($file -and $officeExt -contains $ext) {
        return [pscustomobject]@{
            Route = 'OFFICE_EXISTING_SAFE'
            Target = $LegacyPath
            Request = $r
            File = $file
        }
    }

    # ============================================================
    # JUNGLE OPERATIONAL v1.1 - INTENT / ARTIFACT PLANNER
    #
    # Fileless creation/show requests are planned BEFORE TEXT_FAST.
    # Planner decides whether the result should remain text or become
    # a durable artifact. Existing file routes remain untouched.
    # ============================================================

    $filelessCreationIntent =
        ($r -match '(?i)만들어\s*줘|만들어줘|만들어\s*주세요|생성해\s*줘|생성해\s*주세요') -or
        ($r -match '(?i)작성해\s*줘|작성해\s*주세요|그려\s*줘|그려\s*주세요') -or
        ($r -match '(?i)제작해\s*줘|구현해\s*줘|파일로\s*(?:만들|작성|생성)') -or
        ($r -match '(?i)산출물(?:로|\s*만들)|HTML로\s*보여|Mermaid로\s*보여|마크다운으로\s*보여')

    if (
        -not $file `
        -and $filelessCreationIntent
    ) {
        return [pscustomobject]@{
            Route   = 'INTENT_ARTIFACT_V11'
            Target  = $PlannerPath
            Request = $r
            File    = $null
        }
    }

    # No file -> existing FAST verified text path
    if (-not $file) {
        return [pscustomobject]@{
            Route = 'TEXT_FAST'
            Target = $FastPath
            Request = $r
            File = $null
        }
    }

    # Unknown file extension must not silently fall through.
    throw "FAIL-CLOSED: unsupported file type or route. File=$file"
}

$request = ($RequestParts -join ' ').Trim()

$route = Get-JungleRoute -Request $request

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " JUNGLE OPERATIONAL ROUTER v1" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Route   : $($route.Route)"
Write-Host "Target  : $($route.Target)"

if ($route.File) {
    Write-Host "File    : $($route.File)"
}

if ($RouteOnly) {
    Write-Host ""
    Write-Host "ROUTE ONLY - NO AI / NO FILE MODIFICATION" -ForegroundColor Yellow
    Write-Host "STATUS : JUNGLE_OPERATIONAL_ROUTE_PASS" -ForegroundColor Green
    return
}

switch ($route.Route) {
    'GENERAL_FILE_V11' {
        Assert-Dependency -Name 'GENERAL' -Path $GeneralPath
    }

    'SMART_DIRECT' {
        Assert-Dependency -Name 'SMART' -Path $SmartPath
    }

    'PPT_REFLOW_SMART' {
        Assert-Dependency -Name 'SMART' -Path $SmartPath
    }

    'OFFICE_EXISTING_SAFE' {
        Assert-Dependency -Name 'LEGACY' -Path $LegacyPath
    }

    'INTENT_ARTIFACT_V11' {
        Assert-Dependency -Name 'PLANNER' -Path $PlannerPath
    }

    'TEXT_FAST' {
        Assert-Dependency -Name 'FAST' -Path $FastPath
    }

    default {
        throw "Unknown route: $($route.Route)"
    }
}

Write-Host ""
Write-Host "Dependency integrity : PASS" -ForegroundColor Green
Write-Host "Executing target..." -ForegroundColor Yellow
Write-Host ""

Invoke-JungleTarget -Path $route.Target -Request $route.Request

# [ROUTER-V1-END]
