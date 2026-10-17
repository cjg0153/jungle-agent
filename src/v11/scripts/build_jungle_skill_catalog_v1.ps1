$ErrorActionPreference = "Stop"

function Read-Meta {
    param([string]$FilePath,[string]$FallbackName,[string]$Kind)

    $raw = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    $name = $FallbackName
    $desc = ""

    if ($raw -match '(?m)^name:\s*"?([^"\r\n]+)"?\s*$') { $name = $matches[1].Trim() }

    if ($raw -match '(?m)^description:\s*(?:>\s*)?\r?\n((?:[ \t]+.*\r?\n?){1,10})') {
        $desc = (($matches[1] -split "\r?\n") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join " "
    }
    elseif ($raw -match '(?m)^description:\s*"?(.+?)"?\s*$') {
        $desc = $matches[1].Trim()
    }

    if ($desc.Length -gt 500) { $desc = $desc.Substring(0,500) }

    [PSCustomObject]@{
        kind = $Kind
        name = $name
        description = $desc
        path = $FilePath
    }
}

$items = @()

$claudeAgentDir = Join-Path $env:USERPROFILE ".claude\agents"
$claudeSkillDir = Join-Path $env:USERPROFILE ".claude\skills"
$codexSkillDir = Join-Path $env:USERPROFILE ".codex\skills"

if (Test-Path -LiteralPath $claudeAgentDir -PathType Container) {
    foreach ($f in @(Get-ChildItem -LiteralPath $claudeAgentDir -Filter "*.md" -File -ErrorAction SilentlyContinue)) {
        $items += (Read-Meta -FilePath $f.FullName -FallbackName $f.BaseName -Kind "CLAUDE_AGENT")
    }
}

foreach ($pair in @(
    @($claudeSkillDir,"CLAUDE_SKILL"),
    @($codexSkillDir,"CODEX_SKILL")
)) {
    $dir = $pair[0]
    $kind = $pair[1]

    if (Test-Path -LiteralPath $dir -PathType Container) {
        foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue)) {
            $sf = Join-Path $d.FullName "SKILL.md"
            if (Test-Path -LiteralPath $sf -PathType Leaf) {
                $items += (Read-Meta -FilePath $sf -FallbackName $d.Name -Kind $kind)
            }
        }
    }
}

$catalog = [PSCustomObject]@{
    built_at = (Get-Date).ToString("o")
    count = $items.Count
    items = $items
}

New-Item -ItemType Directory -Path "C:\AI\config" -Force | Out-Null
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText(
    "C:\AI\config\jungle_skill_catalog.json",
    ($catalog | ConvertTo-Json -Depth 8),
    $enc
)

Write-Host "Skill catalog items :" $items.Count
Write-Host "STATUS: JUNGLE_SKILL_CATALOG_BUILT"
