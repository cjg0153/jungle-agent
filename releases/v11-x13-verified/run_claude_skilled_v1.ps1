param(
    [Parameter(Mandatory=$true)][string]$Request
)

$ErrorActionPreference = "Stop"

function Get-ClaudeCatalog {
    $root = Join-Path $env:USERPROFILE ".claude"
    $lines = New-Object System.Collections.Generic.List[string]

    $agentDir = Join-Path $root "agents"
    if (Test-Path -LiteralPath $agentDir -PathType Container) {
        foreach ($f in @(Get-ChildItem -LiteralPath $agentDir -Filter "*.md" -File -ErrorAction SilentlyContinue | Select-Object -First 30)) {
            $raw = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            $name = $f.BaseName
            if ($raw -match '(?m)^name:\s*"?([^"\r\n]+)"?\s*$') { $name = $matches[1].Trim() }
            $desc = ""
            if ($raw -match '(?m)^description:\s*"?(.+?)"?\s*$') { $desc = $matches[1].Trim() }
            if ($desc.Length -gt 300) { $desc = $desc.Substring(0,300) }
            $lines.Add("AGENT $name :: $desc")
        }
    }

    $skillDir = Join-Path $root "skills"
    if (Test-Path -LiteralPath $skillDir -PathType Container) {
        foreach ($d in @(Get-ChildItem -LiteralPath $skillDir -Directory -ErrorAction SilentlyContinue | Select-Object -First 50)) {
            $skillFile = Join-Path $d.FullName "SKILL.md"
            if (!(Test-Path -LiteralPath $skillFile -PathType Leaf)) { continue }
            $raw = Get-Content -LiteralPath $skillFile -Raw -Encoding UTF8
            $name = $d.Name
            if ($raw -match '(?m)^name:\s*"?([^"\r\n]+)"?\s*$') { $name = $matches[1].Trim() }
            $desc = ""
            if ($raw -match '(?m)^description:\s*(?:>\s*)?\r?\n((?:[ \t]+.*\r?\n?){1,8})') {
                $desc = (($matches[1] -split "\r?\n") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join " "
            }
            elseif ($raw -match '(?m)^description:\s*"?(.+?)"?\s*$') { $desc = $matches[1].Trim() }
            if ($desc.Length -gt 350) { $desc = $desc.Substring(0,350) }
            $lines.Add("SKILL $name :: $desc")
        }
    }

    return ($lines -join "`r`n")
}

$claudePath = $null
$cmd = Get-Command claude -ErrorAction SilentlyContinue
if ($null -ne $cmd) { $claudePath = $cmd.Source }
elseif (Test-Path -LiteralPath "C:\Users\Zephyrus G14\.local\bin\claude.exe" -PathType Leaf) {
    $claudePath = "C:\Users\Zephyrus G14\.local\bin\claude.exe"
}
else { throw "Claude Code를 찾을 수 없습니다." }

$catalog = Get-ClaudeCatalog

$drawioAvailable = ($null -ne (Get-Command drawio -ErrorAction SilentlyContinue))
$mmdcAvailable = ($null -ne (Get-Command mmdc -ErrorAction SilentlyContinue))
$pythonAvailable =
    ($null -ne (Get-Command python -ErrorAction SilentlyContinue)) -or
    ($null -ne (Get-Command python3 -ErrorAction SilentlyContinue))

$tempDir = "C:\AI\work\temp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$requestPath = Join-Path $tempDir "claude_skilled_request_${stamp}.txt"

$encBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($requestPath,$Request,$encBom)

$instruction = @"
Read the UTF-8 request:
$requestPath

Execute the user's request.

INSTALLED CLAUDE AGENTS/SKILLS:
$catalog

SKILL POLICY:
- Select only agents/skills actually relevant to this request.
- For Korean technical Markdown documents (architecture, requirements, design, technical guide, structured research documentation), prefer the installed docs-writer-kr agent when available.
- For NEW diagrams, use drawio-figure when available. Do not use Mermaid for a new figure.
- Use mermaid-figure only when maintaining an existing .mmd/Mermaid figure.
- Do not force documentation skills onto PPTX, DOCX, XLSX, coding, or unrelated tasks unless their own descriptions clearly match.
- Follow the user's requested output format over a skill's default format.
- Do not invent facts.

LOCAL CAPABILITIES:
drawio CLI available: $drawioAvailable
mmdc available: $mmdcAvailable
Python available: $pythonAvailable

If a selected rendering workflow needs a missing executable:
- do not auto-install system software,
- produce the safe source/draft work that can be completed,
- clearly report what rendering step could not be executed.

Use local file tools when needed.
Output only the useful result/status, never chain-of-thought.
"@

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE -> CLAUDE + SKILLS"
Write-Host "========================================"

# Task is allowed because docs-writer-kr is a Claude subagent.
& $claudePath -p $instruction --allowedTools "Read" "Write" "Edit" "Task"

if (-not $?) { throw "Claude skilled 실행 실패" }
