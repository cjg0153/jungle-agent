[CmdletBinding()]
param(
    [Parameter(Position=0,ValueFromRemainingArguments=$true)]
    [string[]]$InputArgs
)

$ErrorActionPreference = "Stop"

function Invoke-JungleSecurityPreflight {
    param([string]$Request)

    $securityScript = "C:\AI\scripts\jungle_security_gate_v1.ps1"
    if (Test-Path -LiteralPath $securityScript -PathType Leaf) {
        . $securityScript
        $s = Invoke-JungleSecurityGate -Prompt $Request

        if ($s.Block) {
            Write-Host ""
            Write-Host "STATUS: SECURITY_BLOCKED"
            Write-Host "명시적 credential/private-key 패턴이 포함되어 있어 실행하지 않습니다."
            return $false
        }
    }

    return $true
}

function Invoke-JungleClaudeDirect {
    param([string]$Request)

    if (-not (Invoke-JungleSecurityPreflight -Request $Request)) { return }

    $claudePath = $null
    $cmd = Get-Command claude -ErrorAction SilentlyContinue

    if ($null -ne $cmd) {
        $claudePath = $cmd.Source
    }
    elseif (Test-Path -LiteralPath "C:\Users\Zephyrus G14\.local\bin\claude.exe" -PathType Leaf) {
        $claudePath = "C:\Users\Zephyrus G14\.local\bin\claude.exe"
    }
    else {
        throw "Claude Code를 찾을 수 없습니다."
    }

    $tempDir = "C:\AI\work\temp"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $instructionPath = Join-Path $tempDir ("jungle_claude_direct_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($instructionPath,$Request,$enc)

    $short = "Read the UTF-8 instruction file at $instructionPath and execute the user's request. Use local file tools when needed."

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE -> CLAUDE CODE"
    Write-Host "========================================"
    Write-Host ""

    & $claudePath -p $short --allowedTools "Read" "Write" "Edit"

    if (-not $?) {
        throw "Claude Code 실행에 실패했습니다."
    }
}

function Invoke-JungleCodexDirect {
    param(
        [string]$Request,
        [switch]$Search
    )

    if (-not (Invoke-JungleSecurityPreflight -Request $Request)) { return }

    $codexPath = "C:\Users\Zephyrus G14\AppData\Local\Programs\OpenAI\Codex\bin\codex.exe"
    if (!(Test-Path -LiteralPath $codexPath -PathType Leaf)) {
        $cmd = Get-Command codex -ErrorAction SilentlyContinue
        if ($null -eq $cmd) {
            throw "Codex CLI를 찾을 수 없습니다."
        }
        $codexPath = $cmd.Source
    }

    $writeIntent =
        ($Request -match "수정해줘") -or
        ($Request -match "수정해 줘") -or
        ($Request -match "고쳐줘") -or
        ($Request -match "바꿔줘") -or
        ($Request -match "만들어줘") -or
        ($Request -match "생성해줘") -or
        ($Request -match "작성해줘") -or
        ($Request -match "반영해줘")

    $sandbox = "read-only"
    if ($writeIntent) { $sandbox = "workspace-write" }

    $workDir = (Get-Location).Path
    $m = [regex]::Match(
        $Request,
        '(?i)([A-Z]:\\.+?\.(?:pptx?|docx?|xlsx?|xlsm|xls|hwp|hwp|pdf|html?|csv|txt|md|ps1|py|js|ts|json))'
    )

    if ($m.Success) {
        $candidate = $m.Groups[1].Value.Trim()
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $workDir = Split-Path -LiteralPath $candidate -Parent
        }
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $codexPath
    $psi.WorkingDirectory = $workDir

    if ($Search) {
        $psi.Arguments = "--search exec --skip-git-repo-check --sandbox $sandbox --config model_reasoning_effort=medium -"
    }
    else {
        $psi.Arguments = "exec --skip-git-repo-check --sandbox $sandbox --config model_reasoning_effort=medium -"
    }

    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError = $false
    $psi.CreateNoWindow = $false

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    Write-Host ""
    Write-Host "========================================"
    if ($Search) { Write-Host " JUNGLE -> CODEX RESEARCH" }
    else { Write-Host " JUNGLE -> CODEX" }
    Write-Host "========================================"
    Write-Host "Sandbox :" $sandbox
    Write-Host "WorkDir :" $workDir
    Write-Host ""

    [void]$p.Start()

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $bytes = $utf8.GetBytes($Request)
    $p.StandardInput.BaseStream.Write($bytes,0,$bytes.Length)
    $p.StandardInput.BaseStream.Flush()
    $p.StandardInput.Close()

    $p.WaitForExit()

    if ($p.ExitCode -ne 0) {
        throw "Codex CLI 실행에 실패했습니다. ExitCode=$($p.ExitCode)"
    }
}

function Invoke-JungleAutoLite {
    param([string]$Request)

    if (-not (Invoke-JungleSecurityPreflight -Request $Request)) { return }

    $m = [regex]::Match(
        $Request,
        '(?i)([A-Z]:\\.+?\.(?:pptx?|docx?|xlsx?|xlsm|xls|hwp|hwp|pdf|html?|csv|txt|md))'
    )

    if (-not $m.Success) {
        Write-Host "AUTO: 파일 경로가 없어 Claude direct로 전환합니다."
        Invoke-JungleClaudeDirect -Request $Request
        return
    }

    $path = $m.Groups[1].Value.Trim()
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "파일을 찾을 수 없습니다: $path"
    }

    $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

    $isModify =
        ($Request -match "수정해줘") -or
        ($Request -match "수정해 줘") -or
        ($Request -match "고쳐줘") -or
        ($Request -match "바꿔줘") -or
        ($Request -match "반영해줘") -or
        ($Request -match "다듬어줘") -or
        ($Request -match "추가해줘") -or
        ($Request -match "삭제해줘") -or
        ($Request -match "변경해줘")

    $isReview =
        ($Request -match "봐줘") -or
        ($Request -match "검토") -or
        ($Request -match "리뷰") -or
        ($Request -match "점검") -or
        ($Request -match "분석해줘") -or
        ($Request -match "확인해줘")

    $isCreate =
        ($Request -match "새.*만들어줘") -or
        ($Request -match "새로.*만들어줘") -or
        ($Request -match "작성해줘") -or
        ($Request -match "생성해줘") -or
        ($Request -match "보고서.*만들어줘") -or
        ($Request -match "PPT.*만들어줘") -or
        ($Request -match "문서.*만들어줘")

    $wantsPpt =
        ($Request -match "(?i)\bPPT\b") -or
        ($Request -match "(?i)\bPPTX\b") -or
        ($Request -match "프레젠테이션")

    $wantsDocx =
        ($Request -match "(?i)\bWord\b") -or
        ($Request -match "워드") -or
        ($Request -match "(?i)\bDOCX\b")

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE AUTO LITE"
    Write-Host "========================================"
    Write-Host "Input :" $path
    Write-Host "Type  :" $ext
    Write-Host ""

    if ($ext -in @(".pptx",".ppt")) {
        if ($isModify) {
            & "C:\AI\scripts\run_pptx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Request
            return
        }

        if ($isCreate -and $wantsDocx) {
            & "C:\AI\scripts\run_docx_create_quality_v1.ps1" -SourcePath $path -UserRequest $Request
            return
        }

        if ($isCreate -or $wantsPpt) {
            $create = "C:\AI\scripts\run_pptx_create_v2_pipeline.ps1"
            $quality = "C:\AI\scripts\run_ppt_quality_pipeline_v1.ps1"
            $outputDir = "C:\AI\work\output"
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
            $started = Get-Date

            & $create -SourcePath $path -UserRequest $Request
            if (-not $?) { throw "PPT CREATE 실패" }

            $draft = @(Get-ChildItem -LiteralPath $outputDir -Filter "${baseName}_design_v2_*.pptx" -File |
                Where-Object {
                    $_.LastWriteTime -ge $started.AddSeconds(-2) -and
                    $_.Name -notmatch "_modified_" -and
                    $_.Name -notmatch "_FINAL_"
                } |
                Sort-Object LastWriteTime -Descending)[0]

            if ($null -eq $draft) { throw "새 PPT draft를 찾지 못했습니다." }

            & $quality -DraftPath $draft.FullName -SourcePath $path -UserRequest $Request
            return
        }

        if ($isReview) {
            & "C:\AI\scripts\run_codex_qa_gate_v3.ps1" -FilePath $path -UserRequest $Request
            return
        }
    }

    if ($ext -eq ".hwp") {

        $hwpPipeline = "C:\AI\scripts\run_hwp_modify_quality_v1.ps1"

        if (!(Test-Path -LiteralPath $hwpPipeline -PathType Leaf)) {
            throw "HWP 품질 파이프라인 누락: $hwpPipeline"
        }

        Write-Host ""
        Write-Host "AUTO HWP : HWP -> DOCX -> QA -> HWP"

        & $hwpPipeline `
            -FilePath $path `
            -UserRequest $Request

        if (-not $?) {
            throw "HWP quality pipeline 실패"
        }

        return
    }

    if ($ext -eq ".docx") {
        if ($isModify) {
            & "C:\AI\scripts\run_docx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Request
            return
        }

        if ($isCreate -and $wantsDocx) {
            & "C:\AI\scripts\run_docx_create_quality_v1.ps1" -SourcePath $path -UserRequest $Request
            return
        }

        if ($isReview) {
            & "C:\AI\scripts\run_office_codex_qa_v1.ps1" -FilePath $path -UserRequest $Request
            return
        }
    }

    if ($ext -in @(".xlsx",".xlsm",".xls")) {
        if ($isModify) {
            & "C:\AI\scripts\run_xlsx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Request
            return
        }

        if ($isReview) {
            & "C:\AI\scripts\run_office_codex_qa_v1.ps1" -FilePath $path -UserRequest $Request
            return
        }
    }

    Write-Host "AUTO: 전용 파이프라인이 없는 요청이라 Claude direct로 전환합니다."
    Invoke-JungleClaudeDirect -Request $Request
}

function Show-JungleHelp {
    Write-Host ""
    Write-Host "Jungle Lite 사용법"
    Write-Host ""
    Write-Host "  jungle '요청'              -> Claude 1차 + Codex 검증 + 필요 시 Claude 1회 수정"
    Write-Host "  jungle claude '요청'       -> Claude Code 직결"
    Write-Host "  jungle codex '요청'        -> Codex 직결"
    Write-Host "  jungle research '요청'     -> Codex + Web Search"
    Write-Host "  jungle auto '요청'         -> 기존 품질 파이프라인"
    Write-Host "  jungle smart '요청'        -> 기존 Jev/Router"
    Write-Host "  jungle logic               -> 로직트리"
    Write-Host "  jungle status              -> 상태"
    Write-Host "  jungle help                -> 도움말"
    Write-Host ""
}

function Show-JungleLogic {
    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE LITE LOGIC TREE"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "기본 요청      -> Claude(+관련 Skills) -> Codex(+관련 Skills) QA -> 필요 시 Claude 1회 수정 -> Codex 재QA"
    Write-Host "claude         -> Claude Code direct"
    Write-Host "codex          -> Codex direct"
    Write-Host "research       -> Codex --search"
    Write-Host "auto           -> 파일 품질 파이프라인"
    Write-Host "                  Claude 제작/수정"
    Write-Host "                  -> Codex QA"
    Write-Host "                  -> Triage"
    Write-Host "                  -> AUTO_FIX 1회"
    Write-Host "                  -> Re-QA"
    Write-Host "                  -> Acceptance/Rollback"
    Write-Host "smart          -> 기존 ai.ps1 -> Security Gate -> Jev -> Core"
    Write-Host ""
    Write-Host "원칙: 모든 기본 요청은 Claude 1차 + Codex 검증. Claude/Codex 단독은 명시할 때만. Jev=smart에서만."
    Write-Host ""
}

function Show-JungleStatus {
    $claude = Get-Command claude -ErrorAction SilentlyContinue
    $codexPath = "C:\Users\Zephyrus G14\AppData\Local\Programs\OpenAI\Codex\bin\codex.exe"

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE LITE STATUS"
    Write-Host "========================================"
    if ($null -ne $claude) { Write-Host "Claude     : OK" } else { Write-Host "Claude     : CHECK" }
    if (Test-Path -LiteralPath $codexPath -PathType Leaf) { Write-Host "Codex      : OK" } else { Write-Host "Codex      : CHECK" }
    if (Test-Path "C:\AI\scripts\run_ppt_quality_pipeline_v1.ps1") { Write-Host "Auto PPT   : OK" } else { Write-Host "Auto PPT   : CHECK" }
    if (Test-Path "C:\AI\scripts\run_office_quality_pipeline_v1.ps1") { Write-Host "Auto DOCX  : OK" } else { Write-Host "Auto DOCX  : CHECK" }
    if (Test-Path "C:\AI\scripts\run_xlsx_modify_quality_v1.ps1") { Write-Host "Auto XLSX  : OK" } else { Write-Host "Auto XLSX  : CHECK" }
    if (Test-Path "C:\AI\scripts\run_hwp_modify_quality_v1.ps1") { Write-Host "Auto HWP   : OK" } else { Write-Host "Auto HWP   : CHECK" }
    if (Test-Path "C:\AI\ai.ps1") { Write-Host "Legacy/Jev : AVAILABLE (smart only)" } else { Write-Host "Legacy/Jev : MISSING" }
    Write-Host ""
}

if ($null -eq $InputArgs -or $InputArgs.Count -eq 0) {
    Show-JungleHelp
    return
}

$mode = [string]$InputArgs[0]
$modeLower = $mode.ToLowerInvariant()

if ($modeLower -in @("help","도움말","-h","--help")) { Show-JungleHelp; return }
if ($modeLower -in @("logic","tree","로직","로직트리","설정","설정트리")) { Show-JungleLogic; return }
if ($modeLower -in @("status","상태")) { Show-JungleStatus; return }

$request = ""
if ($InputArgs.Count -gt 1) {
    $request = ($InputArgs[1..($InputArgs.Count-1)] -join " ").Trim()
}

if ($modeLower -eq "claude") {
    if ([string]::IsNullOrWhiteSpace($request)) { throw "Claude 요청이 비어 있습니다." }
    & "C:\AI\run_claude_skilled_v1.ps1" -Request $request
    return
}

if ($modeLower -eq "codex") {
    if ([string]::IsNullOrWhiteSpace($request)) { throw "Codex 요청이 비어 있습니다." }
    Invoke-JungleCodexDirect -Request $request
    return
}

if ($modeLower -in @("research","리서치","조사")) {
    if ([string]::IsNullOrWhiteSpace($request)) { throw "Research 요청이 비어 있습니다." }
    Invoke-JungleCodexDirect -Request $request -Search
    return
}

if ($modeLower -eq "auto") {
    if ([string]::IsNullOrWhiteSpace($request)) { throw "Auto 요청이 비어 있습니다." }
    Invoke-JungleAutoLite -Request $request
    return
}

if ($modeLower -in @("smart","route","legacy","jev")) {
    if ([string]::IsNullOrWhiteSpace($request)) { throw "Smart 요청이 비어 있습니다." }
    & "C:\AI\ai.ps1" $request
    return
}

$joined = ($InputArgs -join " ").Trim()

if (
    ($joined -match "로직트리") -or
    ($joined -match "로직 트리") -or
    ($joined -match "설정된 로직") -or
    ($joined -match "어떻게 동작") -or
    ($joined -match "작동 구조")
) {
    Show-JungleLogic
    return
}

if (
    ($joined -match "현재 상태") -or
    ($joined -match "설치 상태") -or
    ($joined -match "상태 보여")
) {
    Show-JungleStatus
    return
}

# Default = verified pipeline. No Jev.
$fileMatchDefault = [regex]::Match(
    $joined,
    '(?i)([A-Z]:\\.+?\.(?:pptx?|docx?|xlsx?|xlsm|xls|hwp))'
)

if ($fileMatchDefault.Success) {
    Invoke-JungleAutoLite -Request $joined
}
else {
    & "C:\AI\run_verified_text_v3_fast.ps1" -Request $joined
}