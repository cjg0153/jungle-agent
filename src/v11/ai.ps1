param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$Prompt,

    [switch]$Claude,
    [switch]$Codex,
    [switch]$Route,
    [switch]$XHigh
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$RoutingPath = "C:\AI\config\routing.json"

if (!(Test-Path $RoutingPath)) {
    throw "routing.json 파일을 찾을 수 없습니다: $RoutingPath"
}

$routing = Get-Content `
    -LiteralPath $RoutingPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json


function Invoke-JevRoute {

    param(
        [Parameter(Mandatory=$true)]
        [string]$RequestText
    )

    if (-not $env:TYPESAFE_API_KEY) {
        throw "TYPESAFE_API_KEY 환경변수가 없습니다."
    }

    $jevHeaders = @{
        Authorization = "Bearer $env:TYPESAFE_API_KEY"
    }

    $jevBody = @{
        state = $RequestText
        model = $routing.defaults.jev_model

        questions = @{

            workflow = @{
                type = "choice"
                instructions = "사용자의 요청에서 가장 핵심적인 업무 유형 하나를 선택한다."

                criteria = @{
                    BRIEF     = "임원 보고, 보고서, 브리프, PPT 작성이 핵심"
                    REVIEW    = "기존 문서나 자료의 검토, 피드백, QA가 핵심"
                    RESEARCH  = "외부 자료 검색, 사례 조사, 시장 조사, 최신 정보 확인이 핵심"
                    COMPLY    = "규제, 법률, 보안, 컴플라이언스 검토가 핵심"
                    MEETING   = "회의록, 회의 정리, 액션아이템 도출이 핵심"
                    MESSAGE   = "이메일, 메신저, 공지문 작성이 핵심"
                    EDU       = "교육자료, 설명자료, 학습 콘텐츠 제작이 핵심"
                    MOTION    = "모션그래픽, 발표 영상, 움직이는 HTML 제작이 핵심"
                    PROTOTYPE = "웹 프로토타입, 목업, UI 구현이 핵심"
                    DATA      = "엑셀, CSV, 데이터 가공 및 분석 자체가 핵심"
                    SYSTEM    = "코드, 자동화, CLI, 시스템 구축이 핵심"
                    OTHER     = "위 범주에 명확히 해당하지 않음"
                }
            }

            artifact = @{
                type = "choice"
                instructions = "사용자가 최종적으로 원하는 산출물 형식을 선택한다."

                criteria = @{
                    XLSX = "Excel 스프레드시트"
                    PPTX = "PowerPoint 프레젠테이션"
                    DOCX = "Word 문서"
                    PDF  = "PDF"
                    HTML = "HTML 또는 웹 화면"
                    MD   = "Markdown"
                    TEXT = "일반 텍스트"
                    NONE = "특정 파일 산출물이 없음"
                }
            }
        }
    }

    $jevJson = $jevBody | ConvertTo-Json -Depth 10
    $jevBytes = [System.Text.Encoding]::UTF8.GetBytes($jevJson)

    $jevResult = Invoke-RestMethod `
        -Uri "https://api.typesafe.ai/v1/systemone" `
        -Headers $jevHeaders `
        -Method Post `
        -ContentType "application/json; charset=utf-8" `
        -Body $jevBytes

    return [PSCustomObject]@{
        Workflow           = $jevResult.answers.workflow.choice
        WorkflowConfidence = [double]$jevResult.answers.workflow.confidence
        Artifact           = $jevResult.answers.artifact.choice
        ArtifactConfidence = [double]$jevResult.answers.artifact.confidence
    }
}


Write-Host ""
Write-Host "========================================"
Write-Host " Jungle Router"
Write-Host "========================================"
Write-Host ""
Write-Host "요청:"
Write-Host $Prompt
Write-Host ""


# === JUNGLE PHASE2 SECURITY GATE START ===
. "C:\AI\scripts\jungle_security_gate_v1.ps1"

$script:JungleOriginalPromptPhase2 = $Prompt
$securityResultPhase2 = Invoke-JungleSecurityGate -Prompt $Prompt

Write-Host ""
Write-Host "=== JUNGLE SECURITY GATE v1 ==="
Write-Host "Findings :" $securityResultPhase2.FindingCount

if ($securityResultPhase2.Findings.Count -gt 0) {
    Write-Host "Types    :" ($securityResultPhase2.Findings -join ", ")
}

if ($securityResultPhase2.Block) {
    Write-Host "STATUS: SECURITY_BLOCKED"
    Write-Host "Reason: explicit credential/private-key pattern detected."
    return
}

$routeResult = Invoke-JevRoute -RequestText $securityResultPhase2.RouterPrompt

# Restore the full local request after Jev routing.
$Prompt = $script:JungleOriginalPromptPhase2
Write-Host "STATUS: SECURITY_GATE_PASSED"
# === JUNGLE PHASE2 SECURITY GATE END ===

Write-Host "=== JEV ROUTING RESULT ==="
Write-Host "Workflow :" $routeResult.Workflow
Write-Host "Confidence:" $routeResult.WorkflowConfidence
Write-Host "Artifact :" $routeResult.Artifact
Write-Host "Confidence:" $routeResult.ArtifactConfidence
Write-Host ""



# ============================================================

# ============================================================
# === JUNGLE PHASE2 V5 OVERLAY START ===
. "C:\AI\scripts\jungle_core_v5_phase2.ps1"

$phase2DryRun = ($env:JUNGLE_DRYRUN -eq "1")
$v5ResultPhase2 = Invoke-JungleCoreV5Phase2 -Prompt $Prompt -DryRun $phase2DryRun
$v5FinalPhase2 = @($v5ResultPhase2)[-1]

if (
    $null -ne $v5FinalPhase2 -and
    $v5FinalPhase2.PSObject.Properties.Name -contains "Handled" -and
    $v5FinalPhase2.Handled
) {
    return
}
# === JUNGLE PHASE2 V5 OVERLAY END ===
# JUNGLE CORE v4 PRIMARY HOOK
# Supported intents return here before legacy overrides/handlers.
# Unsupported intents fall through to legacy compatibility code.
# ============================================================

$coreV4Path = "C:\AI\scripts\jungle_core_v4.ps1"

if (Test-Path -LiteralPath $coreV4Path -PathType Leaf) {

    . $coreV4Path

    $v4DryRun = ($env:JUNGLE_DRYRUN -eq "1")

    $v4Result = Invoke-JungleCoreV4 `
        -Prompt $Prompt `
        -RouteResult $routeResult `
        -DryRun $v4DryRun `
        -XHigh ([bool]$XHigh)

    $v4Final = @($v4Result)[-1]

    if ($null -ne $v4Final -and $v4Final.PSObject.Properties.Name -contains "Handled") {
        if ($v4Final.Handled) {
            return
        }
    }
}

# ============================================================
# LEGACY COMPATIBILITY PATH
# Reached only when Core v4 does not handle the request.
# ============================================================
# LOCAL INTENT OVERRIDE
#
# 파일경로 + 명확한 수정 동사가 있으면
# Jev Confidence와 관계없이 MODIFY 요청으로 확정한다.
# 현재 Workflow taxonomy에는 MODIFY가 없으므로
# REVIEW를 기반 Workflow로 사용하고 실제 실행은
# GENERIC_FILE_MODIFY Handler가 담당한다.
# ============================================================

$localModifyIntent = $false

$preGateFilePattern = '(?i)([A-Z]:\\.+?\.(?:html?|pptx?|docx?|pdf|xlsx?|xls|csv|txt|md))'
$preGateFileMatch = [regex]::Match($Prompt, $preGateFilePattern)

if ($preGateFileMatch.Success) {

    $preGatePath = $preGateFileMatch.Groups[1].Value.Trim()

    $hasModifyVerb =
        ($Prompt -match "수정해줘") -or
        ($Prompt -match "수정해 줘") -or
        ($Prompt -match "고쳐줘") -or
        ($Prompt -match "바꿔줘") -or
        ($Prompt -match "반영해줘") -or
        ($Prompt -match "다듬어줘") -or
        ($Prompt -match "추가해줘") -or
        ($Prompt -match "삭제해줘") -or
        ($Prompt -match "변경해줘")

    if (
        (Test-Path -LiteralPath $preGatePath -PathType Leaf) -and
        $hasModifyVerb
    ) {

        $localModifyIntent = $true

        # MODIFY는 실행 Intent이며 현재 Jev Workflow에는 없으므로
        # 가장 가까운 REVIEW를 기반 Workflow로 사용
        $routeResult.Workflow = "REVIEW"
        $routeResult.WorkflowConfidence = 1.0

        Write-Host "=== LOCAL INTENT OVERRIDE ==="
        Write-Host "Intent   : MODIFY"
        Write-Host "File     :" $preGatePath
        Write-Host "Workflow : REVIEW (base)"
        Write-Host ""
    }
}


# ============================================================
# LOCAL CREATE INTENT OVERRIDE
# 입력 파일을 참고하여 새 Word 문서를 만드는 요청을
# Jev Confidence와 관계없이 로컬에서 확정한다.
# ============================================================

$localCreateIntent = $false
$localCreateArtifact = $null

$createFilePattern = '(?i)([A-Z]:\\.+?\.(?:html?|pptx?|docx?|pdf|xlsx?|xls|csv|txt|md))'
$createFileMatch = [regex]::Match($Prompt, $createFilePattern)

if ($createFileMatch.Success) {

    $createSourcePath = $createFileMatch.Groups[1].Value.Trim()

    $hasCreateVerb =
        ($Prompt -match "새.*만들어줘") -or
        ($Prompt -match "새로.*만들어줘") -or
        ($Prompt -match "작성해줘") -or
        ($Prompt -match "생성해줘") -or
        ($Prompt -match "문서.*만들어줘") -or
        ($Prompt -match "보고서.*만들어줘")

    $wantsDocx =
        ($Prompt -match "Word") -or
        ($Prompt -match "워드") -or
        ($Prompt -match "DOCX") -or
        ($Prompt -match "문서")

    if (
        (Test-Path -LiteralPath $createSourcePath -PathType Leaf) -and
        $hasCreateVerb -and
        $wantsDocx
    ) {

        $localCreateIntent = $true
        $localCreateArtifact = "DOCX"

        # 현재 Workflow taxonomy에는 CREATE가 없으므로
        # BRIEF를 기반 Workflow로 사용한다.
        $routeResult.Workflow = "BRIEF"
        $routeResult.WorkflowConfidence = 1.0

        # 사용자가 Word 문서를 명시했으므로 출력 Artifact 확정
        $routeResult.Artifact = "DOCX"
        $routeResult.ArtifactConfidence = 1.0

        Write-Host "=== LOCAL INTENT OVERRIDE ==="
        Write-Host "Intent   : CREATE"
        Write-Host "Source   :" $createSourcePath
        Write-Host "Artifact : DOCX"
        Write-Host ""
    }
}



# Confidence Gate
$threshold = [double]$routing.defaults.confidence_threshold

# Workflow는 반드시 충분히 확신해야 한다.
if ($routeResult.WorkflowConfidence -lt $threshold) {
    Write-Host "ROUTE WARNING: Workflow Confidence가 기준보다 낮습니다."
    Write-Host "Threshold :" $threshold
    Write-Host "STATUS: ROUTE_NEEDS_CONFIRMATION"
    return
}


# ============================================================
# INPUT FILE DETECTION
# 기존 파일 경로와 출력 Artifact를 분리한다.
# ============================================================

$detectedFilePath = $null
$detectedInputType = $null

$filePattern = '(?i)([A-Z]:\\.+?\.(?:html?|pptx?|docx?|pdf|xlsx?|xls|csv|txt|md))'

$fileMatch = [regex]::Match(
    $Prompt,
    $filePattern
)

if ($fileMatch.Success) {

    $candidatePath = $fileMatch.Groups[1].Value.Trim()

    if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {

        $detectedFilePath = $candidatePath

        $detectedInputType = `
            [System.IO.Path]::GetExtension($candidatePath).
            TrimStart(".").
            ToUpperInvariant()

        Write-Host "=== INPUT FILE DETECTED ==="
        Write-Host "File :" $detectedFilePath
        Write-Host "Type :" $detectedInputType
        Write-Host ""
    }
}


# ============================================================
# ARTIFACT CONFIDENCE GATE
#
# REVIEW 요청에서 발견된 HTML/PPTX/PDF/XLSX 등의 파일은
# 출력물이 아니라 입력 파일일 수 있다.
# ============================================================

if ($routeResult.ArtifactConfidence -lt $threshold) {

    if ($routeResult.Artifact -eq "NONE") {

        Write-Host "Artifact Notice: 산출물 형식이 지정되지 않았습니다. NONE으로 계속합니다."
        Write-Host ""
    }
    elseif (
        $routeResult.Workflow -eq "REVIEW" -and
        $null -ne $detectedFilePath
    ) {

        Write-Host "Artifact Notice:"
        Write-Host "기존 파일 검토 요청으로 판단했습니다."
        Write-Host "Artifact 값은 입력 파일 형식으로 취급하고 계속 진행합니다."
        Write-Host ""
    }
    else {

        Write-Host "ARTIFACT WARNING: 산출물 형식 판단 Confidence가 낮습니다."
        Write-Host "Threshold :" $threshold
        Write-Host "STATUS: ARTIFACT_NEEDS_CONFIRMATION"
        return
    }
}

# Workflow 설정 조회
$workflowConfig = $routing.workflows.($routeResult.Workflow)

if ($null -eq $workflowConfig) {
    throw "routing.json에 Workflow 설정이 없습니다: $($routeResult.Workflow)"
}


$selectedPrimary = $workflowConfig.primary
$selectedSecondary = $workflowConfig.secondary


# 수동 강제 지정
if ($Claude) {
    $selectedPrimary = "claude"
    $selectedSecondary = "codex"
}

if ($Codex) {
    $selectedPrimary = "codex"
    $selectedSecondary = "claude"
}


# Artifact Handler
$artifactConfig = $routing.artifact_handlers.($routeResult.Artifact)


# Special Route 검색
$specialRoute = $routing.special_routes |
    Where-Object {
        $_.workflow -eq $routeResult.Workflow -and
        $_.artifact -eq $routeResult.Artifact
    } |
    Select-Object -First 1


Write-Host "=== JUNGLE EXECUTION PLAN ==="
Write-Host "Primary   :" $selectedPrimary
Write-Host "Secondary :" $selectedSecondary
Write-Host "Web Search:" $workflowConfig.web_search
Write-Host "Quality   :" $workflowConfig.quality_priority

if ($null -ne $artifactConfig) {

    if ($artifactConfig.content_engine) {
        Write-Host "Content   :" $artifactConfig.content_engine
    }

    if ($artifactConfig.file_executor) {
        Write-Host "Executor  :" $artifactConfig.file_executor
    }

    if ($artifactConfig.final_qa) {
        Write-Host "Final QA  :" $artifactConfig.final_qa
    }
}

Write-Host ""

if ($null -ne $specialRoute) {

    Write-Host "Special Route:"

    foreach ($step in $specialRoute.steps) {
        Write-Host " ->" $step
    }

    Write-Host ""
}


if ($XHigh) {
    Write-Host "Codex Reasoning: XHIGH"
}
else {
    Write-Host "Codex Reasoning:" $routing.defaults.codex_reasoning
}


Write-Host ""

if ($Route) {
    Write-Host "STATUS: ROUTED_ONLY"
    return
}

# ============================================================

# ============================================================

# ============================================================

# ============================================================
# GENERIC DOCX CREATE HANDLER
# 참고 파일을 기반으로 새 Word 문서를 생성한다.
# ============================================================

if (
    $localCreateIntent -eq $true -and
    $localCreateArtifact -eq "DOCX"
) {

    Write-Host ""
    Write-Host "=== JUNGLE HANDLER ==="
    Write-Host "Handler : GENERIC_DOCX_CREATE"
    Write-Host "Source  :" $createSourcePath
    Write-Host "Output  : DOCX"
    Write-Host ""

    if ($env:JUNGLE_DRYRUN -eq "1") {
        Write-Host "STATUS: DOCX_CREATE_HANDLER_MATCHED_DRYRUN"
        return
    }

    $createScript = "C:\AI\scripts\run_docx_create.ps1"

    if (!(Test-Path -LiteralPath $createScript -PathType Leaf)) {
        throw "DOCX Create Engine을 찾을 수 없습니다: $createScript"
    }

    & $createScript `
        -SourcePath $createSourcePath `
        -UserRequest $Prompt

    if (-not $?) {
        throw "GENERIC_DOCX_CREATE 실행에 실패했습니다."
    }

    return
}

# GENERIC FILE MODIFY HANDLER
# 기존 파일을 수정해달라는 요청이면 복사본을 만들어 수정한다.
# ============================================================

$isModifyRequest =
    ($null -ne $detectedFilePath) -and
    (
        ($localModifyIntent -eq $true) -or
        ($Prompt -match "수정해줘") -or
        ($Prompt -match "수정해 줘") -or
        ($Prompt -match "고쳐줘") -or
        ($Prompt -match "바꿔줘") -or
        ($Prompt -match "반영해줘") -or
        ($Prompt -match "다듬어줘") -or
        ($Prompt -match "추가해줘") -or
        ($Prompt -match "삭제해줘") -or
        ($Prompt -match "변경해줘")
    )

if ($isModifyRequest) {

    Write-Host ""
    Write-Host "=== JUNGLE HANDLER ==="
    Write-Host "Handler : GENERIC_FILE_MODIFY"
    Write-Host "Input   :" $detectedFilePath
    Write-Host ""

    if ($env:JUNGLE_DRYRUN -eq "1") {
        Write-Host "STATUS: FILE_MODIFY_HANDLER_MATCHED_DRYRUN"
        return
    }

    # 파일 형식별 MODIFY 엔진 선택
    switch ($detectedInputType) {

        "HTML" {
            $modifyScript = "C:\AI\scripts\run_file_modify.ps1"
            $modifyEngine = "TEXT_HTML_MODIFY"
        }

        "HTM" {
            $modifyScript = "C:\AI\scripts\run_file_modify.ps1"
            $modifyEngine = "TEXT_HTML_MODIFY"
        }

        "MD" {
            $modifyScript = "C:\AI\scripts\run_file_modify.ps1"
            $modifyEngine = "TEXT_MODIFY"
        }

        "TXT" {
            $modifyScript = "C:\AI\scripts\run_file_modify.ps1"
            $modifyEngine = "TEXT_MODIFY"
        }

        "PPTX" {
            $modifyScript = "C:\AI\scripts\run_pptx_modify.ps1"
            $modifyEngine = "POWERPOINT_COM_MODIFY"
        }

        "PPT" {
            $modifyScript = "C:\AI\scripts\run_pptx_modify.ps1"
            $modifyEngine = "POWERPOINT_COM_MODIFY"
        }


        "DOCX" {
            $modifyScript = "C:\AI\scripts\run_docx_modify.ps1"
            $modifyEngine = "WORD_COM_MODIFY"
        }

        "DOC" {
            $modifyScript = "C:\AI\scripts\run_docx_modify.ps1"
            $modifyEngine = "WORD_COM_MODIFY"
        }

        default {
            throw "현재 MODIFY에서 아직 지원하지 않는 파일 형식입니다: $detectedInputType"
        }
    }

    if (!(Test-Path -LiteralPath $modifyScript -PathType Leaf)) {
        throw "Modify Engine을 찾을 수 없습니다: $modifyScript"
    }

    Write-Host "Engine  :" $modifyEngine
    Write-Host ""

    & $modifyScript `
        -FilePath $detectedFilePath `
        -UserRequest $Prompt

    if (-not $?) {
        throw "GENERIC_FILE_MODIFY 실행에 실패했습니다."
    }

    return
}

# GENERIC FILE REVIEW HANDLER
# REVIEW + 실제 입력 파일이 있으면 파일을 직접 읽어 검토한다.
# ============================================================

if (
    $routeResult.Workflow -eq "REVIEW" -and
    $null -ne $detectedFilePath
) {

    Write-Host ""
    Write-Host "=== JUNGLE HANDLER ==="
    Write-Host "Handler : GENERIC_FILE_REVIEW"
    Write-Host "Input   :" $detectedFilePath
    Write-Host ""

    # Handler 연결 테스트용
    if ($env:JUNGLE_DRYRUN -eq "1") {
        Write-Host "STATUS: FILE_REVIEW_HANDLER_MATCHED_DRYRUN"
        return
    }

    $fileReviewScript = "C:\AI\scripts\run_file_review.ps1"

    if (!(Test-Path -LiteralPath $fileReviewScript -PathType Leaf)) {
        throw "File Review Handler를 찾을 수 없습니다: $fileReviewScript"
    }

    & powershell.exe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $fileReviewScript `
        -FilePath $detectedFilePath `
        -UserRequest $Prompt

    if ($LASTEXITCODE -ne 0) {
        throw "GENERIC_FILE_REVIEW 실행에 실패했습니다."
    }

    return
}

# JUNGLE EXECUTION DISPATCHER
# ============================================================

$targetChecks = @(
    ($Prompt -match "신한"),
    ($Prompt -match "하나"),
    ($Prompt -match "우리"),
    ($Prompt -match "국민"),
    ($Prompt -match "카카오"),
    (($Prompt -match "K뱅크") -or ($Prompt -match "케이뱅크")),
    ($Prompt -match "부산")
)

$targetCount = @(
    $targetChecks |
    Where-Object { $_ -eq $true }
).Count

$isBankAgentResearch =
    ($routeResult.Workflow -eq "RESEARCH") -and
    ($routeResult.Artifact -eq "XLSX") -and
    ($Prompt -match "(AI\s*Agent|AI\s*에이전트|에이전트)") -and
    ($Prompt -match "(엑셀|Excel|XLSX|조사표)") -and
    ($targetCount -eq 7)

if ($isBankAgentResearch) {

    Write-Host ""
    Write-Host "=== JUNGLE HANDLER ==="
    Write-Host "Handler : BANK_AGENT_RESEARCH_XLSX"
    Write-Host ""

    # Dispatcher 검증용: 실제 AI/Excel 실행 없이 Handler 선택만 확인
    if ($env:JUNGLE_DRYRUN -eq "1") {
        Write-Host "STATUS: HANDLER_MATCHED_DRYRUN"
        return
    }

    $pipeline =
        "C:\AI\scripts\run_bank_agent_pipeline.ps1"

    if (!(Test-Path -LiteralPath $pipeline)) {
        throw "Bank Agent pipeline script not found: $pipeline"
    }

    if ($XHigh) {

        & powershell.exe `
            -NoProfile `
            -ExecutionPolicy Bypass `
            -File $pipeline `
            -UserRequest $Prompt `
            -XHigh
    }
    else {

        & powershell.exe `
            -NoProfile `
            -ExecutionPolicy Bypass `
            -File $pipeline `
            -UserRequest $Prompt
    }

    if ($LASTEXITCODE -ne 0) {
        throw "BANK_AGENT_RESEARCH_XLSX pipeline failed."
    }

    return
}

Write-Host "STATUS: EXECUTION_PLAN_READY"

