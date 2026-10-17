# Jungle Core v4
Set-StrictMode -Version 2.0

function Get-JungleFileInfo {
    param([Parameter(Mandatory=$true)][string]$Prompt)
    $pattern = '(?i)([A-Z]:\\.+?\.(?:html?|pptx?|ppt|docx?|doc|pdf|xlsx?|xls|csv|txt|md))'
    $match = [regex]::Match($Prompt, $pattern)
    if (-not $match.Success) {
        return [PSCustomObject]@{ Found=$false; Path=$null; Extension=$null; Exists=$false }
    }
    $path = $match.Groups[1].Value.Trim()
    $ext = [System.IO.Path]::GetExtension($path).TrimStart(".").ToUpperInvariant()
    $exists = Test-Path -LiteralPath $path -PathType Leaf
    return [PSCustomObject]@{ Found=$true; Path=$path; Extension=$ext; Exists=[bool]$exists }
}

function Get-JungleExplicitOutputArtifact {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt
    )

    # 입력 파일경로의 확장자는 출력 형식 판단에서 제외
    $cleanPrompt = [regex]::Replace(
        $Prompt,
        '(?i)[A-Z]:\\.+?\.(?:html?|pptx?|ppt|docx?|doc|pdf|xlsx?|xls|csv|txt|md)',
        ' '
    )

    # 실제로 결과물을 만들거나 변환하는 요청만 출력 형식을 판별
    $actions = [regex]::Matches(
        $cleanPrompt,
        '(?i)만들어|작성해|생성해|변환해|바꿔|내보내|저장해'
    )

    if ($actions.Count -eq 0) {
        return $null
    }

    $actionIndex = $actions[$actions.Count - 1].Index

    $artifactPattern = '(?i)PowerPoint|파워포인트|PPTX?|프레젠테이션|Word|워드|DOCX?|Excel|엑셀|스프레드시트|XLSX?|PDF|HTML?|웹페이지|웹 페이지|Markdown|마크다운|\bMD\b'

    $candidates = @(
        [regex]::Matches($cleanPrompt, $artifactPattern) |
        Where-Object { $_.Index -lt $actionIndex }
    )

    if ($candidates.Count -eq 0) {
        return $null
    }

    $target = $candidates[$candidates.Count - 1].Value

    if ($target -match '(?i)PowerPoint|파워포인트|PPTX?|프레젠테이션') { return "PPTX" }
    if ($target -match '(?i)Word|워드|DOCX?') { return "DOCX" }
    if ($target -match '(?i)Excel|엑셀|스프레드시트|XLSX?') { return "XLSX" }
    if ($target -match '(?i)PDF') { return "PDF" }
    if ($target -match '(?i)HTML?|웹페이지|웹 페이지') { return "HTML" }
    if ($target -match '(?i)Markdown|마크다운|\bMD\b') { return "MD" }

    return $null
}
function Get-JungleIntent {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)]$FileInfo,
        $RouteResult = $null
    )

    if ($FileInfo.Found -and $FileInfo.Exists) {
        if ($Prompt -match '변환해줘|변환해 줘|변환해서|파일로 변환|형식으로 변환|포맷.*변환') {
            $explicit = Get-JungleExplicitOutputArtifact -Prompt $Prompt
            if ($null -ne $explicit -and $explicit -ne $FileInfo.Extension) { return "CONVERT" }
        }

        if ($Prompt -match '새.*만들어줘|새로.*만들어줘|새 문서|새로운 문서|작성해줘|생성해줘|요약.*문서.*만들어줘|보고서.*만들어줘|자료.*만들어줘') {
            $explicit = Get-JungleExplicitOutputArtifact -Prompt $Prompt
            if ($null -ne $explicit) { return "CREATE" }
        }

        if ($Prompt -match '수정해줘|수정해 줘|고쳐줘|바꿔줘|반영해줘|다듬어줘|추가해줘|삭제해줘|변경해줘|교체해줘') {
            return "MODIFY"
        }

        if ($Prompt -match '검토해줘|검토해 줘|분석해줘|분석해 줘|봐줘|의견을 줘|의견줘|평가해줘|점검해줘|확인해줘') {
            return "REVIEW"
        }
    }

    if ($Prompt -match '(?i)AI\s*Agent|AI\s*에이전트|에이전트') {
        $bankHits = @(
            ($Prompt -match '신한'),
            ($Prompt -match '하나'),
            ($Prompt -match '우리'),
            ($Prompt -match '국민'),
            ($Prompt -match '카카오'),
            (($Prompt -match 'K뱅크') -or ($Prompt -match '케이뱅크')),
            ($Prompt -match '부산')
        ) | Where-Object { $_ -eq $true }

        if ($bankHits.Count -eq 7 -and $Prompt -match '엑셀|Excel|XLSX|조사표') {
            return "BANK_AGENT_RESEARCH_XLSX"
        }
    }

    if ($null -ne $RouteResult) { return [string]$RouteResult.Workflow }
    return "OTHER"
}

function Get-JungleModifyEngine {
    param([Parameter(Mandatory=$true)][string]$Extension)

    switch ($Extension) {
        "HTML" { return [PSCustomObject]@{ Name="TEXT_HTML_MODIFY"; Script="C:\AI\scripts\run_file_modify.ps1" } }
        "HTM"  { return [PSCustomObject]@{ Name="TEXT_HTML_MODIFY"; Script="C:\AI\scripts\run_file_modify.ps1" } }
        "MD"   { return [PSCustomObject]@{ Name="TEXT_MODIFY"; Script="C:\AI\scripts\run_file_modify.ps1" } }
        "TXT"  { return [PSCustomObject]@{ Name="TEXT_MODIFY"; Script="C:\AI\scripts\run_file_modify.ps1" } }
        "PPTX" { return [PSCustomObject]@{ Name="POWERPOINT_COM_MODIFY"; Script="C:\AI\scripts\run_pptx_modify.ps1" } }
        "PPT"  { return [PSCustomObject]@{ Name="POWERPOINT_COM_MODIFY"; Script="C:\AI\scripts\run_pptx_modify.ps1" } }
        "DOCX" { return [PSCustomObject]@{ Name="WORD_COM_MODIFY"; Script="C:\AI\scripts\run_docx_modify.ps1" } }
        "DOC"  { return [PSCustomObject]@{ Name="WORD_COM_MODIFY"; Script="C:\AI\scripts\run_docx_modify.ps1" } }
        default { return $null }
    }
}

function Invoke-JungleCoreV4 {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        $RouteResult = $null,
        [bool]$DryRun = $false,
        [bool]$XHigh = $false
    )

    $fileInfo = Get-JungleFileInfo -Prompt $Prompt
    $intent = Get-JungleIntent -Prompt $Prompt -FileInfo $fileInfo -RouteResult $RouteResult
    $explicitOutput = Get-JungleExplicitOutputArtifact -Prompt $Prompt

    Write-Host ""
    Write-Host "=== JUNGLE CORE v4 ==="
    Write-Host "Intent :" $intent
    if ($fileInfo.Found) {
        Write-Host "Input  :" $fileInfo.Path
        Write-Host "Type   :" $fileInfo.Extension
        Write-Host "Exists :" $fileInfo.Exists
    }
    if ($null -ne $explicitOutput) { Write-Host "Output :" $explicitOutput }
    Write-Host ""

    if ($fileInfo.Found -and (-not $fileInfo.Exists)) {
        Write-Host "STATUS: INPUT_FILE_NOT_FOUND"
        return [PSCustomObject]@{ Handled=$true; Status="INPUT_FILE_NOT_FOUND" }
    }

    if ($intent -eq "REVIEW" -and $fileInfo.Exists) {
        $script = "C:\AI\scripts\run_file_review.ps1"
        Write-Host "Handler: GENERIC_FILE_REVIEW"
        if ($DryRun) {
            Write-Host "STATUS: V4_REVIEW_MATCHED_DRYRUN"
            return [PSCustomObject]@{ Handled=$true; Status="V4_REVIEW_MATCHED_DRYRUN" }
        }
        if (!(Test-Path -LiteralPath $script -PathType Leaf)) { throw "Review engine not found: $script" }
        & $script -FilePath $fileInfo.Path -UserRequest $Prompt
        if (-not $?) { throw "V4 REVIEW execution failed." }
        return [PSCustomObject]@{ Handled=$true; Status="REVIEW_COMPLETE" }
    }

    if ($intent -eq "MODIFY" -and $fileInfo.Exists) {
        $engine = Get-JungleModifyEngine -Extension $fileInfo.Extension
        if ($null -eq $engine) {
            Write-Host "STATUS: MODIFY_ENGINE_NOT_AVAILABLE"
            return [PSCustomObject]@{ Handled=$true; Status="MODIFY_ENGINE_NOT_AVAILABLE" }
        }
        Write-Host "Handler: GENERIC_FILE_MODIFY"
        Write-Host "Engine :" $engine.Name
        if ($DryRun) {
            Write-Host "STATUS: V4_MODIFY_MATCHED_DRYRUN"
            return [PSCustomObject]@{ Handled=$true; Status="V4_MODIFY_MATCHED_DRYRUN" }
        }
        if (!(Test-Path -LiteralPath $engine.Script -PathType Leaf)) { throw "Modify engine not found: $($engine.Script)" }
        & $engine.Script -FilePath $fileInfo.Path -UserRequest $Prompt
        if (-not $?) { throw "V4 MODIFY execution failed." }
        return [PSCustomObject]@{ Handled=$true; Status="MODIFY_COMPLETE" }
    }

    if ($intent -eq "CREATE" -and $fileInfo.Exists) {
        if ($explicitOutput -eq "DOCX") {
            $script = "C:\AI\scripts\run_docx_create.ps1"
            Write-Host "Handler: GENERIC_DOCX_CREATE"
            Write-Host "Engine : CLAUDE_PLAN + OPENXML_DOCX"
            if ($DryRun) {
                Write-Host "STATUS: V4_DOCX_CREATE_MATCHED_DRYRUN"
                return [PSCustomObject]@{ Handled=$true; Status="V4_DOCX_CREATE_MATCHED_DRYRUN" }
            }
            if (!(Test-Path -LiteralPath $script -PathType Leaf)) { throw "DOCX create engine not found: $script" }
            & $script -SourcePath $fileInfo.Path -UserRequest $Prompt
            if (-not $?) { throw "V4 DOCX CREATE execution failed." }
            return [PSCustomObject]@{ Handled=$true; Status="DOCX_CREATE_COMPLETE" }
        }

        if ($explicitOutput -eq "PPTX") {
            $script = "C:\AI\scripts\run_pptx_create_v2_pipeline.ps1"
            $qualityScript = "C:\AI\scripts\run_ppt_quality_pipeline_v1.ps1"

            Write-Host "Handler: GENERIC_PPTX_CREATE"
            Write-Host "Engine : PPT_DESIGN_V2_AUTO"
            Write-Host "Design : AUTO ANALYZE + LAYOUT CACHE + CLAUDE"
            Write-Host "Quality: CODEX QA v3 + TRIAGE + CLAUDE AUTO_FIX 1X + ACCEPTANCE"

            if ($DryRun) {
                Write-Host "STATUS: V4_PPTX_CREATE_QUALITY_MATCHED_DRYRUN"
                return [PSCustomObject]@{
                    Handled = $true
                    Status  = "V4_PPTX_CREATE_QUALITY_MATCHED_DRYRUN"
                }
            }

            if (!(Test-Path -LiteralPath $script -PathType Leaf)) {
                throw "PPTX create engine not found: $script"
            }

            if (!(Test-Path -LiteralPath $qualityScript -PathType Leaf)) {
                throw "PPT quality pipeline not found: $qualityScript"
            }

            $createStarted = Get-Date

            & $script -SourcePath $fileInfo.Path -UserRequest $Prompt

            if (-not $?) {
                throw "V4 PPTX CREATE execution failed."
            }

            $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($fileInfo.Path)

            $draftFile = Get-ChildItem "C:\AI\work\output" -Filter "${sourceBase}_design_v2_*.pptx" -File -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.LastWriteTime -ge $createStarted.AddSeconds(-2) -and
                    $_.Name -notmatch "_FINAL_" -and
                    $_.Name -notmatch "_modified_"
                } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($null -eq $draftFile) {
                throw "새로 생성된 PPTX draft를 찾지 못했습니다. Quality Pipeline을 시작하지 않습니다."
            }

            Write-Host ""
            Write-Host "=== JUNGLE PPT QUALITY AUTO START ==="
            Write-Host "Draft  :" $draftFile.FullName
            Write-Host "Source :" $fileInfo.Path
            Write-Host ""

            & $qualityScript `
                -DraftPath $draftFile.FullName `
                -SourcePath $fileInfo.Path `
                -UserRequest $Prompt

            if (-not $?) {
                throw "V4 PPT Quality Pipeline execution failed."
            }

            return [PSCustomObject]@{
                Handled = $true
                Status  = "PPTX_CREATE_QUALITY_COMPLETE"
            }
        }
        Write-Host "STATUS: CREATE_ENGINE_NOT_AVAILABLE"
        return [PSCustomObject]@{ Handled=$true; Status="CREATE_ENGINE_NOT_AVAILABLE" }
    }

    if ($intent -eq "CONVERT" -and $fileInfo.Exists) {
        Write-Host "Handler: GENERIC_CONVERT"
        Write-Host "STATUS: CONVERT_ENGINE_NOT_AVAILABLE"
        return [PSCustomObject]@{ Handled=$true; Status="CONVERT_ENGINE_NOT_AVAILABLE" }
    }

    if ($intent -eq "BANK_AGENT_RESEARCH_XLSX") {
        $script = "C:\AI\scripts\run_bank_agent_pipeline.ps1"
        Write-Host "Handler: BANK_AGENT_RESEARCH_XLSX"
        if ($DryRun) {
            Write-Host "STATUS: V4_BANK_AGENT_MATCHED_DRYRUN"
            return [PSCustomObject]@{ Handled=$true; Status="V4_BANK_AGENT_MATCHED_DRYRUN" }
        }
        if (!(Test-Path -LiteralPath $script -PathType Leaf)) { throw "Bank Agent pipeline not found: $script" }
        if ($XHigh) { & $script -UserRequest $Prompt -XHigh } else { & $script -UserRequest $Prompt }
        if (-not $?) { throw "V4 BANK AGENT pipeline failed." }
        return [PSCustomObject]@{ Handled=$true; Status="BANK_AGENT_PIPELINE_COMPLETE" }
    }

    Write-Host "STATUS: V4_FALLBACK_TO_LEGACY"
    return [PSCustomObject]@{ Handled=$false; Status="FALLBACK" }
}
