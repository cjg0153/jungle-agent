function Invoke-JungleCoreV5Phase2 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [bool]$DryRun = $false
    )

    $filePattern = '(?i)([A-Z]:\\.+?\.(?:pptx?|docx?|xlsx?|xlsm|xls|pdf|html?|csv|txt|md))'
    $m = [regex]::Match($Prompt,$filePattern)

    if (-not $m.Success) {
        return [PSCustomObject]@{ Handled=$false; Status="V5_NO_FILE" }
    }

    $path = $m.Groups[1].Value.Trim()
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) {
        return [PSCustomObject]@{ Handled=$false; Status="V5_FILE_NOT_FOUND" }
    }

    $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

    $isModify =
        ($Prompt -match "수정해줘") -or
        ($Prompt -match "수정해 줘") -or
        ($Prompt -match "고쳐줘") -or
        ($Prompt -match "바꿔줘") -or
        ($Prompt -match "반영해줘") -or
        ($Prompt -match "다듬어줘") -or
        ($Prompt -match "추가해줘") -or
        ($Prompt -match "삭제해줘") -or
        ($Prompt -match "변경해줘")

    $isReview =
        ($Prompt -match "봐줘") -or
        ($Prompt -match "검토") -or
        ($Prompt -match "리뷰") -or
        ($Prompt -match "점검") -or
        ($Prompt -match "분석해줘") -or
        ($Prompt -match "확인해줘")

    $hasCreateVerb =
        ($Prompt -match "새.*만들어줘") -or
        ($Prompt -match "새로.*만들어줘") -or
        ($Prompt -match "작성해줘") -or
        ($Prompt -match "생성해줘") -or
        ($Prompt -match "문서.*만들어줘") -or
        ($Prompt -match "보고서.*만들어줘")

    $wantsDocx =
        ($Prompt -match "(?i)\bWord\b") -or
        ($Prompt -match "워드") -or
        ($Prompt -match "(?i)\bDOCX\b") -or
        ($Prompt -match "Word 문서")

    if ($isModify -and $ext -in @(".pptx",".ppt")) {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : PPTX_MODIFY_QUALITY"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_PPTX_MODIFY_QUALITY_DRYRUN" } }

        & "C:\AI\scripts\run_pptx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Prompt
        if (-not $?) { throw "PPTX MODIFY QUALITY 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="PPTX_MODIFY_QUALITY_COMPLETE" }
    }

    if ($isModify -and $ext -eq ".docx") {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : DOCX_MODIFY_QUALITY"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_DOCX_MODIFY_QUALITY_DRYRUN" } }

        & "C:\AI\scripts\run_docx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Prompt
        if (-not $?) { throw "DOCX MODIFY QUALITY 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="DOCX_MODIFY_QUALITY_COMPLETE" }
    }

    if ($isModify -and $ext -in @(".xlsx",".xlsm",".xls")) {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : XLSX_MODIFY_QUALITY"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_XLSX_MODIFY_QUALITY_DRYRUN" } }

        & "C:\AI\scripts\run_xlsx_modify_quality_v1.ps1" -FilePath $path -UserRequest $Prompt
        if (-not $?) { throw "XLSX MODIFY QUALITY 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="XLSX_MODIFY_QUALITY_COMPLETE" }
    }

    if ($isReview -and $ext -eq ".docx") {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : DOCX_QA_REVIEW"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_DOCX_QA_REVIEW_DRYRUN" } }

        & "C:\AI\scripts\run_office_codex_qa_v1.ps1" -FilePath $path -UserRequest $Prompt
        if (-not $?) { throw "DOCX QA REVIEW 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="DOCX_QA_REVIEW_COMPLETE" }
    }

    if ($isReview -and $ext -in @(".xlsx",".xlsm",".xls")) {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : XLSX_QA_REVIEW"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_XLSX_QA_REVIEW_DRYRUN" } }

        & "C:\AI\scripts\run_office_codex_qa_v1.ps1" -FilePath $path -UserRequest $Prompt
        if (-not $?) { throw "XLSX QA REVIEW 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="XLSX_QA_REVIEW_COMPLETE" }
    }

    if ($hasCreateVerb -and $wantsDocx) {
        Write-Host ""
        Write-Host "=== JUNGLE CORE v5 PHASE2 ==="
        Write-Host "Handler : DOCX_CREATE_QUALITY"
        if ($DryRun) { return [PSCustomObject]@{ Handled=$true; Status="V5_DOCX_CREATE_QUALITY_DRYRUN" } }

        & "C:\AI\scripts\run_docx_create_quality_v1.ps1" -SourcePath $path -UserRequest $Prompt
        if (-not $?) { throw "DOCX CREATE QUALITY 실패" }
        return [PSCustomObject]@{ Handled=$true; Status="DOCX_CREATE_QUALITY_COMPLETE" }
    }

    return [PSCustomObject]@{ Handled=$false; Status="V5_NOT_HANDLED" }
}
