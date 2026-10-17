param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$true)]
    [string]$UserRequest
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()

if ($ext -ne ".docx" -and $ext -ne ".doc") {
    throw "DOC/DOCX 파일만 지원합니다."
}

$outputDir = "C:\AI\work\output"
$tempDir   = "C:\AI\work\temp"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath = Join-Path $outputDir "${baseName}_modified_${dateStamp}${ext}"
$extractPath = Join-Path $tempDir "docx_extract_${dateStamp}.json"
$planPath = Join-Path $tempDir "docx_edit_plan_${dateStamp}.json"
$instructionPath = Join-Path $tempDir "docx_instruction_${dateStamp}.txt"

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle DOCX Modify"
Write-Host "========================================"
Write-Host ""
Write-Host "ORIGINAL :" $FilePath
Write-Host "OUTPUT   :" $outputPath
Write-Host ""

# ------------------------------------------------------------
# 1. 원본 보존용 복사
# ------------------------------------------------------------

Copy-Item `
    -LiteralPath $FilePath `
    -Destination $outputPath `
    -Force


# ------------------------------------------------------------
# 2. Word 문서 내용 추출
# ------------------------------------------------------------

& "C:\AI\scripts\extract_office_content.ps1" `
    -FilePath $FilePath `
    -OutputPath $extractPath

if (!(Test-Path -LiteralPath $extractPath -PathType Leaf)) {
    throw "Word 내용 추출에 실패했습니다."
}


# ------------------------------------------------------------
# 3. Claude 수정계획 작성
# 긴 한국어 Prompt는 파일로 저장해서 전달
# ------------------------------------------------------------

$prompt = @"
다음 Word 문서의 추출 내용을 읽고
사용자의 요청에 필요한 최소한의 수정계획을 작성하세요.

[Word 추출 내용]
$extractPath

[사용자 요청]
$UserRequest

[중요 원칙]

1. 사용자가 요청한 부분만 수정하세요.
2. 기존 문서의 순서, 스타일, 구조를 최대한 유지하세요.
3. 요청하지 않은 문단은 수정하지 마세요.
4. 문단을 임의로 삭제하거나 새로 만들지 마세요.
5. paragraph_number는 추출 JSON에 존재하는 실제 번호를 사용하세요.
6. old_text는 추출 JSON에 존재하는 원문과 정확히 동일하게 작성하세요.
7. new_text에는 수정 후 해당 문단 전체 텍스트를 작성하세요.
8. 새로운 사실을 추측해서 추가하지 마세요.
9. 자연스러운 한국어를 사용하세요.
10. 수정이 필요 없는 문단은 edits에 포함하지 마세요.

아래 형식의 JSON으로만 작성하세요.

{
  "edits": [
    {
      "paragraph_number": 1,
      "old_text": "기존 문구",
      "new_text": "수정 문구",
      "reason": "수정 이유"
    }
  ]
}

JSON 파일을 다음 경로에 UTF-8로 저장하세요.

$planPath

저장이 끝나면 다른 설명 없이
DOCX_EDIT_PLAN_CREATED
만 출력하세요.
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllText(
    $instructionPath,
    $prompt,
    $utf8Bom
)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified in that instruction. Do not use web search, Google Drive, connectors, or external sources. When finished, output only DOCX_EDIT_PLAN_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $planPath -PathType Leaf)) {
    throw "Claude DOCX 수정계획 JSON이 생성되지 않았습니다."
}


# ------------------------------------------------------------
# 4. 수정계획 검증
# ------------------------------------------------------------

$plan = Get-Content `
    -LiteralPath $planPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

if ($null -eq $plan.edits) {
    throw "DOCX 수정계획에 edits가 없습니다."
}

Write-Host ""
Write-Host "EDIT PLAN :" $plan.edits.Count "건"
Write-Host ""


# ------------------------------------------------------------
# 5. Word COM 수정
# ------------------------------------------------------------

$word = $null
$doc  = $null

$applied = 0
$failed = @()

try {

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open($outputPath)

    foreach ($edit in $plan.edits) {

        $paragraphNumber = [int]$edit.paragraph_number
        $oldText = ([string]$edit.old_text).Trim()
        $newText = [string]$edit.new_text

        if (
            $paragraphNumber -lt 1 -or
            $paragraphNumber -gt $doc.Paragraphs.Count
        ) {
            $failed += "Paragraph $paragraphNumber : 존재하지 않는 문단"
            continue
        }

        $paragraph = $doc.Paragraphs.Item($paragraphNumber)

        $currentText = [string]$paragraph.Range.Text
        $currentText = $currentText -replace "[`r`a]", ""
        $currentText = $currentText.Trim()

        if ($currentText -ne $oldText) {
            $failed += "Paragraph $paragraphNumber : 원문 불일치"
            continue
        }

        # 문단 끝 마커는 그대로 두고 본문만 교체
        $range = $paragraph.Range.Duplicate

        if ($range.End -gt $range.Start) {
            $range.End = $range.End - 1
        }

        $range.Text = $newText
        $applied++

        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) |
                Out-Null
        }
        catch {}

        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($paragraph) |
                Out-Null
        }
        catch {}
    }

    $doc.Save()
}
finally {

    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) |
                Out-Null
        }
        catch {}
    }

    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) |
                Out-Null
        }
        catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ------------------------------------------------------------
# 6. 결과
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "DOCX 수정본이 생성되지 않았습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE DOCX MODIFY RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Modify  : OK"
Write-Host "Applied :" $applied
Write-Host "Failed  :" $failed.Count
Write-Host "Original:" $FilePath
Write-Host "File    :" $outputPath
Write-Host "Size    :" (Get-Item $outputPath).Length

if ($failed.Count -gt 0) {

    Write-Host ""
    Write-Host "미적용 항목:"

    foreach ($item in $failed) {
        Write-Host " -" $item
    }

    Write-Host ""
    Write-Host "STATUS: DOCX_MODIFY_PARTIAL"
}
else {

    Write-Host ""
    Write-Host "STATUS: DOCX_MODIFY_COMPLETE"
}