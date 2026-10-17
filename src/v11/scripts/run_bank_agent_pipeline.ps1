param(
    [Parameter(Mandatory=$true)]
    [string]$UserRequest,

    [string]$AsOfDate = (Get-Date -Format "yyyy-MM-dd"),

    [string]$TemplatePath = "C:\AI\work\input\은행별_AI_Agent_구축현황_조사표_20261006.xlsx",

    [switch]$XHigh,

    [switch]$ReuseResearch,

    [switch]$ReuseNormalized
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$root    = "C:\AI"
$tempDir = "C:\AI\work\temp"
$outDir  = "C:\AI\work\output"

if (!(Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
if (!(Test-Path $outDir))  { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

try {
    $dateObj = [datetime]::ParseExact(
        $AsOfDate,
        "yyyy-MM-dd",
        [Globalization.CultureInfo]::InvariantCulture
    )
}
catch {
    throw "AsOfDate 형식은 yyyy-MM-dd 이어야 합니다."
}

$dateStamp = $dateObj.ToString("yyyyMMdd")
$dateDot   = $dateObj.ToString("yyyy.MM.dd")

$ResearchJson   = Join-Path $tempDir "bank_agent_research_final_$dateStamp.json"
$NormalizedJson = Join-Path $tempDir "bank_agent_excel_normalized_$dateStamp.json"
$OutputExcel    = Join-Path $outDir "은행별_AI_Agent_구축현황_조사표_${dateStamp}_완성.xlsx"

$WriterScript = "C:\AI\scripts\write_research_excel.ps1"

Write-Host ""
Write-Host "================================================"
Write-Host " Jungle Bank Agent Pipeline"
Write-Host "================================================"
Write-Host ""
Write-Host "Request :" $UserRequest
Write-Host "As Of   :" $AsOfDate
Write-Host ""

# ============================================================
# PHASE 1 : CODEX LIVE RESEARCH
# ============================================================

if ($ReuseResearch -and (Test-Path $ResearchJson)) {
    Write-Host "[1/4] CODEX RESEARCH"
    Write-Host "      Existing research JSON reused."
    Write-Host "      $ResearchJson"
    Write-Host ""
}
else {
    Write-Host "[1/4] CODEX RESEARCH"
    Write-Host "      Live web research started..."
    Write-Host ""

    if (Test-Path $ResearchJson) {
        Remove-Item $ResearchJson -Force
    }

    $researchPrompt = @"
너는 Jungle AI의 은행 AI Agent 조사 엔진이다.

사용자 요청:
$UserRequest

조사 기준일:
$AsOfDate

조사 대상은 아래 7개 은행으로 한정한다.
- 신한은행
- 하나은행
- 우리은행
- KB국민은행
- 카카오뱅크
- 케이뱅크
- BNK부산은행

NH농협은행, IBK기업은행, iM뱅크, 토스뱅크 등
위 목록에 없는 은행은 결과에 포함하지 마라.

목적:
각 은행의 AI Agent 구축현황과 공식 계획을 조사한다.

조사 원칙:
1. 최신 공개자료를 웹에서 조사한다.
2. 공식 홈페이지/보도자료/IR/공공기관/공식 사업자 자료를 우선한다.
3. 필요한 경우 주요 언론을 보조 출처로 사용한다.
4. 같은 사실을 반복 검색하지 않는다.
5. 숫자를 추정하지 않는다.
6. 구축 완료, 상반기 목표, 하반기 목표, 연간 목표, 향후 계획을 구분한다.
7. 단순 챗봇, 검색, OCR, 번역, 요약 기능을 근거 없이 독립 Agent 구축 실적으로 계산하지 않는다.
8. 개별 업무 사례 수를 Agent 구축 개수로 임의 환산하지 않는다.
9. 금융그룹 전체 수치를 특정 은행의 수치로 임의 환산하지 않는다.
10. 출처가 충돌하면 conflicting_sources와 notes에 기록한다.
11. 확인할 수 없는 숫자는 null로 둔다. null은 0을 의미하지 않는다.
12. Excel 파일은 수정하지 않는다.
13. 로컬의 다른 문서나 Skill 파일은 읽지 않는다.

반드시 아래 경로에 UTF-8 JSON 파일을 직접 생성한다:
$ResearchJson

상위 구조:
{
  "as_of_date": "$AsOfDate",
  "request": "",
  "subjects": [
    {
      "name": "",
      "records": [
        {
          "category": "",
          "description": "",
          "status": "",
          "agent_classification": "",
          "completed_count": null,
          "first_half_target": null,
          "second_half_target": null,
          "annual_target": null,
          "future_plan": "",
          "sources": [
            {
              "source_type": "",
              "publisher": "",
              "title": "",
              "published_date": "",
              "url": ""
            }
          ],
          "conflicting_sources": [],
          "notes": ""
        }
      ]
    }
  ]
}

subjects는 반드시 정확히 7개여야 하며 name은 아래 표기만 사용한다:
신한은행
하나은행
우리은행
KB국민은행
카카오뱅크
케이뱅크
BNK부산은행

JSON 생성 후 문법 유효성을 확인한다.
긴 JSON을 stdout으로 출력하지 말고 짧게 완료만 알린다.
"@

    $reasoning = if ($XHigh) { "xhigh" } else { "medium" }

    Push-Location $root
    try {
        $researchPrompt |
            codex --search exec `
                --skip-git-repo-check `
                --sandbox workspace-write `
                --config "model_reasoning_effort=$reasoning" `
                - |
            Out-Null

        $codexExit = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($codexExit -ne 0) {
        throw "Codex research failed. ExitCode: $codexExit"
    }

    if (!(Test-Path $ResearchJson)) {
        throw "Codex가 Research JSON을 생성하지 못했습니다."
    }

    Write-Host "      Research JSON created."
    Write-Host ""
}

# ============================================================
# RESEARCH VALIDATION
# ============================================================

$researchText = Get-Content -LiteralPath $ResearchJson -Raw -Encoding UTF8

try {
    $research = $researchText | ConvertFrom-Json
}
catch {
    throw "Research JSON parsing failed."
}

$expectedResearchBanks = @(
    "신한은행",
    "하나은행",
    "우리은행",
    "KB국민은행",
    "카카오뱅크",
    "케이뱅크",
    "BNK부산은행"
)

$actualResearchBanks = @($research.subjects | ForEach-Object { $_.name })
$missingResearch = @($expectedResearchBanks | Where-Object { $_ -notin $actualResearchBanks })
$extraResearch   = @($actualResearchBanks | Where-Object { $_ -notin $expectedResearchBanks })

if (@($research.subjects).Count -ne 7) {
    throw "Research 대상이 7개 은행이 아닙니다."
}
if ($missingResearch.Count -gt 0) {
    throw "Research 누락 은행: $($missingResearch -join ', ')"
}
if ($extraResearch.Count -gt 0) {
    throw "Research 불필요 은행: $($extraResearch -join ', ')"
}

Write-Host "      Research Validation: OK"
Write-Host ""

# ============================================================
# PHASE 2 : CLAUDE NORMALIZATION
# ============================================================

if ($ReuseNormalized -and (Test-Path $NormalizedJson)) {
    Write-Host "[2/4] CLAUDE NORMALIZATION"
    Write-Host "      Existing normalized JSON reused."
    Write-Host ""
}
else {
    Write-Host "[2/4] CLAUDE NORMALIZATION"
    Write-Host "      Excel-oriented normalization started..."
    Write-Host ""

    if (Test-Path $NormalizedJson) {
        Remove-Item $NormalizedJson -Force
    }

    $normalizationPrompt = @"
너는 iM뱅크 AX 업무용 조사결과 정규화 편집자다.

새로운 웹검색을 하지 마라.
로컬의 다른 파일을 읽지 마라.
아래 Research JSON만 사용한다.

최종 결과를 stdout에 JSON으로 출력하지 말고
반드시 Write 도구를 사용하여 아래 경로에
UTF-8 순수 JSON 파일로 직접 저장한다:

$NormalizedJson

조사 기준일:
$AsOfDate

목적:
개별 업무 사례 단위 조사결과를
기존 은행별 AI Agent 조사 Excel에 넣기 좋은
은행별 상위 Agent 분류 단위로 정규화한다.

업무 사례 수와 Agent 수를 동일시하지 않는다.
은행 자체의 공식 Agent 분류가 확인되지 않으면
agent_category는 "전체"로 작성한다.

은행 표기:
신한은행
하나은행
우리은행
국민은행
카카오뱅크
K뱅크
부산은행

변환:
KB국민은행 -> 국민은행
케이뱅크 -> K뱅크
BNK부산은행 -> 부산은행

신한은행 사용자 기존값 보호:
- 버티컬 Agent: completed_count = 16
- 일반 Agent: first_half_target = 80
- 일반 Agent: second_half_target = 150
- 일반 Agent: annual_target = 200
- 두 행 모두 protected_user_value = true
공개자료와 차이가 있으면 수치를 바꾸지 말고 notes에 "[교차검증 필요]"라고 기록한다.

판단 규칙:
- 단순 검색/OCR/번역/요약/챗봇을 독립 Agent 개수로 과대계상하지 않는다.
- 개별 업무 사례를 자동으로 각각 Agent 1개로 계산하지 않는다.
- 금융그룹 전체 수치를 개별 은행 수치로 변환하지 않는다.
- 완료 수치가 확인되지 않았으면 completed_count = null.
- 목표 수치가 명확히 확인된 경우만 숫자를 기록한다.
- 확인되지 않은 값은 빈 문자열 또는 null.
- 임의 추정 금지.

특별 주의:
- 국민은행: PB, RM, 금융상담, 여신심사를 각각 별도 Agent 4개로 계산하지 않는다.
- 우리은행: CRM·기업여신·자산관리·내부통제·고객상담·업무자동화 등을 각각 독립 Agent 수량으로 자동 계산하지 않는다.
- 카카오뱅크: AI 이체·AI 검색·AI 모임총무가 독립 AI Agent라는 근거가 불충분하면 Agent 구축 수량으로 계산하지 않는다.
- K뱅크: 개발지원 Agentic AI·이해상충거래 셀프체크·금융사기 메일 자동분류 사례를 자동으로 3개 Agent로 계산하지 않는다.
- 부산은행: BNK금융그룹 공동 플랫폼의 Agent 수를 부산은행 자체 구축 수량으로 환산하지 않는다.

출력 구조:
{
  "as_of_date": "$AsOfDate",
  "banks": [
    {
      "bank": "",
      "rows": [
        {
          "agent_category": "",
          "target_work": "",
          "builder": "",
          "build_method": "",
          "completed_count": null,
          "first_half_target": null,
          "second_half_target": null,
          "annual_target": null,
          "future_plan": "",
          "source_type": "",
          "source_detail": "",
          "source_urls": [],
          "published_dates": [],
          "notes": "",
          "protected_user_value": false
        }
      ]
    }
  ]
}

Research JSON:
$researchText

저장 전 반드시 확인:
- 정확히 7개 은행
- 신한 버티컬 completed_count = 16
- 신한 일반 first_half_target = 80
- 신한 일반 second_half_target = 150
- 신한 일반 annual_target = 200
- 그룹 수치를 개별은행 수치로 환산하지 않음
- 확인되지 않은 완료값을 0으로 만들지 않음
- source URL 최대한 보존
- 한국어 UTF-8 정상

파일 생성 후 JSON 문법도 확인한다.
채팅 응답은 짧게 끝내라.
"@

    Push-Location $root
    try {
        $normalizationPrompt |
            claude -p --allowedTools "Write" |
            Out-Null

        $claudeExit = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($claudeExit -ne 0) {
        throw "Claude normalization failed. ExitCode: $claudeExit"
    }

    if (!(Test-Path $NormalizedJson)) {
        throw "Claude가 Normalized JSON을 생성하지 못했습니다."
    }
}

# ============================================================
# NORMALIZED VALIDATION
# ============================================================

$normalizedText = Get-Content -LiteralPath $NormalizedJson -Raw -Encoding UTF8

try {
    $normalized = $normalizedText | ConvertFrom-Json
}
catch {
    throw "Normalized JSON parsing failed."
}

$expectedExcelBanks = @(
    "신한은행",
    "하나은행",
    "우리은행",
    "국민은행",
    "카카오뱅크",
    "K뱅크",
    "부산은행"
)

$actualExcelBanks = @($normalized.banks | ForEach-Object { $_.bank })
$missingExcel = @($expectedExcelBanks | Where-Object { $_ -notin $actualExcelBanks })
$extraExcel   = @($actualExcelBanks | Where-Object { $_ -notin $expectedExcelBanks })

if (@($normalized.banks).Count -ne 7) {
    throw "Normalized 결과 은행 수가 7개가 아닙니다."
}
if ($missingExcel.Count -gt 0) {
    throw "Normalized 누락 은행: $($missingExcel -join ', ')"
}
if ($extraExcel.Count -gt 0) {
    throw "Normalized 불필요 은행: $($extraExcel -join ', ')"
}

$shinhan = $normalized.banks | Where-Object { $_.bank -eq "신한은행" } | Select-Object -First 1
$vertical = $shinhan.rows | Where-Object { $_.agent_category -eq "버티컬 Agent" } | Select-Object -First 1
$general  = $shinhan.rows | Where-Object { $_.agent_category -eq "일반 Agent" } | Select-Object -First 1

if ($null -eq $vertical) { throw "신한 버티컬 Agent 행 없음" }
if ($null -eq $general)  { throw "신한 일반 Agent 행 없음" }

if ([double]$vertical.completed_count -ne 16) { throw "신한 버티컬 보호값 16 오류" }
if ([double]$general.first_half_target -ne 80) { throw "신한 일반 상반기 보호값 80 오류" }
if ([double]$general.second_half_target -ne 150) { throw "신한 일반 하반기 보호값 150 오류" }
if ([double]$general.annual_target -ne 200) { throw "신한 일반 연간 보호값 200 오류" }

$totalRows = 0
foreach ($b in $normalized.banks) {
    $totalRows += @($b.rows).Count
}

Write-Host "      Normalized Banks:" @($normalized.banks).Count
Write-Host "      Normalized Rows :" $totalRows
Write-Host "      Protected Data  : OK"
Write-Host ""

# ============================================================
# PHASE 3 : EXCEL COM
# ============================================================

Write-Host "[3/4] EXCEL COM"
Write-Host "      Workbook generation started..."
Write-Host ""

if (!(Test-Path $WriterScript)) {
    throw "Excel writer script not found: $WriterScript"
}
if (!(Test-Path $TemplatePath)) {
    throw "Excel template not found: $TemplatePath"
}

& powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File $WriterScript `
    -TemplatePath $TemplatePath `
    -JsonPath $NormalizedJson `
    -OutputPath $OutputExcel

if ($LASTEXITCODE -ne 0) {
    throw "Excel writer failed."
}
if (!(Test-Path $OutputExcel)) {
    throw "Excel output was not created."
}

# Dynamic date patch
$excel = $null
$wb = $null
$ws = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($OutputExcel)
    $ws = $wb.Worksheets.Item("조사표")

    $ws.Range("A1").Value2 = "은행별 AI Agent 구축 현황 조사표 (기준일: $dateDot)"

    foreach ($r in 6..19) {
        $category = "$($ws.Cells.Item($r,3).Value2)"
        if (![string]::IsNullOrWhiteSpace($category)) {
            $ws.Cells.Item($r,14).Value2 = $dateDot
        }
    }

    $wb.Save()
}
finally {
    if ($null -ne $wb) { $wb.Close($true) }
    if ($null -ne $excel) { $excel.Quit() }

    if ($null -ne $ws) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) }
    if ($null -ne $wb) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "      Excel created:"
Write-Host "      $OutputExcel"
Write-Host ""

# ============================================================
# PHASE 4 : FINAL QA
# ============================================================

Write-Host "[4/4] FINAL QA"
Write-Host ""

$excel = $null
$wb = $null
$ws = $null
$summary = $null
$errors = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($OutputExcel, 0, $true)
    $ws = $wb.Worksheets.Item("조사표")
    $summary = $wb.Worksheets.Item("은행별 요약")

    $excel.CalculateFull()

    if ([double]$ws.Range("G6").Value2 -ne 16) { $errors += "신한 버티컬 완료 16 손상" }
    if ([double]$ws.Range("H7").Value2 -ne 80) { $errors += "신한 일반 상반기 80 손상" }
    if ([double]$ws.Range("I7").Value2 -ne 150) { $errors += "신한 일반 하반기 150 손상" }
    if ([double]$ws.Range("J7").Value2 -ne 200) { $errors += "신한 일반 연간 200 손상" }

    foreach ($r in 6..19) {
        if (!$ws.Cells.Item($r,11).HasFormula) {
            $errors += "K$r 자동계산 수식 없음"
        }
    }

    foreach ($r in 5..11) {
        foreach ($c in 2..7) {
            if (!$summary.Cells.Item($r,$c).HasFormula) {
                $addr = $summary.Cells.Item($r,$c).Address($false,$false)
                $errors += "은행별 요약 $addr 수식 없음"
            }
        }
    }

    $templateBanks = @(
        "신한은행",
        "우리은행",
        "국민은행",
        "하나은행",
        "부산은행",
        "카카오뱅크",
        "K뱅크"
    )

    foreach ($bankName in $templateBanks) {
        $found = $false
        foreach ($r in 6..19) {
            if ("$($ws.Cells.Item($r,2).Value2)" -eq $bankName) {
                $found = $true
                break
            }
        }
        if (!$found) {
            $errors += "조사표 은행 누락: $bankName"
        }
    }

    if ("$($ws.Range("A1").Value2)" -notlike "*$dateDot*") {
        $errors += "기준일 제목 오류"
    }
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }

    if ($null -ne $ws) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) }
    if ($null -ne $summary) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($summary) }
    if ($null -ne $wb) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "================================================"
Write-Host " JUNGLE PIPELINE RESULT"
Write-Host "================================================"
Write-Host ""

if ($errors.Count -gt 0) {
    Write-Host "STATUS: BANK_AGENT_PIPELINE_QA_FAIL"
    foreach ($e in $errors) {
        Write-Host " -" $e
    }
    throw "Final Excel QA failed."
}

$resultFile = Get-Item $OutputExcel

Write-Host "Research   : OK"
Write-Host "Normalize  : OK"
Write-Host "Excel      : OK"
Write-Host "QA         : PASS"
Write-Host ""
Write-Host "File       :" $resultFile.FullName
Write-Host "Size       :" $resultFile.Length
Write-Host ""
Write-Host "STATUS: BANK_AGENT_PIPELINE_COMPLETE"
