param(
    [string]$InputJson = "C:\AI\work\temp\bank_agent_research_final_20261007.json",
    [string]$OutputJson = "C:\AI\work\temp\bank_agent_excel_normalized_20261007.json"
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle Research Normalizer"
Write-Host "========================================"
Write-Host ""

# ------------------------------------------------------------
# 1. 입력 파일 확인
# ------------------------------------------------------------

if (!(Test-Path $InputJson)) {
    throw "입력 JSON을 찾을 수 없습니다: $InputJson"
}

$sourceJson = Get-Content `
    -LiteralPath $InputJson `
    -Raw `
    -Encoding UTF8

# JSON 자체가 정상인지 먼저 확인
try {
    $sourceData = $sourceJson | ConvertFrom-Json
}
catch {
    throw "입력 JSON 형식이 올바르지 않습니다."
}

Write-Host "Input Subjects :" $sourceData.subjects.Count


# ------------------------------------------------------------
# 2. Claude 정규화 Prompt
# ------------------------------------------------------------

$prompt = @"
너는 iM뱅크 AX 업무 조사결과를 Excel 입력 형태로 정규화하는 편집자다.

중요:
- 새로운 웹검색을 하지 마라.
- 로컬 파일을 직접 읽지 마라.
- 아래에 제공된 JSON 내용만 사용한다.
- 제공되지 않은 사실, 수치, 일정, 구축주체, 구축방식을 추정하지 마라.
- 확인되지 않은 값은 null 또는 빈 문자열로 둔다.
- 출력은 JSON만 반환한다.
- Markdown code fence를 사용하지 마라.
- 설명문이나 서론을 붙이지 마라.

==================================================
[작업 목적]
==================================================

현재 조사결과는 개별 업무 사례 단위로 너무 잘게 나뉘어 있다.

이를 기존 은행별 AI Agent 조사 Excel의 입력 구조에 맞게 정규화한다.

은행별로 하나의 업무 사례마다 행을 만드는 것이 아니라,
해당 은행이 실제 사용하는 Agent 분류 또는 합리적으로 묶을 수 있는 상위 분류 단위로 행을 만든다.

예:

잘못된 형태:
- 자산관리(PB)
- 기업금융(RM)
- 금융상담
- 여신심사

이 네 개를 반드시 4개의 Agent로 계산해서는 안 된다.

공식자료에서 모두 동일한 Agent 체계의 업무영역으로 설명된다면
하나의 상위 Agent 분류 행 안의 대상업무로 묶는다.

은행 자체의 공식 Agent 분류가 확인되지 않으면
agent_category는 "전체"로 처리한다.

==================================================
[대상 은행]
==================================================

아래 7개 은행만 포함한다.

1. 신한은행
2. 하나은행
3. 우리은행
4. 국민은행
5. 카카오뱅크
6. K뱅크
7. 부산은행

은행명은 반드시 위 표기를 사용한다.

다음 이름은 변경한다.

- KB국민은행 → 국민은행
- 케이뱅크 → K뱅크
- BNK부산은행 → 부산은행

==================================================
[Agent 판단 원칙]
==================================================

1. 단순 AI 기능을 Agent 개수로 과대계상하지 않는다.

2. 단순 검색, OCR, 번역, 요약, 챗봇이라는 이유만으로
   독립 Agent로 집계하지 않는다.

3. 원본 조사자료에서 AI Agent / 에이전트라고 명시되었거나
   Agentic Workflow가 명확한 경우만 Agent 관련 항목으로 사용한다.

4. 구축 완료와 구축 계획을 반드시 구분한다.

5. 그룹 전체 수치를 개별 은행 수치로 변환하지 않는다.

6. "100개 구축 예정"과 같이 명확한 숫자가 있는 경우에만
   해당 목표 필드에 숫자를 넣는다.

7. 완료 수치가 확인되지 않았으면 completed_count는 null이다.
   null은 0을 의미하지 않는다.

==================================================
[신한은행 기존 사용자 제공값 보호]
==================================================

신한은행은 기존 Excel에 사용자가 직접 입력한 아래 수치를 보호한다.

행 1
- agent_category: "버티컬 Agent"
- completed_count: 16

행 2
- agent_category: "일반 Agent"
- first_half_target: 80
- second_half_target: 150
- annual_target: 200

이 값은 외부 조사결과와 다르더라도 절대 변경하지 않는다.

외부 조사결과가 다른 경우:
- protected_user_value = true
- notes에 "[교차검증 필요]"라고 표시
- 사용자 기존값과 공개자료 내용을 모두 적는다.

신한은행에 대한 외부 조사내용은 위 두 분류의
target_work, source_detail, notes를 보완하는 용도로 사용한다.

새로운 수치로 기존값을 덮어쓰지 않는다.

==================================================
[Excel 입력용 출력 구조]
==================================================

반드시 아래 구조로 출력한다.

{
  "as_of_date": "2026-10-07",
  "banks": [
    {
      "bank": "신한은행",
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

==================================================
[필드 작성 규칙]
==================================================

bank
- 지정된 7개 은행 표기만 사용

agent_category
- 해당 은행의 공식 Agent 구분 우선
- 공식 분류가 확인되지 않으면 "전체"
- 개별 업무명을 무조건 Agent 분류로 사용하지 않는다

target_work
- 동일 Agent 분류 안의 주요 대상업무를 간결하게 묶어서 작성
- 예: "PB·RM·금융상담·여신심사"

builder
- 누가 구축하는지 원본에서 확인되는 경우만 작성
- 예: 현업, IT, 공동 구축, 외부사업자 등
- 확인되지 않으면 ""

build_method
- Agent Builder, 자체개발, 공동 플랫폼 등
  원본에서 확인되는 내용만 작성
- 추정 금지

completed_count
- 실제 구축 완료 수량만 숫자로 입력

first_half_target
- 상반기 목표 수량

second_half_target
- 하반기 목표 수량

annual_target
- 연간 목표 수량

future_plan
- 향후 구축 계획 요약

source_type
- 주요 출처 유형을 요약
- 복수면 "공식자료 + 주요언론"처럼 작성 가능

source_detail
- 출처명과 제목을 간결하게 연결하여 작성

source_urls
- 실제 원본 URL 배열
- 중복 URL 제거

published_dates
- 출처 발표일 배열
- 중복 날짜 제거

notes
- 수치 해석 주의
- 그룹 수치 여부
- 계획/완료 구분
- 교차검증 필요사항
- Agent 포함 여부 판단 근거
등을 기록

protected_user_value
- 사용자 기존 입력값을 보호해야 하는 행만 true

==================================================
[정규화 품질 기준]
==================================================

- 7개 은행이 정확히 모두 있어야 한다.
- 조사대상 외 은행은 포함하지 않는다.
- 업무 사례 수와 Agent 개수를 동일시하지 않는다.
- 동일한 상위 Agent 아래 업무는 target_work에 묶는다.
- 공식 분류를 확인할 수 없는 경우 억지로 분류명을 만들지 않는다.
- 원본 출처는 가능한 한 유지한다.
- 하나은행처럼 공개된 Agent 수치가 부족한 경우
  한 행 "전체"로 두는 것이 여러 가상 Agent를 만드는 것보다 낫다.
- 부산은행의 그룹 공동 플랫폼 관련 숫자는 부산은행 자체 구축 수치로 변환하지 않는다.
- 카카오뱅크의 AI 검색 등은 Agent 여부가 불명확하면
  notes에 판단 주의를 남기고 독립 Agent 수량으로 계산하지 않는다.

==================================================
[원본 조사 JSON]
==================================================

$sourceJson

==================================================
[최종 지시]
==================================================

위 JSON만 근거로 Excel 입력용 데이터를 정규화하라.

JSON 이외의 텍스트는 절대 출력하지 마라.
"@

# ------------------------------------------------------------
# 3. Claude 실행
# ------------------------------------------------------------

Write-Host "Claude normalization started..."
Write-Host ""

# 긴 Prompt는 stdin으로 전달
$claudeOutput = $prompt | & claude -p

if ([string]::IsNullOrWhiteSpace($claudeOutput)) {
    throw "Claude가 결과를 반환하지 않았습니다."
}

# 배열로 반환되는 경우 문자열로 결합
if ($claudeOutput -is [array]) {
    $claudeOutput = $claudeOutput -join "`n"
}

$claudeOutput = [string]$claudeOutput


# ------------------------------------------------------------
# 4. 혹시 code fence가 있으면 제거
# ------------------------------------------------------------

$clean = $claudeOutput.Trim()

$clean = $clean -replace '^```json\s*', ''
$clean = $clean -replace '^```\s*', ''
$clean = $clean -replace '\s*```$', ''

# JSON 앞뒤에 설명이 붙은 경우 JSON 객체만 추출
$firstBrace = $clean.IndexOf("{")
$lastBrace  = $clean.LastIndexOf("}")

if ($firstBrace -lt 0 -or $lastBrace -le $firstBrace) {
    throw "Claude 결과에서 JSON 객체를 찾을 수 없습니다."
}

$clean = $clean.Substring(
    $firstBrace,
    $lastBrace - $firstBrace + 1
)


# ------------------------------------------------------------
# 5. JSON 파싱 검증
# ------------------------------------------------------------

try {
    $normalized = $clean | ConvertFrom-Json
}
catch {
    $debugPath = "C:\AI\work\temp\normalize_claude_raw.txt"

    [System.IO.File]::WriteAllText(
        $debugPath,
        $claudeOutput,
        (New-Object System.Text.UTF8Encoding($false))
    )

    throw "Claude 결과 JSON 파싱 실패. 원본 저장: $debugPath"
}


# ------------------------------------------------------------
# 6. 은행 목록 검증
# ------------------------------------------------------------

$expectedBanks = @(
    "신한은행",
    "하나은행",
    "우리은행",
    "국민은행",
    "카카오뱅크",
    "K뱅크",
    "부산은행"
)

$actualBanks = @(
    $normalized.banks |
    ForEach-Object { $_.bank }
)

$missing = @(
    $expectedBanks |
    Where-Object { $_ -notin $actualBanks }
)

$extra = @(
    $actualBanks |
    Where-Object { $_ -notin $expectedBanks }
)

if ($normalized.banks.Count -ne 7) {
    throw "정규화 결과 은행 수가 7개가 아닙니다: $($normalized.banks.Count)"
}

if ($missing.Count -gt 0) {
    throw "정규화 결과 누락 은행: $($missing -join ', ')"
}

if ($extra.Count -gt 0) {
    throw "정규화 결과 불필요 은행: $($extra -join ', ')"
}


# ------------------------------------------------------------
# 7. 신한은행 보호값 검증
# ------------------------------------------------------------

$shinhan = $normalized.banks |
    Where-Object { $_.bank -eq "신한은행" } |
    Select-Object -First 1

if ($null -eq $shinhan) {
    throw "신한은행 결과가 없습니다."
}

$vertical = $shinhan.rows |
    Where-Object { $_.agent_category -eq "버티컬 Agent" } |
    Select-Object -First 1

$general = $shinhan.rows |
    Where-Object { $_.agent_category -eq "일반 Agent" } |
    Select-Object -First 1

if ($null -eq $vertical) {
    throw "신한은행 버티컬 Agent 행이 없습니다."
}

if ($null -eq $general) {
    throw "신한은행 일반 Agent 행이 없습니다."
}

if ([int]$vertical.completed_count -ne 16) {
    throw "신한 버티컬 Agent 보호값 오류: 구축완료 16이 아닙니다."
}

if ([int]$general.first_half_target -ne 80) {
    throw "신한 일반 Agent 보호값 오류: 상반기 80이 아닙니다."
}

if ([int]$general.second_half_target -ne 150) {
    throw "신한 일반 Agent 보호값 오류: 하반기 150이 아닙니다."
}

if ([int]$general.annual_target -ne 200) {
    throw "신한 일반 Agent 보호값 오류: 연간 200이 아닙니다."
}


# ------------------------------------------------------------
# 8. 출력 저장 - UTF-8 no BOM
# ------------------------------------------------------------

$normalizedJson = $normalized |
    ConvertTo-Json -Depth 30

[System.IO.File]::WriteAllText(
    $OutputJson,
    $normalizedJson,
    (New-Object System.Text.UTF8Encoding($false))
)


# ------------------------------------------------------------
# 9. 결과 요약
# ------------------------------------------------------------

Write-Host ""
Write-Host "=== NORMALIZATION RESULT ==="

$totalRows = 0

foreach ($bank in $normalized.banks) {

    $rowCount = @($bank.rows).Count
    $totalRows += $rowCount

    Write-Host ""
    Write-Host "[$($bank.bank)] Rows:" $rowCount

    foreach ($row in $bank.rows) {

        Write-Host (
            " - {0} | 완료:{1} | 상:{2} | 하:{3} | 연간:{4}" -f `
            $row.agent_category,
            $(if ($null -eq $row.completed_count) {"-"} else {$row.completed_count}),
            $(if ($null -eq $row.first_half_target) {"-"} else {$row.first_half_target}),
            $(if ($null -eq $row.second_half_target) {"-"} else {$row.second_half_target}),
            $(if ($null -eq $row.annual_target) {"-"} else {$row.annual_target})
        )
    }
}

Write-Host ""
Write-Host "Banks     :" $normalized.banks.Count
Write-Host "Total Rows:" $totalRows
Write-Host "Output    :" $OutputJson
Write-Host ""
Write-Host "STATUS: NORMALIZED_JSON_READY"
