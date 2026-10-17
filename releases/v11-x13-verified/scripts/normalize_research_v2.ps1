param(
    [string]$InputJson  = "C:\AI\work\temp\bank_agent_research_final_20261007.json",
    [string]$OutputJson = "C:\AI\work\temp\bank_agent_excel_normalized_20261007.json"
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle Research Normalizer v2"
Write-Host " Direct File Write Mode"
Write-Host "========================================"
Write-Host ""

# ------------------------------------------------------------
# 1. 입력 JSON 확인
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $InputJson)) {
    throw "입력 JSON을 찾을 수 없습니다: $InputJson"
}

$sourceJson = Get-Content `
    -LiteralPath $InputJson `
    -Raw `
    -Encoding UTF8

try {
    $sourceData = $sourceJson | ConvertFrom-Json
}
catch {
    throw "입력 JSON 형식이 올바르지 않습니다."
}

if ($sourceData.subjects.Count -ne 7) {
    throw "입력 조사대상이 7개 은행이 아닙니다: $($sourceData.subjects.Count)"
}

Write-Host "Input Subjects :" $sourceData.subjects.Count
Write-Host "Input File     :" $InputJson
Write-Host ""

# ------------------------------------------------------------
# 2. 이전 출력 제거
# ------------------------------------------------------------

if (Test-Path -LiteralPath $OutputJson) {
    Remove-Item -LiteralPath $OutputJson -Force
}

# ------------------------------------------------------------
# 3. Claude에게 전달할 Prompt
#    중요: 결과를 stdout으로 출력하지 않고 파일에 직접 저장
# ------------------------------------------------------------

$prompt = @"
너는 iM뱅크 AX 업무용 조사결과 정규화 편집자다.

이번 작업에서 가장 중요한 규칙:

1. 새로운 웹검색을 하지 마라.
2. 로컬의 다른 파일을 읽지 마라.
3. 아래 프롬프트에 포함된 원본 조사 JSON만 근거로 사용한다.
4. 사실이나 숫자를 추정하지 마라.
5. 최종 결과를 채팅 응답으로 출력하지 마라.
6. 반드시 Write 도구를 사용하여 아래 경로에 JSON 파일을 직접 생성하라.

저장 경로:
$OutputJson

파일은 반드시 순수 JSON이어야 한다.
Markdown code fence를 넣지 마라.
JSON 앞뒤에 설명문을 넣지 마라.
UTF-8 한글을 그대로 유지하라.

==================================================
작업 목적
==================================================

현재 조사결과는 개별 업무 사례 단위로 세분화되어 있다.

이를 은행별 AI Agent 구축현황 Excel에 입력할 수 있도록
은행별 상위 Agent 분류 기준으로 정규화한다.

개별 업무사례 수와 Agent 개수를 동일시하지 않는다.

예:

자산관리(PB)
기업금융(RM)
금융상담
여신심사

위 업무들이 동일한 Agent 체계의 적용업무로 설명된다면
각각 별도 Agent로 만들지 않고 하나의 상위 행으로 통합하고
target_work에 업무를 묶어서 기록한다.

은행 자체의 공식 Agent 분류를 확인할 수 없으면

"agent_category": "전체"

로 작성한다.

==================================================
대상 은행
==================================================

반드시 아래 7개 은행만 포함한다.

신한은행
하나은행
우리은행
국민은행
카카오뱅크
K뱅크
부산은행

원본 은행명은 다음과 같이 변경한다.

KB국민은행 -> 국민은행
케이뱅크 -> K뱅크
BNK부산은행 -> 부산은행

==================================================
Agent 판단 원칙
==================================================

- 단순 AI 기능을 독립 Agent 개수로 과대계상하지 않는다.
- 단순 검색, OCR, 번역, 요약, 챗봇이라는 이유만으로 Agent로 계산하지 않는다.
- AI Agent, 에이전트, Agentic Workflow 근거가 있는 경우를 우선한다.
- 구축 완료와 향후 계획을 구분한다.
- 그룹 전체 수치를 개별 은행 수치로 변환하지 않는다.
- 목표 숫자가 명시된 경우에만 숫자를 기록한다.
- 완료 수치가 확인되지 않은 경우 completed_count는 null이다.
- null은 0을 의미하지 않는다.
- 공식 분류가 없으면 임의의 Agent 명칭을 만들지 말고 "전체"로 한다.

==================================================
신한은행 사용자 기존값 보호
==================================================

신한은행 기존 Excel에는 사용자가 직접 입력한 값이 있다.

이 값은 절대 변경하지 않는다.

첫 번째 행:

agent_category = "버티컬 Agent"
completed_count = 16
protected_user_value = true

두 번째 행:

agent_category = "일반 Agent"
first_half_target = 80
second_half_target = 150
annual_target = 200
protected_user_value = true

공개 조사자료와 값이 다르더라도 위 숫자를 변경하지 않는다.

차이가 발견되면 notes에:

[교차검증 필요]

라고 표시하고 사용자 입력값과 공개자료의 차이를 설명한다.

공개 조사자료는 target_work, future_plan, source 정보,
notes를 보완하는 용도로 사용한다.

==================================================
출력 JSON 구조
==================================================

반드시 아래 구조를 따른다.

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
필드 작성 규칙
==================================================

agent_category
- 은행의 공식 Agent 분류를 우선한다.
- 공식 분류가 확인되지 않으면 "전체".
- 단순 업무명을 Agent 분류명으로 만들지 않는다.

target_work
- 동일 Agent 분류에 속하는 업무를 묶어 간결하게 작성한다.
- 예: PB·RM·금융상담·여신심사

builder
- 원본 조사자료에서 구축 주체가 확인되는 경우만 작성한다.
- 확인되지 않으면 빈 문자열.

build_method
- Agent Builder, 자체개발, 공동 플랫폼 등의 방식이
  원본에 명시된 경우만 작성한다.
- 추정 금지.

completed_count
- 실제 구축 완료 수량만 기록한다.

first_half_target
- 상반기 목표.

second_half_target
- 하반기 목표.

annual_target
- 연간 목표.

future_plan
- 확인된 향후 구축계획만 요약한다.

source_type
- 출처 유형을 요약한다.
- 예: 공식자료
- 예: 공식자료 + 주요언론

source_detail
- 출처기관과 자료 제목을 간결하게 작성한다.

source_urls
- 실제 원본 URL을 배열로 기록한다.
- 중복 URL을 제거한다.

published_dates
- 발표일을 배열로 기록한다.
- 중복 날짜를 제거한다.

notes
- 계획과 완료의 차이
- 그룹 수치 여부
- Agent 인정 여부
- 교차검증 필요사항
- 수치 해석 주의사항
등을 기록한다.

==================================================
중요 정규화 사례
==================================================

국민은행:

PB, RM, 금융상담, 여신심사 등을
개별 Agent 4개로 임의 계산하지 않는다.

은행 전체 100개 목표가 확인된 경우
annual_target = 100으로 유지하되,
2026-10-07 현재 실제 완료 여부가 확인되지 않았다면
completed_count는 null이다.

KB금융그룹 전체 수치는 국민은행 자체 수치로 환산하지 않는다.


우리은행:

고객관계관리·기업여신·자산관리·내부통제·고객상담·업무자동화 등을
무조건 개별 Agent 숫자로 환산하지 않는다.

전행 AI Agent 구축사업의 90개 목표는
근거 범위에 맞게 목표값으로 유지한다.

에이전틱 AICC가 별도 사업이더라도
중복 수량 여부가 불명확하면 추가 숫자로 합산하지 않는다.


카카오뱅크:

AI 이체, AI 검색, AI 모임총무가
AI 서비스로 공개되었더라도
독립 AI Agent라는 근거가 부족한 경우
Agent 개수로 계산하지 않는다.

notes에 Agent 판단 주의를 남긴다.


K뱅크:

개발지원 Agentic AI,
이해상충거래 셀프체크,
금융사기 메일 자동분류 사례를
자동으로 3개 Agent로 계산하지 않는다.

공개된 배포 및 완료 수량이 없으면
completed_count는 null이다.


부산은행:

BNK금융그룹 공동 생성형 AI 플랫폼 참여는
부산은행 자체 완료 수량으로 환산하지 않는다.

그룹 공동 사업의 업무용 Agent 수 역시
부산은행 자체 구축수량으로 사용하지 않는다.

==================================================
품질 검증
==================================================

저장하기 전에 스스로 아래를 점검한다.

- banks 배열에 정확히 7개 은행이 있는가
- 다른 은행이 포함되지 않았는가
- 신한은행 버티컬 Agent completed_count가 16인가
- 신한은행 일반 Agent first_half_target이 80인가
- 신한은행 일반 Agent second_half_target이 150인가
- 신한은행 일반 Agent annual_target이 200인가
- 그룹 수치를 개별 은행 수치로 변환하지 않았는가
- 확인되지 않은 완료 수치를 0으로 작성하지 않았는가
- source URL을 가능한 한 보존했는가
- 한국어가 정상 UTF-8로 저장되는가

==================================================
원본 조사 JSON
==================================================

$sourceJson

==================================================
최종 실행 지시
==================================================

위 원본 JSON만 사용해서 정규화한다.

반드시 Write 도구를 이용해 다음 파일을 직접 생성한다.

$OutputJson

파일 생성 후 JSON 문법이 유효한지 스스로 확인한다.

채팅 최종 응답에는 긴 JSON을 출력하지 말고
파일 저장이 완료되었다는 짧은 메시지만 출력한다.
"@

# ------------------------------------------------------------
# 4. Claude 실행
# ------------------------------------------------------------

Write-Host "Claude direct-file normalization started..."
Write-Host "Output target:" $OutputJson
Write-Host ""

Push-Location "C:\AI"

try {

    $prompt | & claude -p --allowedTools "Write"

    $claudeExitCode = $LASTEXITCODE

}
finally {

    Pop-Location

}

if ($claudeExitCode -ne 0) {
    throw "Claude 실행 실패. ExitCode: $claudeExitCode"
}

# ------------------------------------------------------------
# 5. 실제 파일 생성 여부 확인
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $OutputJson)) {
    throw "Claude가 출력 JSON 파일을 생성하지 못했습니다: $OutputJson"
}

$fileInfo = Get-Item -LiteralPath $OutputJson

if ($fileInfo.Length -lt 100) {
    throw "출력 JSON 파일 크기가 비정상적으로 작습니다: $($fileInfo.Length) bytes"
}

Write-Host ""
Write-Host "File created :" $fileInfo.FullName
Write-Host "File size    :" $fileInfo.Length


# ------------------------------------------------------------
# 6. UTF-8 직접 읽기 및 JSON 파싱
# ------------------------------------------------------------

$normalizedText = Get-Content `
    -LiteralPath $OutputJson `
    -Raw `
    -Encoding UTF8

try {
    $normalized = $normalizedText | ConvertFrom-Json
}
catch {

    $badPath = "C:\AI\work\temp\normalize_v2_invalid_json.txt"

    Copy-Item `
        -LiteralPath $OutputJson `
        -Destination $badPath `
        -Force

    throw "v2 JSON 파싱 실패. 원본 보존: $badPath"
}

Write-Host "JSON Parse   : OK"


# ------------------------------------------------------------
# 7. 은행 목록 검증
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

if (@($normalized.banks).Count -ne 7) {
    throw "정규화 결과 은행 수 오류: $(@($normalized.banks).Count)"
}

if ($missing.Count -gt 0) {
    throw "누락 은행: $($missing -join ', ')"
}

if ($extra.Count -gt 0) {
    throw "불필요 은행: $($extra -join ', ')"
}

Write-Host "Bank Check   : OK"


# ------------------------------------------------------------
# 8. 신한 보호값 검증
# ------------------------------------------------------------

$shinhan = $normalized.banks |
    Where-Object { $_.bank -eq "신한은행" } |
    Select-Object -First 1

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

if ($vertical.completed_count -ne 16) {
    throw "신한 버티컬 Agent 보호값 오류: completed_count"
}

if ($general.first_half_target -ne 80) {
    throw "신한 일반 Agent 보호값 오류: first_half_target"
}

if ($general.second_half_target -ne 150) {
    throw "신한 일반 Agent 보호값 오류: second_half_target"
}

if ($general.annual_target -ne 200) {
    throw "신한 일반 Agent 보호값 오류: annual_target"
}

Write-Host "Protected Data: OK"


# ------------------------------------------------------------
# 9. 결과 요약
# ------------------------------------------------------------

Write-Host ""
Write-Host "=== NORMALIZATION V2 RESULT ==="

$totalRows = 0

foreach ($bank in $normalized.banks) {

    $rows = @($bank.rows)
    $totalRows += $rows.Count

    Write-Host ""
    Write-Host "[$($bank.bank)] Rows:" $rows.Count

    foreach ($row in $rows) {

        $completed = if ($null -eq $row.completed_count) {
            "-"
        }
        else {
            $row.completed_count
        }

        $first = if ($null -eq $row.first_half_target) {
            "-"
        }
        else {
            $row.first_half_target
        }

        $second = if ($null -eq $row.second_half_target) {
            "-"
        }
        else {
            $row.second_half_target
        }

        $annual = if ($null -eq $row.annual_target) {
            "-"
        }
        else {
            $row.annual_target
        }

        Write-Host (
            " - {0} | 완료:{1} | 상:{2} | 하:{3} | 연간:{4}" -f `
            $row.agent_category,
            $completed,
            $first,
            $second,
            $annual
        )
    }
}

Write-Host ""
Write-Host "Banks      :" @($normalized.banks).Count
Write-Host "Total Rows :" $totalRows
Write-Host "Output     :" $OutputJson
Write-Host ""
Write-Host "STATUS: NORMALIZED_JSON_V2_READY"