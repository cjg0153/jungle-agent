param(
    [string]$TemplatePath = "C:\AI\work\input\은행별_AI_Agent_구축현황_조사표_20261006.xlsx",
    [string]$JsonPath     = "C:\AI\work\temp\bank_agent_excel_normalized_20261007.json",
    [string]$OutputPath   = "C:\AI\work\output\은행별_AI_Agent_구축현황_조사표_20261007_완성.xlsx"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle Excel Writer"
Write-Host "========================================"
Write-Host ""

# ------------------------------------------------------------
# 1. 파일 확인
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $TemplatePath)) {
    throw "원본 Excel 파일 없음: $TemplatePath"
}

if (!(Test-Path -LiteralPath $JsonPath)) {
    throw "정규화 JSON 없음: $JsonPath"
}

$outputDir = Split-Path -Parent $OutputPath

if (!(Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# ------------------------------------------------------------
# 2. JSON 읽기
# ------------------------------------------------------------

$jsonText = Get-Content `
    -LiteralPath $JsonPath `
    -Raw `
    -Encoding UTF8

try {
    $data = $jsonText | ConvertFrom-Json
}
catch {
    throw "정규화 JSON 파싱 실패"
}

if (@($data.banks).Count -ne 7) {
    throw "은행 수가 7개가 아닙니다."
}

Write-Host "Normalized Banks :" @($data.banks).Count

# ------------------------------------------------------------
# 3. 원본 복사
# ------------------------------------------------------------

Copy-Item `
    -LiteralPath $TemplatePath `
    -Destination $OutputPath `
    -Force

Write-Host "Workbook copied   :" $OutputPath

# ------------------------------------------------------------
# 4. Excel 실행
# ------------------------------------------------------------

$excel = $null
$wb = $null
$ws = $null

try {

    $excel = New-Object -ComObject Excel.Application

    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($OutputPath)

    $ws = $wb.Worksheets.Item("조사표")

    # --------------------------------------------------------
    # 5. 제목 기준일 변경
    # --------------------------------------------------------

    $ws.Range("A1").Value2 =
        "은행별 AI Agent 구축 현황 조사표 (기준일: 2026.10.07)"

    # --------------------------------------------------------
    # 6. 은행별 사용 가능한 행
    # --------------------------------------------------------

    $rowMap = @{
        "신한은행"   = @(6,7)
        "우리은행"   = @(8,9)
        "국민은행"   = @(10,11)
        "하나은행"   = @(12,13)
        "부산은행"   = @(14,15)
        "카카오뱅크" = @(16,17)
        "K뱅크"      = @(18,19)
    }

    # --------------------------------------------------------
    # 7. 기존 K열 자동계산 공식 보장
    # --------------------------------------------------------

    for ($r = 6; $r -le 19; $r++) {

        $k = $ws.Cells.Item($r, 11)

        if (!$k.HasFormula) {

            $k.Formula =
                '=IF(OR(J' + $r + '="",G' + $r +
                '=""),"",MAX(J' + $r + '-G' + $r + ',0))'
        }
    }

    # --------------------------------------------------------
    # 8. 조사 담당 이외 입력영역 초기화
    #
    # 신한 6~7행은 사용자 기존값 보호를 위해 초기화하지 않음
    # 다른 은행은 C:J, L:P만 비움
    # K열 수식은 보호
    # A/B열 번호 및 은행명 보호
    # --------------------------------------------------------

    foreach ($bankName in @(
        "우리은행",
        "국민은행",
        "하나은행",
        "부산은행",
        "카카오뱅크",
        "K뱅크"
    )) {

        foreach ($r in $rowMap[$bankName]) {

            $ws.Range("C$r`:J$r").ClearContents()
            $ws.Range("L$r`:P$r").ClearContents()
        }
    }

    # --------------------------------------------------------
    # Helper : 값 입력
    # --------------------------------------------------------

    function Set-CellValue {

        param(
            $sheet,
            [int]$row,
            [int]$column,
            $value
        )

        $cell = $sheet.Cells.Item($row,$column)

        if ($null -eq $value) {
            $cell.ClearContents()
            return
        }

        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $cell.ClearContents()
            return
        }

        # Excel COM은 Int32 직접 대입 시 형변환 오류가 날 수 있으므로
        # 숫자는 Double, 나머지는 String으로 명시 변환한다.
        if (
            $value -is [byte]    -or
            $value -is [int16]   -or
            $value -is [int32]   -or
            $value -is [int64]   -or
            $value -is [single]  -or
            $value -is [double]  -or
            $value -is [decimal]
        ) {
            $cell.Value2 = [double]$value
        }
        else {
            $cell.Value2 = [string]$value
        }
    }

    # --------------------------------------------------------
    # Helper : 출처 상세 구성
    # --------------------------------------------------------

    function Build-SourceDetail {

        param($record)

        $parts = @()

        if (
            $null -ne $record.source_detail -and
            "$($record.source_detail)".Trim() -ne ""
        ) {
            $parts += "$($record.source_detail)"
        }

        $urls  = @($record.source_urls)
        $dates = @($record.published_dates)

        for ($i = 0; $i -lt $urls.Count; $i++) {

            $url = "$($urls[$i])"

            if ([string]::IsNullOrWhiteSpace($url)) {
                continue
            }

            $date = ""

            if ($i -lt $dates.Count) {
                $date = "$($dates[$i])"
            }

            if ($date -ne "") {
                $parts += "$date / $url"
            }
            else {
                $parts += $url
            }
        }

        return ($parts -join "`n")
    }

    # --------------------------------------------------------
    # 9. 데이터 입력
    # --------------------------------------------------------

    foreach ($bank in $data.banks) {

        $bankName = "$($bank.bank)"

        if (!$rowMap.ContainsKey($bankName)) {
            throw "Excel Row Map에 없는 은행: $bankName"
        }

        $rows = @($bank.rows)
        $slots = @($rowMap[$bankName])

        if ($rows.Count -gt $slots.Count) {
            throw "$bankName 의 정규화 행이 Excel 슬롯보다 많습니다."
        }

        Write-Host ""
        Write-Host "[$bankName]"

        for ($i = 0; $i -lt $rows.Count; $i++) {

            $record = $rows[$i]
            $r = $slots[$i]

            Write-Host " Row $r :" $record.agent_category

            # =================================================
            # 신한은행
            # 기존 사용자 데이터 보호
            # =================================================

            if ($bankName -eq "신한은행") {

                # 사용자 입력 숫자와 기존 분류는 건드리지 않음.
                #
                # 필요한 경우 출처와 비고만 추가한다.

                $existingSource = "$($ws.Cells.Item($r,13).Value2)"
                $existingNote   = "$($ws.Cells.Item($r,16).Value2)"

                $newSource = Build-SourceDetail $record
                $newNote   = "$($record.notes)"

                # 출처 상세 Append
                if (![string]::IsNullOrWhiteSpace($newSource)) {

                    if (![string]::IsNullOrWhiteSpace($existingSource)) {

                        if ($existingSource -notlike "*$newSource*") {

                            $ws.Cells.Item($r,13).Value2 =
                                $existingSource +
                                "`n[공개자료 교차검증]" +
                                "`n" +
                                $newSource
                        }
                    }
                    else {

                        $ws.Cells.Item($r,13).Value2 = $newSource
                    }
                }

                # 비고 Append
                if (![string]::IsNullOrWhiteSpace($newNote)) {

                    if (![string]::IsNullOrWhiteSpace($existingNote)) {

                        if ($existingNote -notlike "*$newNote*") {

                            $ws.Cells.Item($r,16).Value2 =
                                $existingNote +
                                "`n" +
                                $newNote
                        }
                    }
                    else {

                        $ws.Cells.Item($r,16).Value2 = $newNote
                    }
                }

                # 확인일만 입력
                $ws.Cells.Item($r,14).Value2 = "2026.10.07"

                continue
            }

            # =================================================
            # 나머지 은행
            # =================================================

            Set-CellValue $ws $r 3  $record.agent_category
            Set-CellValue $ws $r 4  $record.target_work
            Set-CellValue $ws $r 5  $record.builder
            Set-CellValue $ws $r 6  $record.build_method

            Set-CellValue $ws $r 7  $record.completed_count
            Set-CellValue $ws $r 8  $record.first_half_target
            Set-CellValue $ws $r 9  $record.second_half_target
            Set-CellValue $ws $r 10 $record.annual_target

            # K열은 자동 계산

            Set-CellValue $ws $r 12 $record.source_type

            $sourceDetail = Build-SourceDetail $record

            Set-CellValue $ws $r 13 $sourceDetail

            # 확인일
            Set-CellValue $ws $r 14 "2026.10.07"

            # 조사담당은 자동입력하지 않음
            # O열 유지/공란

            Set-CellValue $ws $r 16 $record.notes
        }
    }

    # --------------------------------------------------------
    # 10. WrapText
    # --------------------------------------------------------

    $ws.Range("C6:P19").WrapText = $true

    # --------------------------------------------------------
    # 11. 자동계산
    # --------------------------------------------------------

    $excel.CalculateFull()

    # --------------------------------------------------------
    # 12. 저장
    # --------------------------------------------------------

    $wb.Save()

    Write-Host ""
    Write-Host "Workbook Saved."
}
finally {

    if ($null -ne $wb) {
        $wb.Close($true)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    if ($null -ne $ws) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws)
    }

    if ($null -ne $wb) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
    }

    if ($null -ne $excel) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ------------------------------------------------------------
# 13. 결과 확인
# ------------------------------------------------------------

if (!(Test-Path -LiteralPath $OutputPath)) {
    throw "최종 Excel 파일이 생성되지 않았습니다."
}

$result = Get-Item -LiteralPath $OutputPath

Write-Host ""
Write-Host "========================================"
Write-Host " EXCEL WRITE RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "File :" $result.FullName
Write-Host "Size :" $result.Length
Write-Host ""
Write-Host "STATUS: EXCEL_WRITE_READY"