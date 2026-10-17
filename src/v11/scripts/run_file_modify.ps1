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

$supportedTextTypes = @(
    ".html",
    ".htm",
    ".md",
    ".txt"
)

if ($ext -notin $supportedTextTypes) {
    throw "현재 MODIFY v1은 HTML/HTM/MD/TXT를 지원합니다. 현재 파일: $ext"
}

$outputDir = "C:\AI\work\output"

if (!(Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$outputPath = Join-Path `
    $outputDir `
    "${baseName}_modified_${dateStamp}${ext}"

# 원본 보존
Copy-Item `
    -LiteralPath $FilePath `
    -Destination $outputPath `
    -Force

Write-Host ""
Write-Host "========================================"
Write-Host " Jungle File Modify"
Write-Host "========================================"
Write-Host ""
Write-Host "ORIGINAL :" $FilePath
Write-Host "COPY     :" $outputPath
Write-Host ""

$prompt = @"
다음 파일의 복사본을 사용자의 요청에 맞게 수정하세요.

[원본 파일]
$FilePath

[수정 대상 복사본]
$outputPath

[사용자 요청]
$UserRequest

[중요 규칙]
- 원본 파일은 절대 수정하지 말 것
- 반드시 수정 대상 복사본만 편집할 것
- 실제 파일 전체를 먼저 읽고 구조와 문맥을 파악할 것
- 사용자가 요청하지 않은 부분은 불필요하게 바꾸지 말 것
- 기존 디자인, CSS, JavaScript, 레이아웃, 기능은 특별한 이유가 없으면 유지할 것
- 수정 요청의 목적과 대상자를 우선할 것
- 기존 내용을 임의로 삭제하거나 축약하지 말 것
- 수정으로 인해 HTML 구조나 기능이 깨지지 않도록 할 것
- 한국어 문장은 자연스럽게 작성할 것
- 수정 후 파일 전체의 일관성을 다시 확인할 것

수정 대상:
$outputPath

수정이 완료되면 다른 설명 없이 MODIFY_FILE_CREATED 만 출력하세요.
"@

claude -p $prompt --allowedTools "Read" "Write" "Edit"

if (!(Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    throw "수정 결과 파일이 생성되지 않았습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE FILE MODIFY RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Modify : OK"
Write-Host "Original:" $FilePath
Write-Host "File    :" $outputPath
Write-Host "Size    :" (Get-Item $outputPath).Length
Write-Host ""
Write-Host "STATUS: FILE_MODIFY_COMPLETE"