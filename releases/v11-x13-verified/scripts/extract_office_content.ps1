param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()

$result = [ordered]@{
    source_file = $FilePath
    file_type   = $ext
    extracted_at = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    content     = @()
}

try {

    if ($ext -eq ".pptx" -or $ext -eq ".ppt") {

        $ppt = New-Object -ComObject PowerPoint.Application
        # PowerPoint는 Visible=False를 허용하지 않음. WithWindow=False로 백그라운드 오픈.

        $presentation = $ppt.Presentations.Open(
            $FilePath,
            $true,
            $true,
            $false
        )

        foreach ($slide in $presentation.Slides) {

            $slideItem = [ordered]@{
                slide_number = $slide.SlideIndex
                texts        = @()
                notes        = @()
            }

            foreach ($shape in $slide.Shapes) {

                try {
                    if (
                        $shape.HasTextFrame -eq -1 -and
                        $shape.TextFrame.HasText -eq -1
                    ) {
                        $text = $shape.TextFrame.TextRange.Text

                        if (![string]::IsNullOrWhiteSpace($text)) {
                            $slideItem.texts += $text
                        }
                    }
                }
                catch {}
            }

            try {
                foreach ($shape in $slide.NotesPage.Shapes) {

                    if (
                        $shape.HasTextFrame -eq -1 -and
                        $shape.TextFrame.HasText -eq -1
                    ) {
                        $noteText = $shape.TextFrame.TextRange.Text

                        if (
                            ![string]::IsNullOrWhiteSpace($noteText) -and
                            $noteText -notmatch '^\s*\d+\s*$'
                        ) {
                            $slideItem.notes += $noteText
                        }
                    }
                }
            }
            catch {}

            $result.content += [PSCustomObject]$slideItem
        }

        $presentation.Close()
        $ppt.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }

    elseif ($ext -eq ".docx" -or $ext -eq ".doc") {

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false

        $document = $word.Documents.Open(
            $FilePath,
            $false,
            $true
        )

        $paragraphNumber = 0

        foreach ($paragraph in $document.Paragraphs) {

            $paragraphNumber++

            $text = $paragraph.Range.Text
            $text = $text -replace "[`r`a]", ""
            $text = $text.Trim()

            if (![string]::IsNullOrWhiteSpace($text)) {

                $result.content += [PSCustomObject]@{
                    paragraph_number = $paragraphNumber
                    text             = $text
                    style            = [string]$paragraph.Range.Style.NameLocal
                }
            }
        }

        $document.Close($false)
        $word.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    else {
        throw "지원하지 않는 Office 파일 형식입니다: $ext"
    }

    $json = $result | ConvertTo-Json -Depth 10

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)

    [System.IO.File]::WriteAllText(
        $OutputPath,
        $json,
        $utf8Bom
    )

    Write-Host ""
    Write-Host "========================================"
    Write-Host " JUNGLE OFFICE EXTRACT"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Source :" $FilePath
    Write-Host "Output :" $OutputPath
    Write-Host "Type   :" $ext
    Write-Host ""
    Write-Host "STATUS: OFFICE_CONTENT_EXTRACTED"
}
finally {

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}