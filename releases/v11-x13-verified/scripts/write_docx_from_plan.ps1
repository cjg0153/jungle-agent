param(
    [Parameter(Mandatory=$true)]
    [string]$PlanPath,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [string]$TemplatePath = "C:\AI\templates\base.docx"
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $PlanPath -PathType Leaf)) {
    throw "생성계획 JSON을 찾을 수 없습니다: $PlanPath"
}

if (!(Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
    throw "DOCX 템플릿을 찾을 수 없습니다: $TemplatePath"
}

$plan = Get-Content `
    -LiteralPath $PlanPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace([string]$plan.title)) {
    throw "생성계획에 title이 없습니다."
}

$outputDir = Split-Path -Parent $OutputPath

if (!(Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 템플릿 복사
Copy-Item `
    -LiteralPath $TemplatePath `
    -Destination $OutputPath `
    -Force

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertTo-XmlEscapedText {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($Text)
}

function New-ParagraphXml {
    param(
        [string]$Text,
        [string]$StyleId
    )

    $escaped = ConvertTo-XmlEscapedText $Text

    $styleXml = ""

    if (![string]::IsNullOrWhiteSpace($StyleId)) {
        $styleXml = "<w:pPr><w:pStyle w:val=`"$StyleId`"/></w:pPr>"
    }

    return @"
<w:p>
$styleXml
<w:r>
<w:t xml:space="preserve">$escaped</w:t>
</w:r>
</w:p>
"@
}

$zip = $null

try {

    # 기존 문서의 페이지 설정(sectPr)은 유지
    $readZip = [System.IO.Compression.ZipFile]::OpenRead($OutputPath)

    $existingEntry = $readZip.GetEntry("word/document.xml")

    if ($null -eq $existingEntry) {
        $readZip.Dispose()
        throw "템플릿에서 word/document.xml을 찾지 못했습니다."
    }

    $reader = New-Object System.IO.StreamReader(
        $existingEntry.Open(),
        [System.Text.Encoding]::UTF8
    )

    $existingXmlText = $reader.ReadToEnd()
    $reader.Close()
    $readZip.Dispose()

    [xml]$existingXml = $existingXmlText

    $ns = New-Object System.Xml.XmlNamespaceManager($existingXml.NameTable)
    $ns.AddNamespace(
        "w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    )

    $sectNode = $existingXml.SelectSingleNode(
        "//w:body/w:sectPr",
        $ns
    )

    $sectXml = ""

    if ($null -ne $sectNode) {
        $sectXml = $sectNode.OuterXml
    }

    # 새 document.xml 본문 생성
    $body = New-Object System.Text.StringBuilder

    [void]$body.Append(
        (New-ParagraphXml `
            -Text ([string]$plan.title) `
            -StyleId "Title")
    )

    foreach ($section in $plan.sections) {

        if (![string]::IsNullOrWhiteSpace([string]$section.heading)) {

            [void]$body.Append(
                (New-ParagraphXml `
                    -Text ([string]$section.heading) `
                    -StyleId "Heading1")
            )
        }

        foreach ($paragraph in $section.paragraphs) {

            if (![string]::IsNullOrWhiteSpace([string]$paragraph)) {

                [void]$body.Append(
                    (New-ParagraphXml `
                        -Text ([string]$paragraph) `
                        -StyleId "Normal")
                )
            }
        }
    }

    if (![string]::IsNullOrWhiteSpace($sectXml)) {
        [void]$body.Append($sectXml)
    }

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
$($body.ToString())
</w:body>
</w:document>
"@

    # DOCX ZIP 내부 document.xml 교체
    $zip = [System.IO.Compression.ZipFile]::Open(
        $OutputPath,
        [System.IO.Compression.ZipArchiveMode]::Update
    )

    $oldEntry = $zip.GetEntry("word/document.xml")

    if ($null -ne $oldEntry) {
        $oldEntry.Delete()
    }

    $newEntry = $zip.CreateEntry(
        "word/document.xml",
        [System.IO.Compression.CompressionLevel]::Optimal
    )

    $stream = $newEntry.Open()

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    $writer = New-Object System.IO.StreamWriter(
        $stream,
        $utf8NoBom
    )

    $writer.Write($documentXml)
    $writer.Flush()
    $writer.Close()

}
finally {

    if ($null -ne $zip) {
        $zip.Dispose()
    }
}

if (!(Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "DOCX 생성에 실패했습니다."
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE OPENXML DOCX RESULT"
Write-Host "========================================"
Write-Host ""
Write-Host "Create : OK"
Write-Host "File   :" $OutputPath
Write-Host "Size   :" (Get-Item $OutputPath).Length
Write-Host ""
Write-Host "STATUS: OPENXML_DOCX_CREATE_COMPLETE"