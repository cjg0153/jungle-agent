param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "DOCX 파일을 찾을 수 없습니다: $FilePath"
}

if ([System.IO.Path]::GetExtension($FilePath).ToLowerInvariant() -ne ".docx") {
    throw "DOCX 전용 추출기입니다."
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $tempDir = "C:\AI\work\temp"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $OutputPath = Join-Path $tempDir "${base}_docx_qa_extract_$(Get-Date -Format yyyyMMdd_HHmmss).json"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)

try {
    $entryNames = New-Object System.Collections.Generic.List[string]
    $entryNames.Add("word/document.xml")

    foreach ($e in $zip.Entries) {
        if ($e.FullName -match '^word/(header|footer)\d+\.xml$') {
            $entryNames.Add($e.FullName)
        }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $paragraphCount = 0
    $tableCount = 0
    $cellCount = 0

    foreach ($entryName in $entryNames) {
        $entry = $zip.GetEntry($entryName)
        if ($null -eq $entry) { continue }

        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        try { $xmlText = $reader.ReadToEnd() }
        finally {
            $reader.Close()
            $stream.Close()
        }

        [xml]$xml = $xmlText
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("w","http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        $lines.Add("SECTION: $entryName")

        # Exclude paragraphs inside tables to prevent duplicate QA text.
        $paras = $xml.SelectNodes("//w:p[not(ancestor::w:tbl)]",$ns)
        foreach ($p in $paras) {
            $texts = @($p.SelectNodes(".//w:t",$ns) | ForEach-Object { $_.'#text' })
            $text = ($texts -join "")
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $paragraphCount++
                $lines.Add("P($paragraphCount): $text")
            }
        }

        $tables = $xml.SelectNodes("//w:tbl",$ns)
        foreach ($tbl in $tables) {
            $tableCount++
            $lines.Add("TABLE($tableCount):")
            $rowIndex = 0

            foreach ($tr in $tbl.SelectNodes("./w:tr",$ns)) {
                $rowIndex++
                $cells = New-Object System.Collections.Generic.List[string]

                foreach ($tc in $tr.SelectNodes("./w:tc",$ns)) {
                    $parts = @($tc.SelectNodes(".//w:t",$ns) | ForEach-Object { $_.'#text' })
                    $cellText = ($parts -join "")
                    $cellCount++
                    $cells.Add($cellText)
                }

                $lines.Add("  ROW($rowIndex): " + ($cells -join " | "))
            }
        }
    }

    $result = [PSCustomObject]@{
        source = $FilePath
        artifact = "DOCX"
        paragraph_count = $paragraphCount
        table_count = $tableCount
        cell_count = $cellCount
        coverage = [PSCustomObject]@{
            body_text = $true
            tables = $true
            headers_footers = $true
            rendered_visual_layout = $false
        }
        all_qa_text = ($lines -join "`r`n")
    }

    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText(
        $OutputPath,
        ($result | ConvertTo-Json -Depth 8),
        $enc
    )
}
finally {
    $zip.Dispose()
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE DOCX QA EXTRACTOR"
Write-Host "========================================"
Write-Host "Paragraphs :" $paragraphCount
Write-Host "Tables     :" $tableCount
Write-Host "Cells      :" $cellCount
Write-Host "File       :" $OutputPath
Write-Host "STATUS: DOCX_QA_EXTRACT_COMPLETE"
