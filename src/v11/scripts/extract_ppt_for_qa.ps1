param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "PPT/PPTX 파일을 찾을 수 없습니다: $FilePath"
}

$ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
if ($ext -notin @(".pptx", ".ppt")) {
    throw "PPT/PPTX 파일만 지원합니다."
}

$tempDir = "C:\AI\work\temp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = Join-Path $tempDir "${baseName}_qa_extract_${stamp}.json"
}

function Get-TextFromShape {
    param($Shape)

    try {
        if ($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1) {
            return [string]$Shape.TextFrame.TextRange.Text
        }
    }
    catch {}

    try {
        if ($Shape.TextFrame2.HasText -eq -1) {
            return [string]$Shape.TextFrame2.TextRange.Text
        }
    }
    catch {}

    return $null
}

function Get-TableData {
    param($Shape)

    $result = @()

    try {
        if ($Shape.HasTable -ne -1) {
            return $result
        }

        $rows = [int]$Shape.Table.Rows.Count
        $cols = [int]$Shape.Table.Columns.Count

        for ($r = 1; $r -le $rows; $r++) {
            $rowItems = @()

            for ($c = 1; $c -le $cols; $c++) {
                $cellText = ""

                try {
                    $cellShape = $Shape.Table.Cell($r, $c).Shape
                    $cellText = [string]$cellShape.TextFrame.TextRange.Text
                }
                catch {}

                $rowItems += [PSCustomObject]@{
                    row = $r
                    col = $c
                    text = $cellText
                }
            }

            $result += ,$rowItems
        }
    }
    catch {}

    return $result
}

function Get-SmartArtData {
    param($Shape)

    $nodes = @()

    try {
        $smartArt = $Shape.SmartArt
        $allNodes = $smartArt.AllNodes

        for ($i = 1; $i -le $allNodes.Count; $i++) {
            $node = $allNodes.Item($i)
            $nodeText = ""

            try {
                $nodeText = [string]$node.TextFrame2.TextRange.Text
            }
            catch {}

            $nodes += [PSCustomObject]@{
                index = $i
                text = $nodeText
            }
        }
    }
    catch {}

    return $nodes
}

function Get-ChartData {
    param($Shape)

    $chartInfo = $null

    try {
        if ($Shape.HasChart -ne -1) {
            return $null
        }

        $chart = $Shape.Chart

        $title = ""
        try {
            if ($chart.HasTitle) {
                $title = [string]$chart.ChartTitle.Text
            }
        }
        catch {}

        $categoryAxisTitle = ""
        $valueAxisTitle = ""

        try {
            $axis = $chart.Axes(1)
            if ($axis.HasTitle) {
                $categoryAxisTitle = [string]$axis.AxisTitle.Text
            }
        }
        catch {}

        try {
            $axis = $chart.Axes(2)
            if ($axis.HasTitle) {
                $valueAxisTitle = [string]$axis.AxisTitle.Text
            }
        }
        catch {}

        $series = @()

        try {
            $seriesCollection = $chart.SeriesCollection()

            for ($s = 1; $s -le $seriesCollection.Count; $s++) {
                $seriesItem = $seriesCollection.Item($s)

                $seriesName = ""
                try { $seriesName = [string]$seriesItem.Name } catch {}

                $labels = @()

                try {
                    $points = $seriesItem.Points()

                    for ($p = 1; $p -le $points.Count; $p++) {
                        $point = $points.Item($p)
                        $labelText = ""

                        try {
                            if ($point.HasDataLabel) {
                                $labelText = [string]$point.DataLabel.Text
                            }
                        }
                        catch {}

                        if (-not [string]::IsNullOrWhiteSpace($labelText)) {
                            $labels += [PSCustomObject]@{
                                point = $p
                                text = $labelText
                            }
                        }
                    }
                }
                catch {}

                $series += [PSCustomObject]@{
                    index = $s
                    name = $seriesName
                    data_labels = $labels
                }
            }
        }
        catch {}

        $chartInfo = [PSCustomObject]@{
            title = $title
            category_axis_title = $categoryAxisTitle
            value_axis_title = $valueAxisTitle
            series = $series
        }
    }
    catch {}

    return $chartInfo
}

function Get-ShapeQaData {
    param(
        $Shape,
        [string]$Prefix = ""
    )

    $items = @()

    $shapeName = ""
    $shapeType = 0
    $altText = ""

    try { $shapeName = [string]$Shape.Name } catch {}
    try { $shapeType = [int]$Shape.Type } catch {}
    try { $altText = [string]$Shape.AlternativeText } catch {}

    $fullName = $shapeName
    if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
        $fullName = "$Prefix > $shapeName"
    }

    $plainText = Get-TextFromShape -Shape $Shape
    $tableData = Get-TableData -Shape $Shape
    $smartArtData = Get-SmartArtData -Shape $Shape
    $chartData = Get-ChartData -Shape $Shape

    $items += [PSCustomObject]@{
        name = $fullName
        type = $shapeType
        text = $plainText
        alternative_text = $altText
        table = $tableData
        smartart = $smartArtData
        chart = $chartData
    }

    # msoGroup = 6
    if ($shapeType -eq 6) {
        try {
            $groupItems = $Shape.GroupItems

            for ($i = 1; $i -le $groupItems.Count; $i++) {
                $child = $groupItems.Item($i)
                $childItems = Get-ShapeQaData -Shape $child -Prefix $fullName
                $items += $childItems
            }
        }
        catch {}
    }

    return $items
}

function Get-NotesText {
    param($Slide)

    $noteTexts = @()

    try {
        $notesPage = $Slide.NotesPage

        foreach ($shape in $notesPage.Shapes) {
            $text = Get-TextFromShape -Shape $shape

            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $trimmed = $text.Trim()

                # Ignore boilerplate slide number/date placeholders when possible.
                if ($trimmed -notmatch '^\d+$' -and
                    $trimmed -notmatch '^<.*>$' -and
                    $trimmed -notmatch '^(Date|Footer|Slide Number)$') {
                    $noteTexts += $trimmed
                }
            }
        }
    }
    catch {}

    return $noteTexts
}

function Build-SlideQaText {
    param(
        [int]$SlideNumber,
        [string]$Title,
        $Shapes,
        $Notes
    )

    $lines = @()
    $lines += "[Slide $SlideNumber]"

    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $lines += "TITLE: $Title"
    }

    foreach ($shape in $Shapes) {
        if (-not [string]::IsNullOrWhiteSpace([string]$shape.text)) {
            $lines += "TEXT($($shape.name)): $($shape.text)"
        }

        if ($null -ne $shape.table -and $shape.table.Count -gt 0) {
            $lines += "TABLE($($shape.name)):"

            foreach ($row in $shape.table) {
                $cells = @()

                foreach ($cell in $row) {
                    $cells += [string]$cell.text
                }

                $lines += ("  " + ($cells -join " | "))
            }
        }

        if ($null -ne $shape.smartart -and $shape.smartart.Count -gt 0) {
            foreach ($node in $shape.smartart) {
                if (-not [string]::IsNullOrWhiteSpace([string]$node.text)) {
                    $lines += "SMARTART($($shape.name))[$($node.index)]: $($node.text)"
                }
            }
        }

        if ($null -ne $shape.chart) {
            if (-not [string]::IsNullOrWhiteSpace([string]$shape.chart.title)) {
                $lines += "CHART TITLE($($shape.name)): $($shape.chart.title)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$shape.chart.category_axis_title)) {
                $lines += "CHART X AXIS($($shape.name)): $($shape.chart.category_axis_title)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$shape.chart.value_axis_title)) {
                $lines += "CHART Y AXIS($($shape.name)): $($shape.chart.value_axis_title)"
            }

            foreach ($series in $shape.chart.series) {
                if (-not [string]::IsNullOrWhiteSpace([string]$series.name)) {
                    $lines += "CHART SERIES($($shape.name)): $($series.name)"
                }

                foreach ($label in $series.data_labels) {
                    $lines += "CHART LABEL($($shape.name))[$($label.point)]: $($label.text)"
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$shape.alternative_text)) {
            $lines += "ALT($($shape.name)): $($shape.alternative_text)"
        }
    }

    foreach ($note in $Notes) {
        $lines += "NOTE: $note"
    }

    return ($lines -join "`n")
}

$ppt = $null
$pres = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application

    $pres = $ppt.Presentations.Open(
        $FilePath,
        $false,
        $true,
        $false
    )

    $slideResults = @()
    $allQaText = @()

    foreach ($slide in $pres.Slides) {
        $slideNumber = [int]$slide.SlideIndex

        $titleText = ""
        try {
            if ($null -ne $slide.Shapes.Title) {
                $titleText = [string]$slide.Shapes.Title.TextFrame.TextRange.Text
            }
        }
        catch {}

        $shapeResults = @()

        foreach ($shape in $slide.Shapes) {
            $shapeResults += Get-ShapeQaData -Shape $shape
        }

        $notes = Get-NotesText -Slide $slide

        $qaText = Build-SlideQaText `
            -SlideNumber $slideNumber `
            -Title $titleText `
            -Shapes $shapeResults `
            -Notes $notes

        $slideResults += [PSCustomObject]@{
            slide_number = $slideNumber
            title = $titleText
            shapes = $shapeResults
            notes = $notes
            qa_text = $qaText
        }

        $allQaText += $qaText
    }

    $result = [PSCustomObject]@{
        source = $FilePath
        extraction_type = "ppt_qa_full"
        slide_count = [int]$pres.Slides.Count
        extracted_at = (Get-Date).ToString("s")
        coverage = [PSCustomObject]@{
            standard_text = $true
            tables = $true
            grouped_shapes = $true
            smartart = $true
            chart_titles_axes_series_labels = $true
            notes = $true
            alternative_text = $true
            rendered_visual_layout = $false
        }
        slides = $slideResults
        all_qa_text = ($allQaText -join "`n`n")
    }

    $json = $result | ConvertTo-Json -Depth 20

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText(
        $OutputPath,
        $json,
        $utf8Bom
    )
}
finally {
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
        }
        catch {}
    }

    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }
        catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (!(Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "QA 추출 JSON 생성에 실패했습니다."
}

$qa = Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8 | ConvertFrom-Json

$tableShapeCount = 0
$smartArtNodeCount = 0
$chartCount = 0
$noteCount = 0

foreach ($slide in $qa.slides) {
    $noteCount += @($slide.notes).Count

    foreach ($shape in $slide.shapes) {
        if ($null -ne $shape.table -and @($shape.table).Count -gt 0) {
            $tableShapeCount++
        }

        if ($null -ne $shape.smartart) {
            $smartArtNodeCount += @($shape.smartart).Count
        }

        if ($null -ne $shape.chart) {
            $chartCount++
        }
    }
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE PPT QA EXTRACTOR"
Write-Host "========================================"
Write-Host ""
Write-Host "Source         :" $FilePath
Write-Host "Slides         :" $qa.slide_count
Write-Host "Table shapes   :" $tableShapeCount
Write-Host "SmartArt nodes :" $smartArtNodeCount
Write-Host "Charts         :" $chartCount
Write-Host "Notes          :" $noteCount
Write-Host "File           :" $OutputPath
Write-Host "Size           :" (Get-Item $OutputPath).Length
Write-Host ""
Write-Host "STATUS: PPT_QA_EXTRACT_COMPLETE"
