param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $FilePath -PathType Leaf)) {
    throw "PPTX 파일을 찾을 수 없습니다: $FilePath"
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
    $OutputPath = Join-Path $tempDir "${baseName}_design_analysis_${stamp}.json"
}

function Get-SafeProperty {
    param($Object, [string]$PropertyName, $Default = $null)
    try {
        $value = $Object.$PropertyName
        if ($null -eq $value) { return $Default }
        return $value
    }
    catch { return $Default }
}

function Get-RgbHex {
    param($ColorFormat)
    try {
        $rgb = [int]$ColorFormat.RGB
        $r = $rgb -band 0xFF
        $g = ($rgb -shr 8) -band 0xFF
        $b = ($rgb -shr 16) -band 0xFF
        return ("#{0:X2}{1:X2}{2:X2}" -f $r, $g, $b)
    }
    catch { return $null }
}

$ppt = $null
$pres = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($FilePath, $false, $true, $false)

    $presentationWidth  = [double]$pres.PageSetup.SlideWidth
    $presentationHeight = [double]$pres.PageSetup.SlideHeight
    $slides = @()

    foreach ($slide in $pres.Slides) {
        $shapeItems = @()

        foreach ($shape in $slide.Shapes) {
            $text = $null
            $fontName = $null
            $fontSize = $null
            $fontBold = $null

            try {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $text = [string]$shape.TextFrame.TextRange.Text
                    if ($text.Length -gt 800) { $text = $text.Substring(0, 800) }
                    try { $fontName = [string]$shape.TextFrame.TextRange.Font.Name } catch {}
                    try { $fontSize = [double]$shape.TextFrame.TextRange.Font.Size } catch {}
                    try { $fontBold = [int]$shape.TextFrame.TextRange.Font.Bold } catch {}
                }
            }
            catch {}

            $placeholderType = $null
            try {
                if ($shape.Type -eq 14) {
                    $placeholderType = [int]$shape.PlaceholderFormat.Type
                }
            }
            catch {}

            $fillVisible = $null
            $fillColor = $null
            $fillTransparency = $null
            try {
                $fillVisible = [int]$shape.Fill.Visible
                if ($fillVisible -eq -1) {
                    $fillColor = Get-RgbHex $shape.Fill.ForeColor
                    try { $fillTransparency = [double]$shape.Fill.Transparency } catch {}
                }
            }
            catch {}

            $lineVisible = $null
            $lineColor = $null
            try {
                $lineVisible = [int]$shape.Line.Visible
                if ($lineVisible -eq -1) {
                    $lineColor = Get-RgbHex $shape.Line.ForeColor
                }
            }
            catch {}

            $hasTable = $false
            $tableRows = 0
            $tableCols = 0
            $tableCells = @()

            try {
                if ($shape.HasTable -eq -1) {
                    $hasTable = $true
                    $tableRows = [int]$shape.Table.Rows.Count
                    $tableCols = [int]$shape.Table.Columns.Count

                    for ($r = 1; $r -le $tableRows; $r++) {
                        $row = @()
                        for ($c = 1; $c -le $tableCols; $c++) {
                            $cellText = ""
                            try {
                                $cellText = [string]$shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text
                            }
                            catch {}
                            $row += $cellText
                        }
                        $tableCells += ,$row
                    }
                }
            }
            catch {}

            $hasChart = $false
            try {
                if ($shape.HasChart -eq -1) { $hasChart = $true }
            }
            catch {}

            $shapeItems += [PSCustomObject]@{
                name = [string](Get-SafeProperty $shape "Name" "")
                type = [int](Get-SafeProperty $shape "Type" 0)
                placeholder_type = $placeholderType
                left = [math]::Round([double](Get-SafeProperty $shape "Left" 0), 2)
                top = [math]::Round([double](Get-SafeProperty $shape "Top" 0), 2)
                width = [math]::Round([double](Get-SafeProperty $shape "Width" 0), 2)
                height = [math]::Round([double](Get-SafeProperty $shape "Height" 0), 2)
                text = $text
                font_name = $fontName
                font_size = $fontSize
                font_bold = $fontBold
                fill_visible = $fillVisible
                fill_color = $fillColor
                fill_transparency = $fillTransparency
                line_visible = $lineVisible
                line_color = $lineColor
                has_table = $hasTable
                table_rows = $tableRows
                table_cols = $tableCols
                table_cells = $tableCells
                has_chart = $hasChart
            }
        }

        $layoutName = $null
        try { $layoutName = [string]$slide.CustomLayout.Name } catch {}

        $titleText = $null
        try {
            if ($null -ne $slide.Shapes.Title) {
                $titleText = [string]$slide.Shapes.Title.TextFrame.TextRange.Text
            }
        }
        catch {}

        $shapeCount = [int]$slide.Shapes.Count
        $textShapeCount = @($shapeItems | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.text) }).Count
        $pictureCount = @($shapeItems | Where-Object { $_.type -in @(11, 13) }).Count
        $tableCount = @($shapeItems | Where-Object { $_.has_table -eq $true }).Count
        $chartCount = @($shapeItems | Where-Object { $_.has_chart -eq $true }).Count

        $slides += [PSCustomObject]@{
            slide_number = [int]$slide.SlideIndex
            layout_name = $layoutName
            title = $titleText
            shape_count = $shapeCount
            text_shape_count = $textShapeCount
            picture_count = $pictureCount
            table_count = $tableCount
            chart_count = $chartCount
            shapes = $shapeItems
        }
    }

    $result = [PSCustomObject]@{
        source = $FilePath
        slide_width = $presentationWidth
        slide_height = $presentationHeight
        slide_count = [int]$pres.Slides.Count
        analyzed_at = (Get-Date).ToString("s")
        slides = $slides
    }

    $json = $result | ConvertTo-Json -Depth 14
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($OutputPath, $json, $utf8Bom)
}
finally {
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (!(Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "디자인 분석 JSON 생성에 실패했습니다."
}

Write-Host ""
Write-Host "STATUS: PPT_DESIGN_ANALYSIS_COMPLETE"
Write-Host "File  :" $OutputPath
