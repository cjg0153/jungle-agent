param(
    [Parameter(Mandatory=$true)]
    [string]$DesignJsonPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $DesignJsonPath -PathType Leaf)) {
    throw "디자인 분석 JSON을 찾을 수 없습니다: $DesignJsonPath"
}

$tempDir = "C:\AI\work\temp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($DesignJsonPath)
    $OutputPath = Join-Path $tempDir "${baseName}_layout_library_${stamp}.json"
}

$instructionPath = Join-Path $tempDir "ppt_layout_library_instruction_${stamp}.txt"

$instruction = @"
You are building a reusable PowerPoint layout library for Jungle.

Read ONLY this local PowerPoint design-analysis JSON:
$DesignJsonPath

Do not use web search, connectors, Google Drive, or external sources.

Identify the strongest reusable reference slides. Analyze shape hierarchy, coordinates, text roles, tables, charts, pictures, fills, lines, and visual balance.

Preferred layout_type values:
title, section, agenda, executive_summary, kpi, cards_3, cards_4, comparison_2col, process, roadmap, table, chart, diagram, detail, closing, other

Save valid UTF-8 JSON to:
$OutputPath

Schema:
{
  "source_design_json": "$DesignJsonPath",
  "layouts": [
    {
      "layout_id": "L01",
      "layout_type": "title",
      "reference_slide": 1,
      "layout_name": "short name",
      "quality_score": 0.0,
      "best_for": ["use case"],
      "content_capacity": {
        "title_max_chars": 40,
        "subtitle_max_chars": 80,
        "body_items_recommended": 0
      },
      "replaceable_fields": [
        {
          "role": "title",
          "shape_name": "exact shape name",
          "field_type": "text",
          "source_text": "existing text",
          "max_chars": 40,
          "table_rows": 0,
          "table_cols": 0
        }
      ],
      "preserve_shapes": ["exact shape name"],
      "notes": "short note"
    }
  ],
  "recommended_defaults": {
    "title": "L01",
    "section": "L02",
    "detail": "L03"
  }
}

Rules:
1. Use exact slide numbers and exact shape names from the JSON.
2. field_type must be "text" or "table".
3. If has_table=true, use field_type="table" and copy the exact table_rows/table_cols.
4. Do not invent shapes.
5. quality_score is 0 to 1.
6. Prefer visually strong reusable slides over plain text.
7. Keep 5 to 12 strong layouts.
8. Exclude highly content-specific slides that cannot be safely reused.
9. Decorative shapes, logos, lines, background blocks and page markers belong in preserve_shapes.
10. Tables may be replaceable when their structure is reusable.
11. If a layout type does not exist, do not fabricate it.

After saving the JSON, output only:
PPT_LAYOUT_LIBRARY_CREATED
"@

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($instructionPath, $instruction, $utf8Bom)

$claudeRequest = "Read the instruction file at $instructionPath and execute it exactly. Use only the local files specified there. Do not use web search, Google Drive, connectors, or external sources. When finished, output only PPT_LAYOUT_LIBRARY_CREATED."

claude -p $claudeRequest --allowedTools "Read" "Write"

if (!(Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "레이아웃 라이브러리 JSON이 생성되지 않았습니다."
}

$library = Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($null -eq $library.layouts -or $library.layouts.Count -lt 1) {
    throw "재사용 가능한 레이아웃이 없습니다."
}

Write-Host ""
Write-Host "STATUS: PPT_LAYOUT_LIBRARY_COMPLETE"
Write-Host "Layouts:" $library.layouts.Count
Write-Host "File   :" $OutputPath
