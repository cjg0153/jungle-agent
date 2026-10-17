# JUNGLE_SMART_SELECTIVE_VISUAL_QA_V1

function Get-JungleSelectiveVisualQaDecision {

    param(
        [int[]]$ChangedSlides = @(),

        [bool]$ByteIdentical = $false,

        [int]$MaxSlidesPerCall = 8
    )


    if ($MaxSlidesPerCall -lt 1) {

        throw "MaxSlidesPerCall must be at least 1."
    }


    $normalized =
        @(
            $ChangedSlides |
            Sort-Object -Unique
        )


    if (
        $ByteIdentical `
        -or
        $normalized.Count -eq 0
    ) {

        return [pscustomobject]@{

            visual_qa_required =
                $false

            action =
                "SKIP_NO_VISUAL_CHANGE"

            provider =
                "NONE"

            mode =
                "NONE"

            changed_count =
                0

            changed_slides =
                @()

            max_slides_per_call =
                $MaxSlidesPerCall

            estimated_ai_calls =
                0

            live_execution =
                $false
        }
    }


    $estimatedCalls =
        [int][Math]::Ceiling(
            $normalized.Count /
            [double]$MaxSlidesPerCall
        )


    return [pscustomobject]@{

        visual_qa_required =
            $true

        action =
            "PREPARE_CHANGED_ONLY_BLIND_PAIRWISE"

        provider =
            "CLAUDE"

        mode =
            "BLIND_PAIRWISE_CHANGED_ONLY"

        changed_count =
            $normalized.Count

        changed_slides =
            @($normalized)

        max_slides_per_call =
            $MaxSlidesPerCall

        estimated_ai_calls =
            $estimatedCalls

        live_execution =
            $false
    }
}


function Write-JungleSelectiveVisualQaRequest {

    param(
        [Parameter(Mandatory=$true)]
        [object]$Decision,

        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )


    $record =
        [ordered]@{

            protocol =
                "JUNGLE_SMART_SELECTIVE_VISUAL_QA_V1"

            source =
                $SourcePath

            artifact =
                $ArtifactPath

            visual_qa_required =
                $Decision.visual_qa_required

            action =
                $Decision.action

            provider =
                $Decision.provider

            mode =
                $Decision.mode

            changed_count =
                $Decision.changed_count

            changed_slides =
                @($Decision.changed_slides)

            max_slides_per_call =
                $Decision.max_slides_per_call

            estimated_ai_calls =
                $Decision.estimated_ai_calls

            live_execution =
                $Decision.live_execution

            created_at =
                (Get-Date).ToString("s")
        }


    $json =
        $record |
        ConvertTo-Json -Depth 6


    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )


    [System.IO.File]::WriteAllText(
        $OutputPath,
        $json,
        $encoding
    )


    return $OutputPath
}