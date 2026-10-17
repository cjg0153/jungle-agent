# JUNGLE_SMART_ROUTER_ASCII_REGEX_V1

function Get-JungleRoute {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Request,

        [Parameter(Mandatory=$true)]
        [object]$Config
    )

    # PPT / PowerPoint detection.
    # Korean "PowerPoint" is encoded as Unicode escapes
    # so Windows PowerShell 5.1 source decoding cannot corrupt it.
    $pptPattern =
        '(?i)(\.pptx?|\bppt\b|powerpoint|\uD30C\uC6CC\uD3EC\uC778\uD2B8)'


    # Visual intent:
    # design
    # layout
    # visual
    # style
    # format
    # align
    # spacing
    #
    # Korean:
    # design / layout / visual / tidy / align / spacing / executive report
    $visualPattern =
        '(?i)(design|layout|visual|style|format|align|spacing|\uB514\uC790\uC778|\uB808\uC774\uC544\uC6C3|\uC2DC\uAC01|\uC815\uB3C8|\uC815\uB82C|\uAC04\uACA9|\uC784\uC6D0\uBCF4\uACE0)'


    $isPpt =
        [regex]::IsMatch(
            $Request,
            $pptPattern
        )


    $visualIntent =
        [regex]::IsMatch(
            $Request,
            $visualPattern
        )


        # JUNGLE_SMART_REFLOW_ROUTE_V1
        #
        # Only strong explicit REFLOW intent enters PPT_REFLOW.
        # Ordinary visual requests remain PPT_SAFE.

        $reflowIntent =
            [regex]::IsMatch(
                $Request,
                '(?i)(?:\breflow\b|\uB9AC\uD50C\uB85C\uC6B0|\uB808\uC774\uC544\uC6C3.{0,30}\uC7AC\uAD6C\uC131|\uBC30\uCE58.{0,30}\uC7AC\uAD6C\uC131|\uB514\uC790\uC778.{0,30}\uC7AC\uAD6C\uC131|\uC801\uADF9\uC801.{0,30}\uC7AC\uBC30\uCE58)'
            )


        if ($isPpt -and $reflowIntent) {

            return [pscustomobject]@{

                route =
                    "PPT_REFLOW"

                provider =
                    "CLAUDE"

                mode =
                    "REFLOW"

                reason =
                    "Explicit PPT REFLOW intent detected."
            }
        }

    if ($isPpt -and $visualIntent) {

        return [pscustomobject]@{

            route =
                "PPT_SAFE"

            provider =
                "CLAUDE"

            mode =
                "SAFE"

            reason =
                "PPT visual intent detected."
        }
    }


    return [pscustomobject]@{

        route =
            "GENERAL_AI"

        provider =
            "AUTO"

        mode =
            "AUTO"

        reason =
            "No hard-coded specialist route required."
    }
}