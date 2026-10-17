# JUNGLE_SMART_CHANGED_SCOPE_V1

function Get-JungleFileSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    return (
        Get-FileHash `
            -LiteralPath $Path `
            -Algorithm SHA256
    ).Hash
}


function Export-JunglePptSlideRenders {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory,

        [int]$Width = 1600,

        [int]$Height = 900
    )


    if (-not (
        Test-Path `
            -LiteralPath $Path `
            -PathType Leaf
    )) {

        throw "PPTX not found: $Path"
    }


    if (Test-Path -LiteralPath $OutputDirectory) {

        Remove-Item `
            -LiteralPath $OutputDirectory `
            -Recurse `
            -Force
    }


    New-Item `
        -ItemType Directory `
        -Path $OutputDirectory `
        -Force |
    Out-Null


    $ppt =
        $null

    $pres =
        $null


    try {

        $ppt =
            New-Object `
                -ComObject PowerPoint.Application


        $pres =
            $ppt.Presentations.Open(
                $Path,
                $true,
                $false,
                $false
            )


        $slideCount =
            [int]$pres.Slides.Count


        for (
            $slideNumber = 1;
            $slideNumber -le $slideCount;
            $slideNumber++
        ) {

            $slide =
                $pres.Slides.Item(
                    $slideNumber
                )


            try {

                $png =
                    Join-Path `
                        $OutputDirectory `
                        (
                            "slide_{0:D3}.png" -f
                            $slideNumber
                        )


                $slide.Export(
                    $png,
                    "PNG",
                    $Width,
                    $Height
                )


                if (-not (
                    Test-Path `
                        -LiteralPath $png `
                        -PathType Leaf
                )) {

                    throw "Slide render failed: $png"
                }
            }
            finally {

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $slide
                    ) | Out-Null
                }
                catch {}
            }
        }


        return $slideCount
    }
    finally {

        if ($null -ne $pres) {

            try {
                $pres.Close()
            }
            catch {}

            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $pres
                ) | Out-Null
            }
            catch {}
        }


        if ($null -ne $ppt) {

            try {
                $ppt.Quit()
            }
            catch {}

            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $ppt
                ) | Out-Null
            }
            catch {}
        }


        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


function Get-JungleChangedScope {

    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$WorkDirectory
    )


    if (-not (
        Test-Path `
            -LiteralPath $SourcePath `
            -PathType Leaf
    )) {

        throw "Source PPTX missing: $SourcePath"
    }


    if (-not (
        Test-Path `
            -LiteralPath $ArtifactPath `
            -PathType Leaf
    )) {

        throw "Artifact PPTX missing: $ArtifactPath"
    }


    $sourceHash =
        Get-JungleFileSha256 `
            -Path $SourcePath


    $artifactHash =
        Get-JungleFileSha256 `
            -Path $ArtifactPath


    # ------------------------------------------------------------
    # Cheapest possible path.
    #
    # If the two PPTX files are byte-identical,
    # there cannot be any changed visual slide.
    # Do not start PowerPoint COM.
    # ------------------------------------------------------------

    if ($sourceHash -eq $artifactHash) {

        return [pscustomobject]@{

            pass =
                $true

            byte_identical =
                $true

            rendered =
                $false

            slide_count_same =
                $true

            source_slide_count =
                $null

            artifact_slide_count =
                $null

            changed_count =
                0

            changed_slides =
                @()

            source_hash =
                $sourceHash

            artifact_hash =
                $artifactHash
        }
    }


    $beforeDir =
        Join-Path `
            $WorkDirectory `
            "before"


    $afterDir =
        Join-Path `
            $WorkDirectory `
            "after"


    New-Item `
        -ItemType Directory `
        -Path $WorkDirectory `
        -Force |
    Out-Null


    $sourceSlideCount =
        Export-JunglePptSlideRenders `
            -Path $SourcePath `
            -OutputDirectory $beforeDir


    $artifactSlideCount =
        Export-JunglePptSlideRenders `
            -Path $ArtifactPath `
            -OutputDirectory $afterDir


    $slideCountSame =
        (
            $sourceSlideCount -eq
            $artifactSlideCount
        )


    if (-not $slideCountSame) {

        return [pscustomobject]@{

            pass =
                $false

            byte_identical =
                $false

            rendered =
                $true

            slide_count_same =
                $false

            source_slide_count =
                $sourceSlideCount

            artifact_slide_count =
                $artifactSlideCount

            changed_count =
                $null

            changed_slides =
                @()

            source_hash =
                $sourceHash

            artifact_hash =
                $artifactHash
        }
    }


    $changed =
        New-Object System.Collections.Generic.List[int]


    for (
        $slideNumber = 1;
        $slideNumber -le $sourceSlideCount;
        $slideNumber++
    ) {

        $beforePng =
            Join-Path `
                $beforeDir `
                (
                    "slide_{0:D3}.png" -f
                    $slideNumber
                )


        $afterPng =
            Join-Path `
                $afterDir `
                (
                    "slide_{0:D3}.png" -f
                    $slideNumber
                )


        $beforeHash =
            Get-JungleFileSha256 `
                -Path $beforePng


        $afterHash =
            Get-JungleFileSha256 `
                -Path $afterPng


        if ($beforeHash -ne $afterHash) {

            $changed.Add(
                $slideNumber
            )
        }
    }


    return [pscustomobject]@{

        pass =
            $true

        byte_identical =
            $false

        rendered =
            $true

        slide_count_same =
            $true

        source_slide_count =
            $sourceSlideCount

        artifact_slide_count =
            $artifactSlideCount

        changed_count =
            $changed.Count

        changed_slides =
            @($changed)

        source_hash =
            $sourceHash

        artifact_hash =
            $artifactHash
    }
}