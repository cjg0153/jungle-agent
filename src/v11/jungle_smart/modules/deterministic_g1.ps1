# JUNGLE_SMART_DETERMINISTIC_G1_V1

function Get-JungleStringSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $sha =
        [System.Security.Cryptography.SHA256]::Create()

    try {

        $bytes =
            [System.Text.Encoding]::UTF8.GetBytes(
                $Text
            )

        $hash =
            $sha.ComputeHash(
                $bytes
            )

        return (
            [System.BitConverter]::ToString(
                $hash
            )
        ).Replace(
            "-",
            ""
        )
    }
    finally {

        $sha.Dispose()
    }
}


function Get-JunglePptZipSummary {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (-not (
        Test-Path `
            -LiteralPath $Path `
            -PathType Leaf
    )) {

        throw "PPTX not found: $Path"
    }


    Add-Type `
        -AssemblyName System.IO.Compression.FileSystem `
        -ErrorAction SilentlyContinue


    $zip =
        [System.IO.Compression.ZipFile]::OpenRead(
            $Path
        )


    try {

        $slides =
            @()


        foreach ($entry in $zip.Entries) {

            if (
                $entry.FullName -match
                '^ppt/slides/slide(\d+)\.xml$'
            ) {

                $slides +=
                    [pscustomobject]@{
                        Number = [int]$Matches[1]
                        Entry  = $entry
                    }
            }
        }


        $slides =
            @(
                $slides |
                Sort-Object Number
            )


        $textParts =
            New-Object System.Collections.Generic.List[string]


        $structureParts =
            New-Object System.Collections.Generic.List[string]


        foreach ($item in $slides) {

            $stream =
                $item.Entry.Open()


            $reader =
                New-Object `
                    System.IO.StreamReader `
                    -ArgumentList $stream


            try {

                $xml =
                    $reader.ReadToEnd()
            }
            finally {

                $reader.Dispose()
                $stream.Dispose()
            }


            $textParts.Add(
                "SLIDE:" + $item.Number
            )


            $textMatches =
                [regex]::Matches(
                    $xml,
                    '(?s)<a:t(?:\s[^>]*)?>(.*?)</a:t>'
                )


            foreach ($match in $textMatches) {

                $value =
                    [System.Net.WebUtility]::HtmlDecode(
                        $match.Groups[1].Value
                    )


                $textParts.Add(
                    $value
                )
            }


            $spCount =
                (
                    [regex]::Matches(
                        $xml,
                        '<p:sp(?=[\s>])'
                    )
                ).Count


            $picCount =
                (
                    [regex]::Matches(
                        $xml,
                        '<p:pic(?=[\s>])'
                    )
                ).Count


            $graphicCount =
                (
                    [regex]::Matches(
                        $xml,
                        '<p:graphicFrame(?=[\s>])'
                    )
                ).Count


            $connectorCount =
                (
                    [regex]::Matches(
                        $xml,
                        '<p:cxnSp(?=[\s>])'
                    )
                ).Count


            $groupCount =
                (
                    [regex]::Matches(
                        $xml,
                        '<p:grpSp(?=[\s>])'
                    )
                ).Count


            $structureParts.Add(
                (
                    "S{0}|SP={1}|PIC={2}|GF={3}|CX={4}|GRP={5}" -f
                    $item.Number,
                    $spCount,
                    $picCount,
                    $graphicCount,
                    $connectorCount,
                    $groupCount
                )
            )
        }


        $mediaCount =
            @(
                $zip.Entries |
                Where-Object {
                    $_.FullName -match '^ppt/media/'
                }
            ).Count


        return [pscustomobject]@{

            slide_count =
                $slides.Count

            text_hash =
                Get-JungleStringSha256 `
                    -Text ($textParts -join "`n")

            structure_hash =
                Get-JungleStringSha256 `
                    -Text ($structureParts -join "`n")

            media_count =
                $mediaCount
        }
    }
    finally {

        $zip.Dispose()
    }
}


function Test-JunglePptSafeG1 {

    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$ExpectedSourceHash
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


    $sourceHashNow =
        (
            Get-FileHash `
                -LiteralPath $SourcePath `
                -Algorithm SHA256
        ).Hash


    $sourcePreserved =
        (
            $sourceHashNow -eq
            $ExpectedSourceHash
        )


    $sourceSummary =
        Get-JunglePptZipSummary `
            -Path $SourcePath


    $artifactSummary =
        Get-JunglePptZipSummary `
            -Path $ArtifactPath


    $slideCountSame =
        (
            $sourceSummary.slide_count -eq
            $artifactSummary.slide_count
        )


    $textSame =
        (
            $sourceSummary.text_hash -eq
            $artifactSummary.text_hash
        )


    $structureSame =
        (
            $sourceSummary.structure_hash -eq
            $artifactSummary.structure_hash
        )


    $mediaCountSame =
        (
            $sourceSummary.media_count -eq
            $artifactSummary.media_count
        )


    $pass =
        (
            $sourcePreserved `
            -and
            $slideCountSame `
            -and
            $textSame `
            -and
            $structureSame `
            -and
            $mediaCountSame
        )


    return [pscustomobject]@{

        pass =
            $pass

        source_preserved =
            $sourcePreserved

        slide_count_same =
            $slideCountSame

        text_same =
            $textSame

        structure_same =
            $structureSame

        media_count_same =
            $mediaCountSame

        source_slide_count =
            $sourceSummary.slide_count

        artifact_slide_count =
            $artifactSummary.slide_count

        source_text_hash =
            $sourceSummary.text_hash

        artifact_text_hash =
            $artifactSummary.text_hash

        source_structure_hash =
            $sourceSummary.structure_hash

        artifact_structure_hash =
            $artifactSummary.structure_hash

        source_media_count =
            $sourceSummary.media_count

        artifact_media_count =
            $artifactSummary.media_count
    }
}