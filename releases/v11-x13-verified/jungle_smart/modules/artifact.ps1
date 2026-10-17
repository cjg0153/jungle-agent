function Get-JungleFileHash {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (-not (
        Test-Path `
            -LiteralPath $Path `
            -PathType Leaf
    )) {

        throw "File not found: $Path"
    }

    return (
        Get-FileHash `
            -LiteralPath $Path `
            -Algorithm SHA256
    ).Hash
}


function Test-JungleSourcePreserved {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [string]$OriginalHash
    )

    $currentHash =
        Get-JungleFileHash `
            -Path $Path

    return (
        $currentHash -eq
        $OriginalHash
    )
}

# ================================================================
# JUNGLE_SMART_NATIVE_ARTIFACT_RESOLUTION_V2
#
# Philosophy:
# Human-readable logs are useful hints, not artifact identity.
#
# Priority:
# 1. Parse a valid existing PPTX from the bridge log.
# 2. Fall back to pre/post path snapshot difference.
#
# This mirrors the proven v3.2 Artifact Discovery V3 principle.
# ================================================================

function Get-JungleNativePptArtifactSnapshot {

    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory
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
            -LiteralPath $OutputDirectory `
            -PathType Container
    )) {

        throw "Output directory missing: $OutputDirectory"
    }


    $baseName =
        [System.IO.Path]::GetFileNameWithoutExtension(
            $SourcePath
        )


    $prefix =
        $baseName +
        "_visual_v1_"


    $paths =
        @(
            Get-ChildItem `
                -LiteralPath $OutputDirectory `
                -Filter "*.pptx" `
                -File `
                -ErrorAction SilentlyContinue |
            Where-Object {

                $_.Name.StartsWith(
                    $prefix,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            } |
            ForEach-Object {
                $_.FullName
            }
        )


    return [string[]]$paths
}


function Resolve-JungleNativePptArtifact {

    param(
        [Parameter(Mandatory=$true)]
        [string]$BridgeLog,

        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory,

        [string[]]$ExistingPaths = @()
    )


    # ------------------------------------------------------------
    # Method 1
    # Parse a path from human-readable output.
    #
    # Be deliberately flexible about spacing/case.
    # ------------------------------------------------------------

    $logPattern =
        '(?im)\[NATIVE PPT\][^\r\n]*ARTIFACT[^\r\n]*?[:=]\s*["'']?(?<p>[A-Z]:\\[^\r\n"'']+?\.pptx)["'']?\s*$'


    $logMatches =
        @(
            [regex]::Matches(
                $BridgeLog,
                $logPattern
            )
        )


    for (
        $i = $logMatches.Count - 1;
        $i -ge 0;
        $i--
    ) {

        $candidate =
            [string]$logMatches[$i].Groups["p"].Value


        $candidate =
            $candidate.Trim()


        if (
            Test-Path `
                -LiteralPath $candidate `
                -PathType Leaf
        ) {

            return [pscustomobject]@{

                path =
                    $candidate

                method =
                    "LOG_PATH"

                new_path_count =
                    $null
            }
        }
    }


    # ------------------------------------------------------------
    # Method 2
    # Proven Artifact Discovery V3 principle:
    #
    # Compare path inventory before and after execution.
    #
    # Do NOT depend on LastWriteTime.
    # Do NOT depend on file-byte changes.
    # ------------------------------------------------------------

    $afterPaths =
        @(
            Get-JungleNativePptArtifactSnapshot `
                -SourcePath $SourcePath `
                -OutputDirectory $OutputDirectory
        )


    $newPaths =
        @(
            $afterPaths |
            Where-Object {
                $_ -notin $ExistingPaths
            }
        )


    if ($newPaths.Count -eq 1) {

        return [pscustomobject]@{

            path =
                [string]$newPaths[0]

            method =
                "NEW_PATH_V3"

            new_path_count =
                1
        }
    }


    if ($newPaths.Count -gt 1) {

        $sorted =
            @(
                $newPaths |
                Sort-Object -Descending
            )


        throw (
            "Native artifact resolution is ambiguous. " +
            "Multiple new PPTX paths found: " +
            ($sorted -join " | ")
        )
    }


    throw (
        "Jungle Smart could not resolve native PPT artifact. " +
        "Log parse failed and no new matching PPTX path was found."
    )
}