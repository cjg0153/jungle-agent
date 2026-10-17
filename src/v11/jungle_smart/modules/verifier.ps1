function Test-JungleArtifact {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [int64]$MinimumBytes = 10000
    )

    if (-not (
        Test-Path `
            -LiteralPath $Path `
            -PathType Leaf
    )) {

        return $false
    }

    $item =
        Get-Item `
            -LiteralPath $Path

    if ($item.Length -lt $MinimumBytes) {

        return $false
    }

    return $true
}


function Test-JungleHardGate {

    param(
        [Parameter(Mandatory=$true)]
        [int]$Score,

        [Parameter(Mandatory=$true)]
        [int]$Critical,

        [Parameter(Mandatory=$true)]
        [int]$Major
    )

    return (
        $Score -ge 85 `
        -and
        $Critical -eq 0 `
        -and
        $Major -eq 0
    )
}