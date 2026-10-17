# JUNGLE_SMART_BLIND_PAIRWISE_QA_V2
# JUNGLE_PS51_NATIVE_ARRAY_FIX_V1

function Get-JungleTextSha256 {

    param(
        [Parameter(Mandatory=$true)]
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


function Get-JungleJsonFromClaudeOutput {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Raw
    )


    $text =
        $Raw.Trim()


    $fence =
        [regex]::Match(
            $text,
            '(?s)```(?:json)?\s*(.*?)\s*```'
        )


    if ($fence.Success) {

        $text =
            $fence.Groups[1].Value.Trim()
    }


    $first =
        $text.IndexOf(
            "{"
        )


    $last =
        $text.LastIndexOf(
            "}"
        )


    if (
        $first -lt 0 `
        -or
        $last -le $first
    ) {

        throw "Claude output contains no JSON object."
    }


    $jsonText =
        $text.Substring(
            $first,
            ($last - $first + 1)
        )


    try {

        return (
            $jsonText |
            ConvertFrom-Json
        )
    }
    catch {

        throw (
            "Claude Blind Pairwise JSON parse failed. " +
            $_.Exception.Message
        )
    }
}


function New-JungleBlindPairSet {

    param(
        [Parameter(Mandatory=$true)]
        [int[]]$ChangedSlides,

        [Parameter(Mandatory=$true)]
        [string]$ScopeDirectory,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$PairDirectory
    )


    New-Item `
        -ItemType Directory `
        -Path $PairDirectory `
        -Force |
    Out-Null


    $artifactHash =
        (
            Get-FileHash `
                -LiteralPath $ArtifactPath `
                -Algorithm SHA256
        ).Hash


    # IMPORTANT:
    # Native PowerShell array.
    #
    # Do NOT use Generic.List[object] here.
    # Windows PowerShell 5.1 can throw
    # "Argument types do not match" when a Generic.List[object]
    # containing PSCustomObject is converted with @().
    $mapping =
        @()


    $normalizedSlides =
        @(
            $ChangedSlides |
            ForEach-Object {
                [int]$_
            } |
            Sort-Object -Unique
        )


    $index =
        0


    foreach ($slideNumber in $normalizedSlides) {

        $index++


        $before =
            Join-Path `
                (Join-Path $ScopeDirectory "before") `
                (
                    "slide_{0:D3}.png" -f
                    $slideNumber
                )


        $after =
            Join-Path `
                (Join-Path $ScopeDirectory "after") `
                (
                    "slide_{0:D3}.png" -f
                    $slideNumber
                )


        foreach ($image in @(
            $before,
            $after
        )) {

            if (-not (
                Test-Path `
                    -LiteralPath $image `
                    -PathType Leaf
            )) {

                throw (
                    "Blind Pairwise source image missing: " +
                    $image
                )
            }
        }


        $pairId =
            "P{0:D3}" -f
            $index


        $orientationSeed =
            Get-JungleTextSha256 `
                -Text (
                    "slide:" +
                    $slideNumber +
                    "|" +
                    $artifactHash
                )


        $firstByte =
            [Convert]::ToInt32(
                $orientationSeed.Substring(
                    0,
                    2
                ),
                16
            )


        $aIsBefore =
            (
                ($firstByte % 2) -eq 0
            )


        $aPath =
            Join-Path `
                $PairDirectory `
                ($pairId + "_A.png")


        $bPath =
            Join-Path `
                $PairDirectory `
                ($pairId + "_B.png")


        if ($aIsBefore) {

            Copy-Item `
                -LiteralPath $before `
                -Destination $aPath `
                -Force


            Copy-Item `
                -LiteralPath $after `
                -Destination $bPath `
                -Force


            $aRole =
                "BEFORE"

            $bRole =
                "AFTER"
        }
        else {

            Copy-Item `
                -LiteralPath $after `
                -Destination $aPath `
                -Force


            Copy-Item `
                -LiteralPath $before `
                -Destination $bPath `
                -Force


            $aRole =
                "AFTER"

            $bRole =
                "BEFORE"
        }


        $entry =
            [pscustomobject]@{

                pair_id =
                    $pairId

                slide =
                    [int]$slideNumber

                a_path =
                    $aPath

                b_path =
                    $bPath

                a_role =
                    $aRole

                b_role =
                    $bRole
            }


        $mapping +=
            $entry
    }


    return $mapping
}


function Write-JungleBlindPairwisePrompt {

    param(
        [Parameter(Mandatory=$true)]
        [object[]]$Mapping,

        [Parameter(Mandatory=$true)]
        [string]$PromptPath
    )


    $lines =
        @()


    $lines +=
        "You are an independent visual QA reviewer."

    $lines +=
        ""

    $lines +=
        "You are given blind A/B pairs of PowerPoint slide renders."

    $lines +=
        "You do NOT know which image is before or after."

    $lines +=
        "Do not infer chronology from filenames or assume one side should win."

    $lines +=
        ""

    $lines +=
        "For every pair, inspect BOTH image files directly."

    $lines +=
        "Evaluate only visual quality and presentation safety:"

    $lines +=
        "- clipping or cutoff"

    $lines +=
        "- unintended overlap"

    $lines +=
        "- alignment"

    $lines +=
        "- spacing"

    $lines +=
        "- hierarchy"

    $lines +=
        "- readability"

    $lines +=
        "- balance"

    $lines +=
        "- obvious visual regression"

    $lines +=
        ""

    $lines +=
        "Do not rewrite slide content."

    $lines +=
        "Judge A versus B only."

    $lines +=
        ""

    $lines +=
        "winner must be exactly one of: A, B, SAME"

    $lines +=
        "severity must be exactly one of: none, minor, major, critical"

    $lines +=
        "confidence must be an integer from 0 to 100."

    $lines +=
        ""

    $lines +=
        "Return JSON only. No markdown. No commentary outside JSON."

    $lines +=
        ""

    $lines +=
        "Required schema:"

    $lines +=
        "{"

    $lines +=
        '  "protocol": "JUNGLE_SMART_BLIND_PAIRWISE_QA_V1",'

    $lines +=
        '  "pairs": ['

    $lines +=
        '    {'

    $lines +=
        '      "pair_id": "P001",'

    $lines +=
        '      "winner": "A",'

    $lines +=
        '      "severity": "minor",'

    $lines +=
        '      "confidence": 90,'

    $lines +=
        '      "reason": "short visual reason"'

    $lines +=
        '    }'

    $lines +=
        '  ]'

    $lines +=
        "}"

    $lines +=
        ""

    $lines +=
        "PAIR FILES:"

    $lines +=
        ""


    foreach ($item in $Mapping) {

        $lines +=
            (
                [string]$item.pair_id +
                ":"
            )


        $lines +=
            (
                "A = " +
                [string]$item.a_path
            )


        $lines +=
            (
                "B = " +
                [string]$item.b_path
            )


        $lines +=
            ""
    }


    $prompt =
        $lines -join
        [Environment]::NewLine


    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )


    [System.IO.File]::WriteAllText(
        $PromptPath,
        $prompt,
        $encoding
    )
}


function Invoke-JungleBlindPairwiseVisualQa {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Root,

        [Parameter(Mandatory=$true)]
        [int[]]$ChangedSlides,

        [Parameter(Mandatory=$true)]
        [string]$ScopeDirectory,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$RunDirectory,

        [int]$TimeoutSeconds = 600,

        [int]$PollMilliseconds = 200
    )


    $normalizedChangedSlides =
        @(
            $ChangedSlides |
            ForEach-Object {
                [int]$_
            } |
            Sort-Object -Unique
        )


    if ($normalizedChangedSlides.Count -eq 0) {

        return [pscustomobject]@{

            pass =
                $true

            skipped =
                $true

            claude_calls =
                0

            before_better_count =
                0

            critical_regressions =
                0

            major_regressions =
                0

            minor_regressions =
                0

            results =
                @()
        }
    }


    New-Item `
        -ItemType Directory `
        -Path $RunDirectory `
        -Force |
    Out-Null


    $pairDirectory =
        Join-Path `
            $RunDirectory `
            "pairs"


    $mapping =
        @(
            New-JungleBlindPairSet `
                -ChangedSlides ([int[]]$normalizedChangedSlides) `
                -ScopeDirectory $ScopeDirectory `
                -ArtifactPath $ArtifactPath `
                -PairDirectory $pairDirectory
        )


    if (
        $mapping.Count -ne
        $normalizedChangedSlides.Count
    ) {

        throw (
            "Blind Pair mapping count mismatch. Expected " +
            $normalizedChangedSlides.Count +
            " but created " +
            $mapping.Count
        )
    }


    $mapRecord =
        [ordered]@{

            protocol =
                "JUNGLE_SMART_BLIND_PAIR_MAP_V1"

            pairs =
                $mapping
        }


    $mapJson =
        $mapRecord |
        ConvertTo-Json -Depth 6


    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )


    $mapPath =
        Join-Path `
            $RunDirectory `
            "pair_map.json"


    [System.IO.File]::WriteAllText(
        $mapPath,
        $mapJson,
        $encoding
    )


    $promptPath =
        Join-Path `
            $RunDirectory `
            "prompt.txt"


    Write-JungleBlindPairwisePrompt `
        -Mapping ([object[]]$mapping) `
        -PromptPath $promptPath


    $rawResultPath =
        Join-Path `
            $RunDirectory `
            "claude_raw.txt"


    $worker =
        Join-Path `
            $Root `
            "executors\visual_qa_worker.ps1"


    if (-not (
        Test-Path `
            -LiteralPath $worker `
            -PathType Leaf
    )) {

        throw "Visual QA worker missing: $worker"
    }


    $powershellExe =
        Join-Path `
            $PSHOME `
            "powershell.exe"


    $processDirectory =
        Join-Path `
            $RunDirectory `
            "process"


    $processResult =
        Invoke-JungleProcessWithWatchdog `
            -FilePath $powershellExe `
            -Arguments @(
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                $worker,
                "-PromptPath",
                $promptPath,
                "-ResultPath",
                $rawResultPath
            ) `
            -TimeoutSeconds $TimeoutSeconds `
            -WorkDirectory $processDirectory `
            -PollMilliseconds $PollMilliseconds


    if ($processResult.timed_out) {

        throw (
            "Claude Blind Pairwise QA timed out and was terminated."
        )
    }


    if (-not $processResult.completed) {

        throw (
            "Claude Blind Pairwise QA worker did not complete."
        )
    }


    if ($processResult.exit_code -ne 0) {

        $stderr =
            ""


        if (
            Test-Path `
                -LiteralPath $processResult.stderr_path `
                -PathType Leaf
        ) {

            $stderr =
                [System.IO.File]::ReadAllText(
                    $processResult.stderr_path
                )
        }


        throw (
            "Claude Blind Pairwise QA worker failed. " +
            $stderr
        )
    }


    if (-not (
        Test-Path `
            -LiteralPath $rawResultPath `
            -PathType Leaf
    )) {

        throw "Claude Blind Pairwise raw result missing."
    }


    $raw =
        [System.IO.File]::ReadAllText(
            $rawResultPath
        )


    if ([string]::IsNullOrWhiteSpace(
        $raw
    )) {

        throw "Claude Blind Pairwise returned empty output."
    }


    $parsed =
        Get-JungleJsonFromClaudeOutput `
            -Raw $raw


    if (
        [string]$parsed.protocol -ne
        "JUNGLE_SMART_BLIND_PAIRWISE_QA_V1"
    ) {

        throw (
            "Blind Pairwise protocol mismatch: " +
            [string]$parsed.protocol
        )
    }


    $results =
        @(
            $parsed.pairs
        )


    if (
        $results.Count -ne
        $mapping.Count
    ) {

        throw (
            "Blind Pairwise result count mismatch. Expected " +
            $mapping.Count +
            " but received " +
            $results.Count
        )
    }


    # Native PowerShell array.
    # Avoid Generic.List[object] on Windows PowerShell 5.1.
    $mappedResults =
        @()


    foreach ($mapItem in $mapping) {

        $matches =
            @(
                $results |
                Where-Object {
                    [string]$_.pair_id -eq
                    [string]$mapItem.pair_id
                }
            )


        if ($matches.Count -ne 1) {

            throw (
                "Expected exactly one result for " +
                [string]$mapItem.pair_id
            )
        }


        $result =
            $matches[0]


        $winner =
            (
                [string]$result.winner
            ).ToUpperInvariant()


        if (
            $winner -notin
            @(
                "A",
                "B",
                "SAME"
            )
        ) {

            throw (
                "Invalid pair winner for " +
                [string]$mapItem.pair_id +
                ": " +
                $winner
            )
        }


        $severity =
            (
                [string]$result.severity
            ).ToLowerInvariant()


        if (
            $severity -notin
            @(
                "none",
                "minor",
                "major",
                "critical"
            )
        ) {

            throw (
                "Invalid severity for " +
                [string]$mapItem.pair_id +
                ": " +
                $severity
            )
        }


        $confidence =
            [int]$result.confidence


        if (
            $confidence -lt 0 `
            -or
            $confidence -gt 100
        ) {

            throw (
                "Invalid confidence for " +
                [string]$mapItem.pair_id
            )
        }


        if ($winner -eq "SAME") {

            $comparison =
                "SAME"
        }
        else {

            if ($winner -eq "A") {

                $winningRole =
                    [string]$mapItem.a_role
            }
            else {

                $winningRole =
                    [string]$mapItem.b_role
            }


            if ($winningRole -eq "AFTER") {

                $comparison =
                    "AFTER_BETTER"
            }
            else {

                $comparison =
                    "BEFORE_BETTER"
            }
        }


        $mappedEntry =
            [pscustomobject]@{

                pair_id =
                    [string]$mapItem.pair_id

                slide =
                    [int]$mapItem.slide

                winner =
                    $winner

                comparison =
                    $comparison

                severity =
                    $severity

                confidence =
                    $confidence

                reason =
                    [string]$result.reason
            }


        $mappedResults +=
            $mappedEntry
    }


    $beforeBetter =
        @(
            $mappedResults |
            Where-Object {
                $_.comparison -eq
                "BEFORE_BETTER"
            }
        )


    $critical =
        @(
            $beforeBetter |
            Where-Object {
                $_.severity -eq
                "critical"
            }
        ).Count


    $major =
        @(
            $beforeBetter |
            Where-Object {
                $_.severity -eq
                "major"
            }
        ).Count


    $minor =
        @(
            $beforeBetter |
            Where-Object {
                $_.severity -eq
                "minor"
            }
        ).Count


    $pass =
        (
            $beforeBetter.Count -eq 0 `
            -and
            $critical -eq 0 `
            -and
            $major -eq 0
        )


    $finalRecord =
        [ordered]@{

            protocol =
                "JUNGLE_SMART_BLIND_PAIRWISE_RESULT_V1"

            pass =
                $pass

            claude_calls =
                1

            before_better_count =
                $beforeBetter.Count

            critical_regressions =
                $critical

            major_regressions =
                $major

            minor_regressions =
                $minor

            results =
                $mappedResults

            created_at =
                (Get-Date).ToString("s")
        }


    $finalJson =
        $finalRecord |
        ConvertTo-Json -Depth 8


    $finalResultPath =
        Join-Path `
            $RunDirectory `
            "visual_qa_result.json"


    [System.IO.File]::WriteAllText(
        $finalResultPath,
        $finalJson,
        $encoding
    )


    return [pscustomobject]@{

        pass =
            $pass

        skipped =
            $false

        claude_calls =
            1

        before_better_count =
            $beforeBetter.Count

        critical_regressions =
            $critical

        major_regressions =
            $major

        minor_regressions =
            $minor

        results =
            $mappedResults

        pair_map_path =
            $mapPath

        raw_result_path =
            $rawResultPath

        final_result_path =
            $finalResultPath
    }
}