param(
    [Parameter(Mandatory=$true)]
    [string]$Request
)

$ErrorActionPreference = "Stop"

# ================================================================
# JUNGLE_V11_GENERAL_SOURCE_PROTECTION_V1
# JUNGLE_V11_ARTIFACT_CONTRACT_V1
#
# Contract:
#   Source        : READ ONLY
#   Immutable copy: Jungle controlled
#   Working copy : AI may modify
#   Final output : C:\AI\work\output
#   Source change: RESTORE + FAIL
# ================================================================

$runner =
    "C:\AI\run_verified_text_v3_fast.ps1"

$expectedRunnerHash =
    "A74D160CA96249143620970ED25014C1B302034C1248759266A50742431C36A0"

$outputRoot =
    "C:\AI\work\output"

$tempRoot =
    "C:\AI\work\temp"


function Get-JungleGeneralFiles {

    param(
        [string]$Directory,
        [string]$Extension
    )

    if (-not (
        Test-Path `
            -LiteralPath $Directory `
            -PathType Container
    )) {
        return @()
    }

    return @(
        Get-ChildItem `
            -LiteralPath $Directory `
            -File `
            -Recurse `
            -Filter ("*" + $Extension) `
            -ErrorAction SilentlyContinue |
        Where-Object {

            $name =
                $_.Name

            -not [regex]::IsMatch(
                $name,
                '^(Jungle_verified_fast_|fast_request_|fast_draft_|fast_claude_instruction_|fast_revision_instruction_|fast_qa_|fast_reqa_)',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        } |
        ForEach-Object {
            $_.FullName
        }
    )
}


function Restore-JungleSource {

    param(
        [string]$BackupPath,
        [string]$SourcePath,
        [string]$ExpectedHash
    )

    [System.IO.File]::Copy(
        $BackupPath,
        $SourcePath,
        $true
    )

    $restoredHash =
        (
            Get-FileHash `
                -LiteralPath $SourcePath `
                -Algorithm SHA256
        ).Hash

    if ($restoredHash -ne $ExpectedHash) {
        throw "CRITICAL: source restore verification failed."
    }
}


if (-not (
    Test-Path `
        -LiteralPath $runner `
        -PathType Leaf
)) {
    throw "FAST runner missing."
}


$runnerHash =
    (
        Get-FileHash `
            -LiteralPath $runner `
            -Algorithm SHA256
    ).Hash


if ($runnerHash -ne $expectedRunnerHash) {
    throw "FAST runner hash mismatch."
}


$pathPattern =
    '(?i)(?:"(?<p>[A-Z]:\\[^"]+?\.(?:html?|md|txt|json|js|ts|py|ps1|css|csv|xml|ya?ml))"|''(?<p>[A-Z]:\\[^'']+?\.(?:html?|md|txt|json|js|ts|py|ps1|css|csv|xml|ya?ml))''|(?<p>[A-Z]:\\.*?\.(?:html?|md|txt|json|js|ts|py|ps1|css|csv|xml|ya?ml))(?=\s|$))'


$sourceMatch =
    [regex]::Match(
        $Request,
        $pathPattern
    )


if (-not $sourceMatch.Success) {
    throw "Jungle v1.1 could not resolve protected source file."
}


$sourcePath =
    [string]$sourceMatch.Groups[
        "p"
    ].Value


if (-not (
    Test-Path `
        -LiteralPath $sourcePath `
        -PathType Leaf
)) {
    throw (
        "Source file missing: " +
        $sourcePath
    )
}


$sourceInfo =
    Get-Item `
        -LiteralPath $sourcePath


$extension =
    [System.IO.Path]::GetExtension(
        $sourcePath
    ).ToLowerInvariant()


$sourceStem =
    [System.IO.Path]::GetFileNameWithoutExtension(
        $sourcePath
    )


$sourceHashBefore =
    (
        Get-FileHash `
            -LiteralPath $sourcePath `
            -Algorithm SHA256
    ).Hash


$stamp =
    Get-Date -Format "yyyyMMdd_HHmmss_fff"


$runDir =
    Join-Path `
        $tempRoot `
        (
            "general_modify_v11_" +
            $stamp +
            "_" +
            [guid]::NewGuid().ToString("N").Substring(0,8)
        )


New-Item `
    -ItemType Directory `
    -Path $runDir `
    -Force |
Out-Null


New-Item `
    -ItemType Directory `
    -Path $outputRoot `
    -Force |
Out-Null


$backupPath =
    Join-Path `
        $runDir `
        (
            "SOURCE_BACKUP" +
            $extension
        )


$workingPath =
    Join-Path `
        $runDir `
        (
            "WORKING_" +
            $sourceInfo.Name
        )


# ---------------------------------------------------------------
# Jungle itself performs both copies.
# AI does not copy the source.
# ---------------------------------------------------------------

[System.IO.File]::Copy(
    $sourcePath,
    $backupPath,
    $false
)


[System.IO.File]::Copy(
    $sourcePath,
    $workingPath,
    $false
)


$backupHash =
    (
        Get-FileHash `
            -LiteralPath $backupPath `
            -Algorithm SHA256
    ).Hash


$workingHashBefore =
    (
        Get-FileHash `
            -LiteralPath $workingPath `
            -Algorithm SHA256
    ).Hash


if (
    $backupHash -ne $sourceHashBefore `
    -or
    $workingHashBefore -ne $sourceHashBefore
) {
    throw "Initial source-copy verification failed."
}


# ---------------------------------------------------------------
# Snapshot AFTER Jungle's own copies.
# Any later path is an AI/runner artifact candidate.
# ---------------------------------------------------------------

$outputBefore =
    @(
        Get-JungleGeneralFiles `
            -Directory $outputRoot `
            -Extension $extension
    )


$tempBefore =
    @(
        Get-JungleGeneralFiles `
            -Directory $tempRoot `
            -Extension $extension
    )


# ---------------------------------------------------------------
# Remove the original path from the AI request completely.
# Claude only sees WORKING copy.
# ---------------------------------------------------------------

$workingRequest =
    $Request.Replace(
        $sourcePath,
        $workingPath
    )


$protectedRequest =
    @"
[JUNGLE MANAGED WORKING COPY]

The file path in this request is a Jungle-managed WORKING COPY.

Mandatory rules:
- Modify only the working-copy path supplied in the request.
- Never search for or edit another copy of the source.
- Do not overwrite files outside C:\AI\work.
- If creating a separate revised artifact, create it inside C:\AI\work\temp.
- Do the actual requested modification, not only a prose recommendation.

[USER REQUEST]
$workingRequest
"@


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE v1.1 GENERAL FILE"
Write-Host " SOURCE PROTECTION ACTIVE"
Write-Host "========================================"
Write-Host ""

Write-Host "Source            :" $sourcePath
Write-Host "Working copy      :" $workingPath
Write-Host "Source SHA256     :" $sourceHashBefore
Write-Host ""
Write-Host "AI-visible source : WORKING COPY ONLY"
Write-Host ""


$runnerItems =
    @()


try {

    & $runner `
        -Request $protectedRequest *>&1 |
    ForEach-Object {

        $runnerItems +=
            $_

        Write-Host (
            [string]$_
        )
    }
}
catch {

    $sourceHashNow =
        (
            Get-FileHash `
                -LiteralPath $sourcePath `
                -Algorithm SHA256
        ).Hash


    if ($sourceHashNow -ne $sourceHashBefore) {

        Restore-JungleSource `
            -BackupPath $backupPath `
            -SourcePath $sourcePath `
            -ExpectedHash $sourceHashBefore
    }

    throw
}


$runnerText =
    (
        @(
            $runnerItems |
            ForEach-Object {
                [string]$_
            }
        ) -join "`r`n"
    )


# ---------------------------------------------------------------
# SOURCE INVARIANT 1
# ---------------------------------------------------------------

$sourceHashAfterAI =
    (
        Get-FileHash `
            -LiteralPath $sourcePath `
            -Algorithm SHA256
    ).Hash


if ($sourceHashAfterAI -ne $sourceHashBefore) {

    Restore-JungleSource `
        -BackupPath $backupPath `
        -SourcePath $sourcePath `
        -ExpectedHash $sourceHashBefore

    throw "SOURCE PROTECTION VIOLATION: original was changed and has been restored."
}


# ---------------------------------------------------------------
# Runner must have passed its quality gate.
# ---------------------------------------------------------------

$runnerPassed =
    [regex]::IsMatch(
        $runnerText,
        '(?im)^\s*STATUS\s*:\s*JUNGLE_FAST_VERIFIED_COMPLETE\s*$'
    )


if (-not $runnerPassed) {

    throw "FAST quality pipeline did not reach VERIFIED COMPLETE."
}


# ---------------------------------------------------------------
# Artifact discovery
#
# Priority:
#   1. Explicit path in runner output
#   2. New output file
#   3. New temp file
#   4. Modified working copy
# ---------------------------------------------------------------

$candidatePath =
    $null


$extEscaped =
    [regex]::Escape(
        $extension
    )


$explicitPattern =
    (
        '(?im)^\s*(?:Verified Artifact|OUTPUT_HTML)\s*:\s*' +
        '(?<p>[A-Z]:\\[^\r\n]+?' +
        $extEscaped +
        ')\s*$'
    )


$explicitMatch =
    [regex]::Match(
        $runnerText,
        $explicitPattern
    )


if ($explicitMatch.Success) {

    $explicitPath =
        [string]$explicitMatch.Groups[
            "p"
        ].Value.Trim()


    if (
        Test-Path `
            -LiteralPath $explicitPath `
            -PathType Leaf
    ) {

        if (
            [System.IO.Path]::GetExtension(
                $explicitPath
            ).ToLowerInvariant() -eq
            $extension
        ) {

            $candidatePath =
                $explicitPath
        }
    }
}


$outputAfter =
    @(
        Get-JungleGeneralFiles `
            -Directory $outputRoot `
            -Extension $extension
    )


$tempAfter =
    @(
        Get-JungleGeneralFiles `
            -Directory $tempRoot `
            -Extension $extension
    )


if ($null -eq $candidatePath) {

    $newOutput =
        @(
            $outputAfter |
            Where-Object {
                $outputBefore -notcontains $_
            }
        )


    if ($newOutput.Count -eq 1) {

        $candidatePath =
            $newOutput[0]
    }
    elseif ($newOutput.Count -gt 1) {

        throw (
            "Artifact resolution ambiguous: multiple new output files. " +
            ($newOutput -join " | ")
        )
    }
}


if ($null -eq $candidatePath) {

    $newTemp =
        @(
            $tempAfter |
            Where-Object {

                $tempBefore -notcontains $_ `
                -and
                $_ -ne $workingPath `
                -and
                $_ -ne $backupPath
            }
        )


    if ($newTemp.Count -eq 1) {

        $candidatePath =
            $newTemp[0]
    }
    elseif ($newTemp.Count -gt 1) {

        throw (
            "Artifact resolution ambiguous: multiple new temp files. " +
            ($newTemp -join " | ")
        )
    }
}


$workingHashAfter =
    (
        Get-FileHash `
            -LiteralPath $workingPath `
            -Algorithm SHA256
    ).Hash


if (
    $null -eq $candidatePath `
    -and
    $workingHashAfter -ne $workingHashBefore
) {

    $candidatePath =
        $workingPath
}


if ($null -eq $candidatePath) {

    throw (
        "ARTIFACT CONTRACT FAILED: " +
        "quality pipeline passed but no modified artifact was found."
    )
}


if (
    [string]::Equals(
        [System.IO.Path]::GetFullPath(
            $candidatePath
        ),
        [System.IO.Path]::GetFullPath(
            $sourcePath
        ),
        [System.StringComparison]::OrdinalIgnoreCase
    )
) {

    throw "ARTIFACT CONTRACT FAILED: original source cannot be the final artifact."
}


if (-not (
    Test-Path `
        -LiteralPath $candidatePath `
        -PathType Leaf
)) {

    throw "Artifact candidate does not exist."
}


$candidateInfo =
    Get-Item `
        -LiteralPath $candidatePath


if ($candidateInfo.Length -lt 1) {
    throw "Artifact candidate is empty."
}


# ---------------------------------------------------------------
# Controlled promotion
# ---------------------------------------------------------------

$finalName =
    (
        $sourceStem +
        "_jungle_v11_" +
        (Get-Date -Format "yyyyMMdd_HHmmss") +
        $extension
    )


$finalPath =
    Join-Path `
        $outputRoot `
        $finalName


if (
    Test-Path `
        -LiteralPath $finalPath
) {

    $finalPath =
        Join-Path `
            $outputRoot `
            (
                $sourceStem +
                "_jungle_v11_" +
                (Get-Date -Format "yyyyMMdd_HHmmss_fff") +
                "_" +
                [guid]::NewGuid().ToString("N").Substring(0,6) +
                $extension
            )
}


[System.IO.File]::Copy(
    $candidatePath,
    $finalPath,
    $false
)


$finalInfo =
    Get-Item `
        -LiteralPath $finalPath


if ($finalInfo.Length -lt 1) {
    throw "Published artifact is empty."
}


$finalHash =
    (
        Get-FileHash `
            -LiteralPath $finalPath `
            -Algorithm SHA256
    ).Hash


# ---------------------------------------------------------------
# SOURCE INVARIANT 2
# ---------------------------------------------------------------

$sourceHashFinal =
    (
        Get-FileHash `
            -LiteralPath $sourcePath `
            -Algorithm SHA256
    ).Hash


if ($sourceHashFinal -ne $sourceHashBefore) {

    Remove-Item `
        -LiteralPath $finalPath `
        -Force `
        -ErrorAction SilentlyContinue

    Restore-JungleSource `
        -BackupPath $backupPath `
        -SourcePath $sourcePath `
        -ExpectedHash $sourceHashBefore

    throw "SOURCE PROTECTION VIOLATION during final promotion."
}


# ---------------------------------------------------------------
# Audit result
# ---------------------------------------------------------------

$resultPath =
    Join-Path `
        $runDir `
        "general_modify_result.json"


$result =
    [ordered]@{
        status =
            "JUNGLE_V11_GENERAL_FILE_VERIFIED_COMPLETE"

        source_path =
            $sourcePath

        source_hash_before =
            $sourceHashBefore

        source_hash_after =
            $sourceHashFinal

        source_unchanged =
            $true

        working_path =
            $workingPath

        candidate_path =
            $candidatePath

        artifact_path =
            $finalPath

        artifact_sha256 =
            $finalHash

        artifact_size =
            $finalInfo.Length

        runner_verified =
            $true

        run_directory =
            $runDir
    }


$json =
    $result |
    ConvertTo-Json -Depth 8


$utf8Bom =
    New-Object System.Text.UTF8Encoding(
        $true
    )


[System.IO.File]::WriteAllText(
    $resultPath,
    $json,
    $utf8Bom
)


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE v1.1 VERIFIED ARTIFACT"
Write-Host "========================================"
Write-Host ""

Write-Host "Source unchanged  : True"
Write-Host "Source SHA256     :" $sourceHashFinal
Write-Host "Verified Artifact :" $finalPath
Write-Host "Artifact SHA256   :" $finalHash
Write-Host "Artifact Size     :" $finalInfo.Length "bytes"
Write-Host "Audit             :" $resultPath
Write-Host ""
Write-Host "STATUS : JUNGLE_V11_GENERAL_FILE_VERIFIED_COMPLETE"
Write-Host "========================================"