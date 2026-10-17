param(
    [Parameter(Mandatory=$true)][string]$Request,
    [int]$ClaudeTimeoutSeconds = 600,
    [int]$CodexTimeoutSeconds = 600,
    [int]$FastPassScore = 85
)

$ErrorActionPreference = "Stop"

$claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
if ($null -ne $claudeCmd) { $claudePath = $claudeCmd.Source }
elseif (Test-Path -LiteralPath "C:\Users\Zephyrus G14\.local\bin\claude.exe" -PathType Leaf) {
    $claudePath = "C:\Users\Zephyrus G14\.local\bin\claude.exe"
}
else { throw "Claude Code를 찾을 수 없습니다." }

$codexPath = "C:\Users\Zephyrus G14\AppData\Local\Programs\OpenAI\Codex\bin\codex.exe"
if (!(Test-Path -LiteralPath $codexPath -PathType Leaf)) {
    $cc = Get-Command codex.cmd -CommandType Application -ErrorAction SilentlyContinue
    if ($null -eq $cc) { throw "Codex CLI를 찾을 수 없습니다." }
    $codexPath = $cc.Source
}

$catalogPath = "C:\AI\config\jungle_skill_catalog.json"
if (!(Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
    & "C:\AI\scripts\build_jungle_skill_catalog_v1.ps1"
}
$catalog = Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8 | ConvertFrom-Json

$claudeCatalog = (@($catalog.items | Where-Object { $_.kind -like "CLAUDE_*" } | ForEach-Object {
    "$($_.kind) $($_.name) :: $($_.description)"
}) -join "`r`n")

$codexCatalog = (@($catalog.items | Where-Object { $_.kind -eq "CODEX_SKILL" } | ForEach-Object {
    "CODEX_SKILL $($_.name) :: $($_.description)"
}) -join "`r`n")

$drawio = ($null -ne (Get-Command drawio -ErrorAction SilentlyContinue))
$mmdc = ($null -ne (Get-Command mmdc -ErrorAction SilentlyContinue))
$python =
    ($null -ne (Get-Command python -ErrorAction SilentlyContinue)) -or
    ($null -ne (Get-Command python3 -ErrorAction SilentlyContinue))

$tempDir = "C:\AI\work\temp"
$outputDir = "C:\AI\work\output"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$requestPath = Join-Path $tempDir "fast_request_${stamp}.txt"
$claudeInstructionPath = Join-Path $tempDir "fast_claude_instruction_${stamp}.txt"
$draftPath = Join-Path $tempDir "fast_draft_${stamp}.txt"
$qaPath = Join-Path $tempDir "fast_qa_${stamp}.json"
$revisionInstructionPath = Join-Path $tempDir "fast_revision_instruction_${stamp}.txt"
$revisionPath = Join-Path $tempDir "fast_revision_${stamp}.txt"
$reQaPath = Join-Path $tempDir "fast_reqa_${stamp}.json"
$finalPath = Join-Path $outputDir "Jungle_verified_fast_${stamp}.txt"

$encBom = New-Object System.Text.UTF8Encoding($true)
$encUtf8 = New-Object System.Text.UTF8Encoding($false)
# ============================================================
# JUNGLE_HTML_ARTIFACT_V1
# Verified HTML deliverables are promoted from temp to output
# only after Jungle quality gate passes.
# ============================================================
function Publish-JungleHtmlArtifact {
    param(
        [string]$HtmlPath,
        [string]$OutputDir,
        [string]$Stamp
    )

    if ([string]::IsNullOrWhiteSpace($HtmlPath)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $HtmlPath -PathType Leaf)) {
        return $null
    }

    $htmlStem =
        [System.IO.Path]::GetFileNameWithoutExtension($HtmlPath)

    $publishedPath =
        Join-Path `
            $OutputDir `
            ($htmlStem + "_verified_" + $Stamp + ".html")

    Copy-Item `
        -LiteralPath $HtmlPath `
        -Destination $publishedPath `
        -Force

    $publishedInfo =
        Get-Item -LiteralPath $publishedPath

    if ($publishedInfo.Length -lt 100) {
        throw "[HTML ARTIFACT] 생성된 HTML 파일 크기가 비정상적으로 작습니다: $($publishedInfo.Length) bytes"
    }

    return $publishedInfo
}

# END JUNGLE_HTML_ARTIFACT_V1
# ============================================================
# JUNGLE_VISUAL_QA_V1
# HTML embedded visual assets are extracted by Jungle itself.
# Claude only receives Read access to the extracted assets.
# No Bash / PowerShell execution permission is granted to Claude.
# ============================================================

function Get-JungleVisualEdgePath {

    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )

    foreach ($candidate in $candidates) {

        if (
            -not [string]::IsNullOrWhiteSpace($candidate) `
            -and (Test-Path -LiteralPath $candidate -PathType Leaf)
        ) {
            return $candidate
        }
    }

    return $null
}


function New-JungleHtmlVisualBundle {

    param(
        [string]$HtmlPath,
        [string]$WorkRoot,
        [string]$Stamp
    )

    $utf8 =
        New-Object System.Text.UTF8Encoding($false)

    $runDir =
        Join-Path $WorkRoot ("visual_" + $Stamp)

    $imageDir =
        Join-Path $runDir "images"

    $videoDir =
        Join-Path $runDir "videos"

    $frameDir =
        Join-Path $runDir "video_frames"

    $profileRoot =
        Join-Path $runDir "edge_profiles"

    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
    New-Item -ItemType Directory -Path $imageDir -Force | Out-Null
    New-Item -ItemType Directory -Path $videoDir -Force | Out-Null
    New-Item -ItemType Directory -Path $frameDir -Force | Out-Null
    New-Item -ItemType Directory -Path $profileRoot -Force | Out-Null

    $html =
        [System.IO.File]::ReadAllText($HtmlPath)

    $srcMatches =
        [regex]::Matches(
            $html,
            '(?is)\bsrc\s*=\s*["''](data:(?:image|video)/[^"'']+)["'']'
        )

    $imageFiles = @()
    $videoFiles = @()
    $frameFiles = @()

    $imageIndex = 0
    $videoIndex = 0

    foreach ($srcMatch in $srcMatches) {

        $src =
            $srcMatch.Groups[1].Value

        if (
            $src -match
            '^data:(image/[^;]+);base64,(.+)$'
        ) {

            $mime =
                $matches[1].ToLowerInvariant()

            $payload =
                $matches[2]

            $extension = $null

            switch ($mime) {

                "image/jpeg" {
                    $extension = "jpg"
                }

                "image/jpg" {
                    $extension = "jpg"
                }

                "image/png" {
                    $extension = "png"
                }

                "image/webp" {
                    $extension = "webp"
                }

                "image/gif" {
                    $extension = "gif"
                }

                "image/svg+xml" {
                    $extension = "svg"
                }
            }

            if (
                -not [string]::IsNullOrWhiteSpace($extension)
            ) {

                $imageIndex++

                try {

                    $bytes =
                        [System.Convert]::FromBase64String(
                            $payload
                        )

                    $imagePath =
                        Join-Path `
                            $imageDir `
                            (
                                "image_{0:D3}.{1}" `
                                -f $imageIndex,
                                   $extension
                            )

                    [System.IO.File]::WriteAllBytes(
                        $imagePath,
                        $bytes
                    )

                    $imageFiles += $imagePath
                }
                catch {
                    # malformed image is simply omitted;
                    # manifest counts will reveal incompleteness.
                }
            }
        }


        if (
            $src -match
            '^data:(video/[^;]+);base64,(.+)$'
        ) {

            $mime =
                $matches[1].ToLowerInvariant()

            $payload =
                $matches[2]

            if ($mime -eq "video/mp4") {

                $videoIndex++

                try {

                    $bytes =
                        [System.Convert]::FromBase64String(
                            $payload
                        )

                    $videoPath =
                        Join-Path `
                            $videoDir `
                            (
                                "video_{0:D2}.mp4" `
                                -f $videoIndex
                            )

                    [System.IO.File]::WriteAllBytes(
                        $videoPath,
                        $bytes
                    )

                    $videoFiles += $videoPath
                }
                catch {
                    # malformed video is omitted;
                    # QA will see missing frame coverage.
                }
            }
        }
    }


    # --------------------------------------------------------
    # Render six representative frames for each MP4
    # --------------------------------------------------------

    $edgePath =
        Get-JungleVisualEdgePath

    $samplePoints = @(
        @{ Name = "02"; Ratio = 0.02 },
        @{ Name = "20"; Ratio = 0.20 },
        @{ Name = "40"; Ratio = 0.40 },
        @{ Name = "60"; Ratio = 0.60 },
        @{ Name = "80"; Ratio = 0.80 },
        @{ Name = "95"; Ratio = 0.95 }
    )

    if (
        -not [string]::IsNullOrWhiteSpace($edgePath)
    ) {

        $videoNumber = 0

        foreach ($videoPath in $videoFiles) {

            $videoNumber++

            $videoBytes =
                [System.IO.File]::ReadAllBytes(
                    $videoPath
                )

            $videoBase64 =
                [System.Convert]::ToBase64String(
                    $videoBytes
                )

            foreach ($point in $samplePoints) {

                $name =
                    $point.Name

                $ratio =
                    $point.Ratio

                $wrapperPath =
                    Join-Path `
                        $frameDir `
                        (
                            "video_{0:D2}_frame_{1}.html" `
                            -f $videoNumber,
                               $name
                        )

                $pngPath =
                    Join-Path `
                        $frameDir `
                        (
                            "video_{0:D2}_frame_{1}.png" `
                            -f $videoNumber,
                               $name
                        )

                $profileDir =
                    Join-Path `
                        $profileRoot `
                        (
                            "video_{0:D2}_{1}" `
                            -f $videoNumber,
                               $name
                        )

                New-Item `
                    -ItemType Directory `
                    -Path $profileDir `
                    -Force |
                    Out-Null

                $frameHtml = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Jungle Video Frame</title>

<style>
html, body {
    margin: 0;
    padding: 0;
    width: 1280px;
    height: 760px;
    overflow: hidden;
    background: #111;
}

video {
    display: none;
}

canvas {
    display: block;
    width: 1280px;
    height: 760px;
}
</style>
</head>

<body>

<video id="v" muted preload="auto">
<source
    src="data:video/mp4;base64,$videoBase64"
    type="video/mp4">
</video>

<canvas
    id="c"
    width="1280"
    height="760">
</canvas>

<script>

const video =
    document.getElementById("v");

const canvas =
    document.getElementById("c");

const ctx =
    canvas.getContext("2d");

const ratio = $ratio;

function drawFrame() {

    const vw = video.videoWidth;
    const vh = video.videoHeight;

    ctx.fillStyle = "#111";
    ctx.fillRect(
        0,
        0,
        canvas.width,
        canvas.height
    );

    ctx.fillStyle = "#ffffff";
    ctx.font = "20px Arial";

    const label =
        "Jungle sample " +
        Math.round(ratio * 100) +
        "% | " +
        video.currentTime.toFixed(1) +
        " sec / " +
        video.duration.toFixed(1) +
        " sec";

    ctx.fillText(
        label,
        16,
        27
    );

    const top = 40;
    const availableW = 1280;
    const availableH = 720;

    let dw = availableW;
    let dh = dw * vh / vw;

    if (dh > availableH) {

        dh = availableH;
        dw = dh * vw / vh;
    }

    const dx =
        (availableW - dw) / 2;

    const dy =
        top +
        (availableH - dh) / 2;

    ctx.drawImage(
        video,
        dx,
        dy,
        dw,
        dh
    );

    document.title =
        "JUNGLE_FRAME_READY";
}

video.addEventListener(
    "loadedmetadata",
    function() {

        let target =
            video.duration * ratio;

        if (target >= video.duration) {

            target =
                Math.max(
                    0,
                    video.duration - 0.1
                );
        }

        video.currentTime = target;
    }
);

video.addEventListener(
    "seeked",
    function() {

        window.setTimeout(
            drawFrame,
            500
        );
    }
);

video.addEventListener(
    "error",
    function() {

        document.title =
            "VIDEO_ERROR";
    }
);

</script>

</body>
</html>
"@

                [System.IO.File]::WriteAllText(
                    $wrapperPath,
                    $frameHtml,
                    $utf8
                )

                if (
                    Test-Path -LiteralPath $pngPath
                ) {

                    Remove-Item `
                        -LiteralPath $pngPath `
                        -Force
                }

                $htmlUri =
                    (
                        New-Object System.Uri(
                            $wrapperPath
                        )
                    ).AbsoluteUri

                $psi =
                    New-Object System.Diagnostics.ProcessStartInfo

                $psi.FileName =
                    $edgePath

                $psi.Arguments = (
                    '--headless=new ' +
                    '--hide-scrollbars ' +
                    '--autoplay-policy=no-user-gesture-required ' +
                    '--run-all-compositor-stages-before-draw ' +
                    '--virtual-time-budget=8000 ' +
                    '--window-size=1280,760 ' +
                    '--user-data-dir="' +
                    $profileDir +
                    '" ' +
                    '--screenshot="' +
                    $pngPath +
                    '" ' +
                    '"' +
                    $htmlUri +
                    '"'
                )

                $psi.UseShellExecute =
                    $false

                $psi.CreateNoWindow =
                    $true

                $psi.RedirectStandardOutput =
                    $true

                $psi.RedirectStandardError =
                    $true

                $proc =
                    New-Object System.Diagnostics.Process

                $proc.StartInfo =
                    $psi

                try {

                    [void]$proc.Start()

                    $stdoutTask =
                        $proc.StandardOutput.ReadToEndAsync()

                    $stderrTask =
                        $proc.StandardError.ReadToEndAsync()

                    $exited =
                        $proc.WaitForExit(20000)

                    if (-not $exited) {

                        try {
                            $proc.Kill()
                        }
                        catch {}

                        continue
                    }

                    $proc.WaitForExit()

                    Start-Sleep `
                        -Milliseconds 300

                    if (
                        Test-Path `
                            -LiteralPath $pngPath `
                            -PathType Leaf
                    ) {

                        $frameInfo =
                            Get-Item `
                                -LiteralPath $pngPath

                        if ($frameInfo.Length -gt 1000) {

                            $frameFiles +=
                                $pngPath
                        }
                    }
                }
                catch {
                    # QA manifest will show incomplete frames.
                }
            }
        }
    }


    # --------------------------------------------------------
    # Manifest
    # --------------------------------------------------------

    $manifestPath =
        Join-Path `
            $runDir `
            "visual_manifest.txt"

    $expectedFrames =
        $videoFiles.Count * 6

    $manifest = @()

    $manifest +=
        "JUNGLE VISUAL QA MANIFEST V1"

    $manifest +=
        "SOURCE_HTML: $HtmlPath"

    $manifest +=
        "IMAGE_COUNT: $($imageFiles.Count)"

    $manifest +=
        "VIDEO_COUNT: $($videoFiles.Count)"

    $manifest +=
        "EXPECTED_VIDEO_FRAMES: $expectedFrames"

    $manifest +=
        "RENDERED_VIDEO_FRAMES: $($frameFiles.Count)"

    $manifest +=
        "EDGE_AVAILABLE: $(-not [string]::IsNullOrWhiteSpace($edgePath))"

    $manifest +=
        ""

    $manifest +=
        "[IMAGES]"

    $n = 0

    foreach ($imagePath in $imageFiles) {

        $n++

        $manifest +=
            (
                "{0:D3} | {1}" `
                -f $n,
                   $imagePath
            )
    }

    $manifest +=
        ""

    $manifest +=
        "[VIDEO FILES - DO NOT READ DIRECTLY WITH CLAUDE]"

    $n = 0

    foreach ($videoPath in $videoFiles) {

        $n++

        $manifest +=
            (
                "{0:D2} | {1}" `
                -f $n,
                   $videoPath
            )
    }

    $manifest +=
        ""

    $manifest +=
        "[VIDEO REPRESENTATIVE FRAMES]"

    $n = 0

    foreach ($framePath in $frameFiles) {

        $n++

        $manifest +=
            (
                "{0:D3} | {1}" `
                -f $n,
                   $framePath
            )
    }

    $manifest +=
        ""

    $manifest +=
        "[REVIEW RULES]"

    $manifest +=
        "- Inspect every file listed under [IMAGES]."

    $manifest +=
        "- Inspect every PNG listed under [VIDEO REPRESENTATIVE FRAMES]."

    $manifest +=
        "- Do not claim that MP4 audio was reviewed."

    $manifest +=
        "- Video frames are representative samples, not every frame."

    $manifest +=
        "- If RENDERED_VIDEO_FRAMES is lower than EXPECTED_VIDEO_FRAMES, visual review is incomplete."

    [System.IO.File]::WriteAllLines(
        $manifestPath,
        $manifest,
        $utf8
    )

    return [pscustomobject]@{

        RunDir =
            $runDir

        ManifestPath =
            $manifestPath

        ImageCount =
            $imageFiles.Count

        VideoCount =
            $videoFiles.Count

        ExpectedFrameCount =
            $expectedFrames

        FrameCount =
            $frameFiles.Count

        EdgeAvailable =
            (-not [string]::IsNullOrWhiteSpace($edgePath))
    }
}

# END JUNGLE_VISUAL_QA_V1
# ============================================================

# ============================================================

[System.IO.File]::WriteAllText($requestPath,$Request,$encBom)
# ============================================================
# JUNGLE_CLAUDE_ADDITIONAL_DIRS_V1
# Claude Code can access Jungle work files plus source-artifact
# directories explicitly referenced in the user's request.
# ============================================================
$claudeAccessDirs = @("C:\AI\work")

$sourcePathPattern = '(?i)([A-Z]:\\[^"\r\n]+?\.(?:html?|md|txt|json|csv|pdf|docx?|pptx?|xlsx?|hwp|hwpx|mp4|mov|avi|wav|mp3))'

$sourcePathMatches = [regex]::Matches(
    $Request,
    $sourcePathPattern
)

foreach ($sourcePathMatch in $sourcePathMatches) {

    $sourceCandidate = $sourcePathMatch.Groups[1].Value

    if (Test-Path -LiteralPath $sourceCandidate -PathType Leaf) {

        $sourceDirCandidate =
            [System.IO.Path]::GetDirectoryName($sourceCandidate)

        if (-not [string]::IsNullOrWhiteSpace($sourceDirCandidate)) {
            $claudeAccessDirs += $sourceDirCandidate
        }
    }
}

$claudeAccessDirs =
    @($claudeAccessDirs | Sort-Object -Unique)

$script:claudeAddDirArgs =
    '--add-dir ' +
    (
        (
            $claudeAccessDirs |
            ForEach-Object { '"' + $_ + '"' }
        ) -join ' '
    )

# END JUNGLE_CLAUDE_ADDITIONAL_DIRS_V1
# ============================================================


function Get-JungleStagePercent {
    param([string]$StageName)

    switch ($StageName) {
        "Claude draft"    { return 25 }
        "Codex QA"        { return 50 }
        "Claude revision" { return 75 }
        "Codex re-QA"     { return 90 }
        default           { return 0 }
    }
}

function Get-JungleProgressBar {
    param([int]$Percent)

    $filled = [Math]::Floor($Percent / 10)

    if ($filled -lt 0)  { $filled = 0 }
    if ($filled -gt 10) { $filled = 10 }

    return (
        ("■" * $filled) +
        ("□" * (10 - $filled))
    )
}

function Invoke-JsonlProcessWithActivity {
    param(
        [string]$FileName,
        [string]$Arguments,
        [string]$InputText,
        [int]$TimeoutSeconds,
        [string]$StageName,
        [ValidateSet("Claude","Codex")]
        [string]$Provider
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo

    $launchFileName = $FileName
    $launchArguments = $Arguments

    # Windows npm shims such as codex.cmd cannot be started directly
    # when UseShellExecute=False. Route .cmd/.bat through cmd.exe.
    if ($FileName -match '\.(cmd|bat)$') {
        $launchFileName = $env:ComSpec
        $launchArguments = '/d /s /c ""' + $FileName + '" ' + $Arguments + '"'
    }

    $psi.FileName = $launchFileName
    $psi.Arguments = $launchArguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    if ($null -ne $InputText) {
        $psi.RedirectStandardInput = $true
    }

    try {
        $utf8 = New-Object System.Text.UTF8Encoding -ArgumentList $false
        $psi.StandardOutputEncoding = $utf8
        $psi.StandardErrorEncoding = $utf8

        if ($null -ne $InputText) {
            $psi.StandardInputEncoding = $utf8
        }
    }
    catch {}

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    [void]$proc.Start()

    if ($null -ne $InputText) {
        # PowerShell 5.1 redirected stdin can corrupt Korean text.
        # Write raw UTF-8 bytes directly to the process stdin stream.
        $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
        $inputBytes = $utf8NoBom.GetBytes($InputText)

        $proc.StandardInput.BaseStream.Write(
            $inputBytes,
            0,
            $inputBytes.Length
        )

        $proc.StandardInput.BaseStream.Flush()
        $proc.StandardInput.Close()
    }

    $stderrTask = $proc.StandardError.ReadToEndAsync()
    $readTask = $proc.StandardOutput.ReadLineAsync()

    $watch = [System.Diagnostics.Stopwatch]::StartNew()

    $lastActivity = [DateTime]::UtcNow
    $eventCount = 0
    $lastEventType = "starting"
    $lastHeartbeat = -1
    $finalResult = $null

    while ($true) {

        # stdout JSONL을 가능한 만큼 즉시 소비
        while ($null -ne $readTask -and $readTask.IsCompleted) {

            $line = $readTask.Result

            if ($null -eq $line) {
                $readTask = $null
                break
            }

            if (-not [string]::IsNullOrWhiteSpace($line)) {

                $eventCount++
                $lastActivity = [DateTime]::UtcNow

                try {
                    $evt = $line | ConvertFrom-Json -ErrorAction Stop

                    if ($null -ne $evt.type) {
                        $lastEventType = [string]$evt.type
                    }

                    # Claude 최종 결과
                    if (
                        $Provider -eq "Claude" `
                        -and $evt.type -eq "result" `
                        -and $null -ne $evt.result
                    ) {
                        $finalResult = [string]$evt.result
                    }

                    # Codex의 마지막 agent_message를 최종 결과로 사용
                    if (
                        $Provider -eq "Codex" `
                        -and $evt.type -eq "item.completed" `
                        -and $null -ne $evt.item `
                        -and $evt.item.type -eq "agent_message" `
                        -and $null -ne $evt.item.text
                    ) {
                        $finalResult = [string]$evt.item.text
                    }
                }
                catch {
                    # JSON 이벤트가 아닌 진단 라인은 무시
                    $lastEventType = "stream-data"
                }
            }

            $readTask = $proc.StandardOutput.ReadLineAsync()
        }

        $elapsed = [int]$watch.Elapsed.TotalSeconds

        # 15초마다 Energy Bar + 실제 activity 표시
        if (
            $elapsed -gt 0 `
            -and ($elapsed % 15) -eq 0 `
            -and $lastHeartbeat -ne $elapsed
        ) {
            $lastHeartbeat = $elapsed

            $idleSeconds = [int](
                ([DateTime]::UtcNow - $lastActivity).TotalSeconds
            )

            $stagePercent = Get-JungleStagePercent -StageName $StageName
            $bar = Get-JungleProgressBar -Percent $stagePercent

            if ($Provider -eq "Codex") {

                Write-Host (
                    "[{0}] {1}%  {2} · xhigh | RUNNING | {3}s/{4}s | events {5} | last {6}s | {7}" `
                    -f $bar,
                       $stagePercent,
                       $StageName,
                       $elapsed,
                       $TimeoutSeconds,
                       $eventCount,
                       $idleSeconds,
                       $lastEventType
                )
            }
            else {

                Write-Host (
                    "[{0}] {1}%  {2} | RUNNING | {3}s/{4}s | events {5} | last {6}s | {7}" `
                    -f $bar,
                       $stagePercent,
                       $StageName,
                       $elapsed,
                       $TimeoutSeconds,
                       $eventCount,
                       $idleSeconds,
                       $lastEventType
                )
            }
        }

        if ($elapsed -ge $TimeoutSeconds) {

            try {
                $proc.Kill()
            }
            catch {}

            throw "$StageName timeout: ${TimeoutSeconds}s"
        }

        if ($proc.HasExited -and $null -eq $readTask) {
            break
        }

        Start-Sleep -Milliseconds 100
    }

    $proc.WaitForExit()
    $watch.Stop()

    $stderr = $stderrTask.Result

    if ($proc.ExitCode -ne 0) {
        throw "$StageName failed. ExitCode=$($proc.ExitCode)`n$stderr"
    }

    if ([string]::IsNullOrWhiteSpace($finalResult)) {
        throw "$StageName completed but final result was not found in the JSONL stream.`n$stderr"
    }

    $stagePercent = Get-JungleStagePercent -StageName $StageName
    $bar = Get-JungleProgressBar -Percent $stagePercent

    Write-Host (
        "[{0}] {1}%  {2} | COMPLETE | {3}s | events {4}" `
        -f $bar,
           $stagePercent,
           $StageName,
           [int]$watch.Elapsed.TotalSeconds,
           $eventCount
    )

    return $finalResult
}


function Invoke-ClaudeFromFile {
    param(
        [string]$InstructionPath,
        [int]$TimeoutSeconds,
        [string]$StageName
    )

    $shortPrompt = "Read the UTF-8 instruction file at `"$InstructionPath`" and execute it exactly. Work non-interactively."

    $escapedPrompt = $shortPrompt.Replace('"','\"')

    $args = (
        '-p "' +
        $escapedPrompt +
        '" --output-format stream-json --verbose --allowedTools "Read" "Write" "Edit" ' +
        $script:claudeAddDirArgs
    )

    return Invoke-JsonlProcessWithActivity `
        -FileName $claudePath `
        -Arguments $args `
        -InputText $null `
        -TimeoutSeconds $TimeoutSeconds `
        -StageName $StageName `
        -Provider "Claude"
}

function Invoke-CodexJson {
    param(
        [string]$Instruction,
        [int]$TimeoutSeconds,
        [string]$StageName
    )

    $args = "exec --json --skip-git-repo-check --sandbox read-only --config model_reasoning_effort=xhigh -"

    return Invoke-JsonlProcessWithActivity `
        -FileName $codexPath `
        -Arguments $args `
        -InputText $Instruction `
        -TimeoutSeconds $TimeoutSeconds `
        -StageName $StageName `
        -Provider "Codex"
}

Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE FAST VERIFIED + SKILLS v3"
Write-Host "========================================"
Write-Host "1/2 Claude draft + relevant skills..."

# ============================================================
# JUNGLE VISUAL QA PREPARATION
# ============================================================

$visualQaRequired = $false
$visualQaEnabled = $false
$visualManifestPath = ""
$visualQaError = ""
$visualQaSummary = "NOT_REQUIRED"

$htmlReviewMatch =
    [regex]::Match(
        $Request,
        '(?i)([A-Z]:\\[^"\r\n]+?\.html?)'
    )

$visualReviewIntent =
    $Request -match
    '(?i)(검토|분석|리뷰|review|check|평가|피드백|점검)'

if (
    $htmlReviewMatch.Success `
    -and $visualReviewIntent
) {

    $visualSourceHtml =
        $htmlReviewMatch.Groups[1].Value

    if (
        Test-Path `
            -LiteralPath $visualSourceHtml `
            -PathType Leaf
    ) {

        try {

            $visualSourceText =
                [System.IO.File]::ReadAllText(
                    $visualSourceHtml
                )

            $visualQaRequired =
                [regex]::IsMatch(
                    $visualSourceText,
                    '(?i)<(?:img|video)\b'
                )

            if ($visualQaRequired) {

                $visualRoot =
                    "C:\AI\work\visual"

                New-Item `
                    -ItemType Directory `
                    -Path $visualRoot `
                    -Force |
                    Out-Null

                Write-Host ""
                Write-Host "[VISUAL QA] Embedded media detected"
                Write-Host "[VISUAL QA] Preparing visual assets..."

                $visualBundle =
                    New-JungleHtmlVisualBundle `
                        -HtmlPath $visualSourceHtml `
                        -WorkRoot $visualRoot `
                        -Stamp $stamp

                $visualManifestPath =
                    $visualBundle.ManifestPath

                $visualQaEnabled =
                    (
                        $visualBundle.ImageCount -gt 0 `
                        -or $visualBundle.FrameCount -gt 0
                    )

                $visualQaSummary =
                    (
                        "images={0}, videos={1}, frames={2}/{3}" `
                        -f $visualBundle.ImageCount,
                           $visualBundle.VideoCount,
                           $visualBundle.FrameCount,
                           $visualBundle.ExpectedFrameCount
                    )

                Write-Host (
                    "[VISUAL QA] {0}" `
                    -f $visualQaSummary
                )

                Write-Host (
                    "[VISUAL QA] Manifest: {0}" `
                    -f $visualManifestPath
                )
            }
        }
        catch {

            $visualQaError =
                $_.Exception.Message

            $visualQaSummary =
                "PREPARATION_FAILED"

            Write-Host (
                "[VISUAL QA] PREPARATION FAILED: {0}" `
                -f $visualQaError
            )
        }
    }
}

# ============================================================
# ============================================================
# JUNGLE_ADAPTIVE_CLAUDE_TIMEOUT_V1
# Default Claude timeout stays unchanged.
# Visual-QA jobs get a longer execution window.
# ============================================================

$effectiveClaudeTimeout =
    $ClaudeTimeoutSeconds

if (
    $visualQaRequired `
    -and $effectiveClaudeTimeout -lt 1200
) {
    $effectiveClaudeTimeout = 1200
}

Write-Host (
    "[CLAUDE TIMEOUT] effective={0}s | visualQA={1}" `
    -f $effectiveClaudeTimeout,
       $visualQaRequired
)

# END JUNGLE_ADAPTIVE_CLAUDE_TIMEOUT_V1
# ============================================================
$claudeInstruction = @"
[USER REQUEST FILE]
$requestPath

Read the request file and create/perform the best possible first result.

[ARTIFACT TRANSFORMATION POLICY]
- When the user says to make/revise an artifact "according to", "based on", or "following" a review, feedback, correction, guide, or instruction file, treat that file as EDIT INSTRUCTIONS, not as the content to publish.
- Read the instruction/review file first.
- Identify the original/target artifact referenced inside it. Resolve relative filenames from the instruction file's directory.
- Read the original/target artifact.
- Apply the requested corrections, deletions, consolidations, wording changes, and additions to the ORIGINAL content.
- The final deliverable must be the REVISED VERSION OF THE ORIGINAL, not a review report or a reformatted copy of the instructions.
- Preserve the original document's structure and intent unless the instructions explicitly require changes.
- For PDF output, create a printable HTML representation of the revised original in C:\AI\work\temp.
- Name the revised HTML from the original target, preferably <original-name>_수정본.html.
- Do not claim completion unless the revised artifact content has actually been created.

[AVAILABLE CLAUDE AGENTS/SKILLS]
$claudeCatalog

[SELECTION POLICY]
- Use relevant installed skill/agent rules only as guidance. Do NOT invoke subagents or Task in the default fast path.
- Korean technical Markdown: prefer docs-writer-kr when applicable.
- New diagram: drawio-figure when applicable.
- mermaid-figure only for existing Mermaid/.mmd maintenance.
- Do not force Markdown conventions onto PPTX/DOCX/XLSX/coding/unrelated work.
- User-requested output format overrides skill defaults.
- NON-INTERACTIVE MODE: do not stop for optional questions. Use sensible defaults; ask only when execution is impossible without a missing decision.
- Never invent facts.
- If a renderer is missing, do not install software automatically; complete safe work and report the unexecuted render step.
- For PDF requests, Jungle performs the final local PDF rendering after your work. Create or edit the required printable HTML, but do NOT claim that the final PDF was not created merely because you cannot run Edge yourself.
- Do NOT instruct the user to press Ctrl+P or manually export a PDF unless Jungle's local renderer actually reports failure.
- Do NOT mention Google Drive, connector authorization, cloud permissions, or unrelated integrations unless the user's request explicitly requires them.
- No chain-of-thought.

[LOCAL CAPABILITIES]
drawio CLI=$drawio
mmdc=$mmdc
Python=$python
[JUNGLE VISUAL QA]
VISUAL_QA_REQUIRED: $visualQaRequired
VISUAL_QA_ENABLED: $visualQaEnabled
VISUAL_QA_MANIFEST: $visualManifestPath
VISUAL_QA_STATUS: $visualQaSummary
VISUAL_QA_ERROR: $visualQaError

If VISUAL_QA_REQUIRED is True:
- Read VISUAL_QA_MANIFEST before completing the review.
- Use the Read tool to visually inspect EVERY file listed under [IMAGES].
- Use the Read tool to visually inspect EVERY PNG listed under [VIDEO REPRESENTATIVE FRAMES].
- Cross-check visual content against the source HTML text and the user's request.
- Include visual findings in the requested deliverable itself, not only in your CLI prose.
- Check readability, clipping, overlap, number/label consistency, screenshots, charts, UI states, visible wording, and obvious mismatch between explanatory text and screen contents.
- For video frames, review visible progression and wording only.
- Never claim that MP4 audio was reviewed.
- If expected video frames are missing, visual QA is incomplete.
- In your final CLI response include exactly one of:
  VISUAL_QA_COMPLETE: TRUE
  VISUAL_QA_COMPLETE: FALSE
- Use TRUE only after every listed image and representative video frame has actually been visually inspected.
"@

[System.IO.File]::WriteAllText($claudeInstructionPath,$claudeInstruction,$encBom)
$jungleRunStart = Get-Date

$draft = Invoke-ClaudeFromFile -InstructionPath $claudeInstructionPath -TimeoutSeconds $effectiveClaudeTimeout -StageName "Claude draft"


# ============================================================
# JUNGLE_LOCAL_PDF_RENDER_V1
#
# MD -> Claude HTML -> Local Edge -> PDF
# Claude/Codex에게 외부 프로그램 실행 권한을 주지 않고
# Jungle PowerShell이 로컬에서 PDF를 생성한다.
# ============================================================

if (
    $Request -match '(?i)\.md\b' `
    -and $Request -match '(?i)pdf'
) {

    Write-Host ""
    Write-Host "[LOCAL PDF] PDF artifact request detected"

    $mdMatch = [regex]::Match(
        $Request,
        '(?i)([A-Z]:\\[^"\r\n]+?\.md)'
    )

    if (-not $mdMatch.Success) {
        throw "[LOCAL PDF] 요청에서 MD 파일 경로를 찾지 못했습니다."
    }

    $mdPath = $mdMatch.Groups[1].Value

    if (-not (Test-Path -LiteralPath $mdPath)) {
        throw "[LOCAL PDF] 원본 MD 파일을 찾지 못했습니다: $mdPath"
    }

    # 기본값: 일반 MD -> PDF
    $pdfOut = [System.IO.Path]::ChangeExtension(
        $mdPath,
        ".pdf"
    )

    # 검토/수정 지시서이면 "검토 대상" 원본 기준으로 수정본 파일명 생성
    $mdContent = Get-Content -LiteralPath $mdPath -Raw -Encoding UTF8
    $mdDirForTarget = [System.IO.Path]::GetDirectoryName($mdPath)

    $targetMatch = [regex]::Match(
        $mdContent,
        '(?im)검토\s*대상\s*:\s*`?([^`\r\n]+?\.(?:pdf|docx|hwp|hwpx|pptx|xlsx))`?'
    )

    $revisionMode = $false
    $originalTargetPath = $null

    if ($targetMatch.Success) {

        $targetRef = $targetMatch.Groups[1].Value.Trim()

        if ([System.IO.Path]::IsPathRooted($targetRef)) {
            $candidateTarget = $targetRef
        }
        else {
            $candidateTarget = Join-Path $mdDirForTarget $targetRef
        }

        if (Test-Path -LiteralPath $candidateTarget) {

            $revisionMode = $true
            $originalTargetPath = $candidateTarget

            $targetDir = [System.IO.Path]::GetDirectoryName($candidateTarget)
            $targetStem = [System.IO.Path]::GetFileNameWithoutExtension($candidateTarget)

            $pdfOut = Join-Path $targetDir ($targetStem + "_수정본.pdf")

            Write-Host "[LOCAL PDF] REVISION MODE"
            Write-Host "[LOCAL PDF] INSTRUCTIONS :" $mdPath
            Write-Host "[LOCAL PDF] ORIGINAL     :" $originalTargetPath
            Write-Host "[LOCAL PDF] OUTPUT       :" $pdfOut
        }
    }

    # 이번 Jungle 실행에서 Claude가 생성한 가장 최근 HTML 찾기
    $htmlCandidate = Get-ChildItem `
        -LiteralPath "C:\AI\work\temp" `
        -Filter "*.html" `
        -File `
        -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LastWriteTime -ge $jungleRunStart.AddSeconds(-5)
        } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $htmlCandidate) {

        $mdDir = [System.IO.Path]::GetDirectoryName($mdPath)

        if (
            -not [string]::IsNullOrWhiteSpace($mdDir) `
            -and (Test-Path -LiteralPath $mdDir)
        ) {

            $htmlCandidate = Get-ChildItem `
                -LiteralPath $mdDir `
                -Filter "*.html" `
                -File `
                -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.LastWriteTime -ge $jungleRunStart.AddMinutes(-15)
                } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
        }
    }

    if ($null -eq $htmlCandidate) {
        throw "[LOCAL PDF] Claude가 생성한 HTML 파일을 찾지 못했습니다."
    }

    $htmlPath = $htmlCandidate.FullName

    Write-Host "[LOCAL PDF] HTML :" $htmlPath
    Write-Host "[LOCAL PDF] PDF  :" $pdfOut

    # Edge 위치 검색
    $edgeCandidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )

    $edgePath = $null

    foreach ($candidate in $edgeCandidates) {
        if (
            -not [string]::IsNullOrWhiteSpace($candidate) `
            -and (Test-Path -LiteralPath $candidate)
        ) {
            $edgePath = $candidate
            break
        }
    }

    if ($null -eq $edgePath) {

        $edgeCmd = Get-Command msedge.exe -ErrorAction SilentlyContinue

        if ($null -ne $edgeCmd) {
            $edgePath = $edgeCmd.Source
        }
    }

    if ($null -eq $edgePath) {
        throw "[LOCAL PDF] Microsoft Edge 실행 파일을 찾지 못했습니다."
    }

    # 기존 PDF가 있으면 제거
    if (Test-Path -LiteralPath $pdfOut) {
        Remove-Item -LiteralPath $pdfOut -Force
    }

    $htmlUri = (New-Object System.Uri($htmlPath)).AbsoluteUri

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $edgePath

    $psi.Arguments = (
        '--headless=new ' +
        '--disable-gpu ' +
        '--no-pdf-header-footer ' +
        '--print-to-pdf="' + $pdfOut + '" ' +
        '"' + $htmlUri + '"'
    )

    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    Write-Host "[LOCAL PDF] Rendering with Edge..."

    $edgeProc = New-Object System.Diagnostics.Process
    $edgeProc.StartInfo = $psi

    [void]$edgeProc.Start()

    $edgeStdout = $edgeProc.StandardOutput.ReadToEndAsync()
    $edgeStderr = $edgeProc.StandardError.ReadToEndAsync()

    $edgeElapsed = 0

    while (-not $edgeProc.HasExited) {

        Start-Sleep -Seconds 1
        $edgeElapsed++

        if (($edgeElapsed % 5) -eq 0) {
            Write-Host (
                "[LOCAL PDF] RUNNING | {0}s" -f $edgeElapsed
            )
        }

        if ($edgeElapsed -ge 120) {

            try {
                $edgeProc.Kill()
            }
            catch {}

            throw "[LOCAL PDF] Edge PDF rendering timeout: 120s"
        }
    }

    $edgeProc.WaitForExit()

    # Edge 종료 직후 파일 flush 대기
    $wait = 0

    while (
        -not (Test-Path -LiteralPath $pdfOut) `
        -and $wait -lt 10
    ) {
        Start-Sleep -Seconds 1
        $wait++
    }

    if (-not (Test-Path -LiteralPath $pdfOut)) {

        $edgeError = ""

        try {
            $edgeError = $edgeStderr.Result
        }
        catch {}

        throw "[LOCAL PDF] PDF 파일이 생성되지 않았습니다.`n$edgeError"
    }

    $pdfInfo = Get-Item -LiteralPath $pdfOut

    if ($pdfInfo.Length -lt 1000) {
        throw "[LOCAL PDF] 생성된 PDF 파일 크기가 비정상적으로 작습니다: $($pdfInfo.Length) bytes"
    }

    Write-Host ""
    Write-Host "[■■■■□□□□□□] 40%  LOCAL PDF | COMPLETE"
    Write-Host "[LOCAL PDF] FILE :" $pdfOut
    Write-Host "[LOCAL PDF] SIZE :" $pdfInfo.Length "bytes"
    Write-Host ""

    # 이후 Codex QA가 실제 PDF 생성 성공을 알 수 있도록
    # Claude 결과에 로컬 artifact 상태를 추가한다.
    $draft = (
        $draft +
        "`r`n`r`n" +
        "[JUNGLE LOCAL ARTIFACT STATUS]`r`n" +
        "PDF_CREATED: TRUE`r`n" +
        "SOURCE_MD: $mdPath`r`n" +
        "SOURCE_HTML: $htmlPath`r`n" +
        "OUTPUT_PDF: $pdfOut`r`n" +
        "PDF_SIZE_BYTES: $($pdfInfo.Length)"
    )
}
# ============================================================
# GENERAL HTML ARTIFACT DISCOVERY
# ============================================================

$htmlDeliverableRequested =
    ($Request -match '(?i)html\s*(?:로|형식(?:으로)?|결과물|output|deliverable)') -or
    ($Request -match '(?i)(?:결과물|출력|output|deliverable).{0,40}html') -or
    ($Request -match '(?i)(?:as|to)\s+html')

$htmlPathValue = $null
$htmlPathVar =
    Get-Variable -Name htmlPath -ErrorAction SilentlyContinue

if ($null -ne $htmlPathVar) {
    $htmlPathValue = [string]$htmlPathVar.Value
}

if (
    $htmlDeliverableRequested `
    -and [string]::IsNullOrWhiteSpace($htmlPathValue)
) {

    $generalHtmlCandidate =
        Get-ChildItem `
            -LiteralPath $tempDir `
            -Filter "*.html" `
            -File `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LastWriteTime -ge $jungleRunStart.AddSeconds(-5)
        } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -ne $generalHtmlCandidate) {
        $htmlPath = $generalHtmlCandidate.FullName
        $htmlPathValue = $htmlPath
    }
}

if (
    $htmlDeliverableRequested `
    -and -not [string]::IsNullOrWhiteSpace($htmlPathValue) `
    -and (Test-Path -LiteralPath $htmlPathValue -PathType Leaf)
) {

    $htmlDraftInfo =
        Get-Item -LiteralPath $htmlPathValue

    $draft = (
        $draft +
        "`r`n`r`n" +
        "[JUNGLE LOCAL HTML ARTIFACT STATUS]`r`n" +
        "HTML_CREATED: TRUE`r`n" +
        "OUTPUT_HTML: $htmlPathValue`r`n" +
        "HTML_SIZE_BYTES: $($htmlDraftInfo.Length)"
    )
}

# ============================================================
[System.IO.File]::WriteAllText($draftPath,$draft,$encBom)

Write-Host "2/2 Codex QA + relevant skills..."

$qaInstruction = @"
You are Jungle's fast universal QA reviewer.

Read:
REQUEST: $requestPath
CLAUDE RESULT: $draftPath
VISUAL QA REQUIRED: $visualQaRequired
VISUAL QA ENABLED: $visualQaEnabled
VISUAL QA MANIFEST: $visualManifestPath

VISUAL QA GATE:
- If VISUAL QA REQUIRED is True, read the manifest.
- The Claude result must contain VISUAL_QA_COMPLETE: TRUE.
- If the marker is missing or FALSE, treat incomplete visual inspection as at least one major issue.
- If RENDERED_VIDEO_FRAMES is lower than EXPECTED_VIDEO_FRAMES, visual QA is incomplete.
- Do not require audio inspection unless the user's request explicitly requires audio review.

AVAILABLE CODEX SKILLS:
$codexCatalog

Use only relevant installed skills.
For Korean business/document deliverables, apply ax-deliverables and humanizer-kr when available and relevant.

Verify:
- request fulfillment
- correctness/grounding
- missing requirements
- contradiction/logical gap
- unsupported facts/numbers
- duplication
- clarity/user fit
- whether requested local actions/file operations were actually completed

Do not use web unless the original request itself requires current/public research.
No chain-of-thought.

Return ONLY JSON:
{
  "status":"PASS | REVISE | FAIL",
  "score":0,
  "critical":0,
  "major":0,
  "minor":0,
  "skills_used":[],
  "issues":[
    {"severity":"critical | major | minor","problem":"","recommendation":""}
  ],
  "summary":""
}

ARTIFACT TRUTH POLICY:
- If the candidate contains [JUNGLE LOCAL ARTIFACT STATUS], treat that block as authoritative for artifact existence, output path, and file size.
- If PDF_CREATED: TRUE and the local artifact status provides a valid output path/size, do NOT penalize stale model prose claiming that PDF creation failed, shell access was unavailable, or manual printing is required.
- Evaluate the user's requested deliverable and the actual artifact status, not obsolete execution commentary.
- Do not treat irrelevant stale connector/permission commentary as a deliverable defect unless it affected the actual artifact.
- Content, structure, omissions, factual errors, and artifact-quality problems must still be scored normally.
- If [JUNGLE LOCAL HTML ARTIFACT STATUS] says HTML_CREATED: TRUE, read OUTPUT_HTML and evaluate the actual HTML deliverable, not only Claude's prose summary.
- Treat OUTPUT_HTML path and HTML_SIZE_BYTES as authoritative for HTML artifact existence.

FAST-PASS POLICY:
- Finalize without another model call when critical=0, major=0, and score >= $FastPassScore.
- Minor polish alone must not trigger a revision cycle.
"@

$qaJson = Invoke-CodexJson -Instruction $qaInstruction -TimeoutSeconds $CodexTimeoutSeconds -StageName "Codex QA"
[System.IO.File]::WriteAllText($qaPath,$qaJson,$encBom)
$qa = $qaJson | ConvertFrom-Json

$needsRevision =
    ([int]$qa.critical -gt 0) -or
    ([int]$qa.major -gt 0) -or
    ([int]$qa.score -lt $FastPassScore)

if (-not $needsRevision) {

    [System.IO.File]::WriteAllText(
        $finalPath,
        $draft,
        $encBom
    )

    Write-Host ""

    # pdfOut 변수가 존재하는 경우에만 확인
    $pdfOutValue = $null
    $pdfOutVar = Get-Variable -Name pdfOut -ErrorAction SilentlyContinue

    if ($null -ne $pdfOutVar) {
        $pdfOutValue = [string]$pdfOutVar.Value
    }

    if (
        -not [string]::IsNullOrWhiteSpace($pdfOutValue) `
        -and (Test-Path -LiteralPath $pdfOutValue)
    ) {

        $fastArtifact = Get-Item -LiteralPath $pdfOutValue

        Write-Host "Verified Artifact :" $fastArtifact.FullName
        Write-Host "Artifact Size     :" $fastArtifact.Length "bytes"
    }
    else {

        $fastHtmlArtifact = $null

        if (
            $htmlDeliverableRequested `
            -and -not [string]::IsNullOrWhiteSpace($htmlPath) `
            -and (Test-Path -LiteralPath $htmlPath -PathType Leaf)
        ) {
            $fastHtmlArtifact =
                Publish-JungleHtmlArtifact `
                    -HtmlPath $htmlPath `
                    -OutputDir $outputDir `
                    -Stamp $stamp
        }

        if ($null -ne $fastHtmlArtifact) {
            Write-Host "Verified Artifact :" $fastHtmlArtifact.FullName
            Write-Host "Artifact Size     :" $fastHtmlArtifact.Length "bytes"
        }
        else {
            # 일반 텍스트 요청은 실제 Claude 결과를 출력
            Write-Host $draft
        }
    }

    Write-Host ""
    Write-Host "QA Score          :" $qa.score
    Write-Host "Critical          :" $qa.critical
    Write-Host "Major             :" $qa.major
    Write-Host "Minor             :" $qa.minor
    Write-Host "Skills            :" (@($qa.skills_used) -join ", ")
    Write-Host "Decision          : FAST_PASS"
    Write-Host "Final Log         :" $finalPath
    Write-Host "[■■■■■■■■■■] 100%  COMPLETE"
    Write-Host "STATUS: JUNGLE_FAST_VERIFIED_COMPLETE"
    return
}

Write-Host ""
Write-Host "Major/Critical/low-score detected -> one revision cycle"
Write-Host "3/4 Claude revision..."

$revisionInstruction = @"
Read:
REQUEST: $requestPath
ORIGINAL RESULT: $draftPath
CODEX QA: $qaPath

VISUAL QA REQUIRED: $visualQaRequired
VISUAL QA ENABLED: $visualQaEnabled
VISUAL QA MANIFEST: $visualManifestPath

If visual QA is required:
- Read the visual manifest.
- Inspect any image or representative video frame needed to resolve the supported QA issues.
- If the original result failed because visual assets were not actually inspected, inspect EVERY listed image and representative video frame now.
- Apply supported visual findings to the actual HTML deliverable.
- Never claim MP4 audio was reviewed.
- End the CLI response with VISUAL_QA_COMPLETE: TRUE only if visual inspection is complete.

Revise ONCE using only supported QA issues.

Relevant Claude agents/skills:
$claudeCatalog

Rules:
- Never invent facts/numbers/dates/owners/targets/sources.
- Preserve correct material and completed file operations.
- User-requested output format overrides skill defaults.
- NON-INTERACTIVE MODE.
- No chain-of-thought.

[ARTIFACT REVISION]
- If CURRENT HTML PATH below is non-empty and exists, do NOT merely describe corrections.
- Read that HTML file and EDIT/WRITE the actual file so the QA issues are applied to the deliverable.
- Preserve all correct content.
- The revised HTML must represent the final revised version of the original target artifact.
- Do not create another review report.

CURRENT HTML PATH: $htmlPath
FINAL PDF PATH: $pdfOut
"@

[System.IO.File]::WriteAllText($revisionInstructionPath,$revisionInstruction,$encBom)

$revision = Invoke-ClaudeFromFile `
    -InstructionPath $revisionInstructionPath `
    -TimeoutSeconds $effectiveClaudeTimeout `
    -StageName "Claude revision"


# ============================================================
# FINAL ARTIFACT RE-RENDER
# Claude revision이 실제 HTML을 수정한 뒤 PDF를 다시 생성한다.
# ============================================================

if (
    -not [string]::IsNullOrWhiteSpace($htmlPath) `
    -and -not [string]::IsNullOrWhiteSpace($pdfOut) `
    -and -not [string]::IsNullOrWhiteSpace($edgePath) `
    -and (Test-Path -LiteralPath $htmlPath)
) {

    Write-Host ""
    Write-Host "[■■■■■■■■□□] 85%  LOCAL PDF RE-RENDER | START"
    Write-Host "[LOCAL PDF] HTML :" $htmlPath
    Write-Host "[LOCAL PDF] PDF  :" $pdfOut

    if (Test-Path -LiteralPath $pdfOut) {
        Remove-Item -LiteralPath $pdfOut -Force
    }

    $htmlUri = (New-Object System.Uri($htmlPath)).AbsoluteUri

    $psiFinal = New-Object System.Diagnostics.ProcessStartInfo
    $psiFinal.FileName = $edgePath

    $psiFinal.Arguments = (
        '--headless=new ' +
        '--disable-gpu ' +
        '--no-pdf-header-footer ' +
        '--print-to-pdf="' + $pdfOut + '" ' +
        '"' + $htmlUri + '"'
    )

    $psiFinal.UseShellExecute = $false
    $psiFinal.CreateNoWindow = $true
    $psiFinal.RedirectStandardOutput = $true
    $psiFinal.RedirectStandardError = $true

    $edgeFinal = New-Object System.Diagnostics.Process
    $edgeFinal.StartInfo = $psiFinal

    [void]$edgeFinal.Start()

    $finalStdout = $edgeFinal.StandardOutput.ReadToEndAsync()
    $finalStderr = $edgeFinal.StandardError.ReadToEndAsync()

    $renderElapsed = 0

    while (-not $edgeFinal.HasExited) {

        Start-Sleep -Seconds 1
        $renderElapsed++

        if (($renderElapsed % 5) -eq 0) {
            Write-Host (
                "[■■■■■■■■□□] 85%  LOCAL PDF RE-RENDER | RUNNING | {0}s" `
                -f $renderElapsed
            )
        }

        if ($renderElapsed -ge 120) {

            try {
                $edgeFinal.Kill()
            }
            catch {}

            throw "[LOCAL PDF] Final PDF re-render timeout: 120s"
        }
    }

    $edgeFinal.WaitForExit()

    $flushWait = 0

    while (
        -not (Test-Path -LiteralPath $pdfOut) `
        -and $flushWait -lt 10
    ) {
        Start-Sleep -Seconds 1
        $flushWait++
    }

    if (-not (Test-Path -LiteralPath $pdfOut)) {

        $renderError = ""

        try {
            $renderError = $finalStderr.Result
        }
        catch {}

        throw "[LOCAL PDF] 최종 수정본 PDF 재생성 실패.`n$renderError"
    }

    $finalPdfInfo = Get-Item -LiteralPath $pdfOut

    if ($finalPdfInfo.Length -lt 1000) {
        throw "[LOCAL PDF] 최종 PDF 크기가 비정상적입니다: $($finalPdfInfo.Length) bytes"
    }

    Write-Host ""
    Write-Host "[■■■■■■■■□□] 85%  LOCAL PDF RE-RENDER | COMPLETE"
    Write-Host "[LOCAL PDF] FINAL FILE :" $pdfOut
    Write-Host "[LOCAL PDF] FINAL SIZE :" $finalPdfInfo.Length "bytes"
    Write-Host ""

    $revision = (
        $revision +
        "`r`n`r`n" +
        "[JUNGLE FINAL ARTIFACT STATUS]`r`n" +
        "FINAL_PDF_RE_RENDERED: TRUE`r`n" +
        "FINAL_HTML: $htmlPath`r`n" +
        "FINAL_PDF: $pdfOut`r`n" +
        "FINAL_PDF_SIZE_BYTES: $($finalPdfInfo.Length)"
    )
}

# ============================================================
# GENERAL HTML REVISION ARTIFACT STATUS
# ============================================================

if (
    $htmlDeliverableRequested `
    -and -not [string]::IsNullOrWhiteSpace($htmlPath) `
    -and (Test-Path -LiteralPath $htmlPath -PathType Leaf)
) {

    $htmlRevisionInfo =
        Get-Item -LiteralPath $htmlPath

    $revision = (
        $revision +
        "`r`n`r`n" +
        "[JUNGLE FINAL HTML ARTIFACT STATUS]`r`n" +
        "FINAL_HTML_CREATED: TRUE`r`n" +
        "FINAL_HTML: $htmlPath`r`n" +
        "FINAL_HTML_SIZE_BYTES: $($htmlRevisionInfo.Length)"
    )
}

# ============================================================
[System.IO.File]::WriteAllText(
    $revisionPath,
    $revision,
    $encBom
)

Write-Host "4/4 Codex re-QA..."

$reQaInstruction = @"
You are Jungle's final QA gate.

Read:
REQUEST: $requestPath
ORIGINAL QA: $qaPath
REVISED RESULT: $revisionPath
VISUAL QA REQUIRED: $visualQaRequired
VISUAL QA ENABLED: $visualQaEnabled
VISUAL QA MANIFEST: $visualManifestPath

FINAL VISUAL QA GATE:
- If VISUAL QA REQUIRED is True, read the manifest.
- The revised result must state VISUAL_QA_COMPLETE: TRUE.
- Missing or FALSE means visual review is still incomplete and must remain a major issue.
- Verify that the revised deliverable incorporates material visual findings where appropriate.
- Do not penalize lack of audio review unless audio was explicitly part of the user's request.

AVAILABLE CODEX SKILLS:
$codexCatalog

Return ONLY JSON:
{
  "decision":"ACCEPT_REVISION | KEEP_ORIGINAL",
  "score":0,
  "critical":0,
  "major":0,
  "minor":0,
  "skills_used":[],
  "summary":""
}

Rules:
- [JUNGLE FINAL ARTIFACT STATUS] is authoritative for final artifact existence, path, re-render status, and file size.
- If FINAL_PDF_RE_RENDERED: TRUE and FINAL_PDF is present in the artifact status, ignore stale model prose that says PDF creation failed, shell access was unavailable, or manual printing is required.
- Do not penalize obsolete Google Drive, shell-permission, or manual-export commentary when it did not affect the actual final artifact.
- Judge the revised deliverable itself: requested edits, omissions, factual accuracy, structure, readability, and artifact completion.
- If [JUNGLE FINAL HTML ARTIFACT STATUS] says FINAL_HTML_CREATED: TRUE, read FINAL_HTML and judge the actual revised HTML deliverable.
- FINAL_HTML path and FINAL_HTML_SIZE_BYTES are authoritative for final HTML artifact existence.
- ACCEPT_REVISION only when critical=0, major=0, and score >= $FastPassScore.
- Otherwise KEEP_ORIGINAL.
- KEEP_ORIGINAL if critical or major increased.
- KEEP_ORIGINAL if overall quality clearly decreased.
- No chain-of-thought.
"@

$reQaJson = Invoke-CodexJson -Instruction $reQaInstruction -TimeoutSeconds $CodexTimeoutSeconds -StageName "Codex re-QA"
[System.IO.File]::WriteAllText($reQaPath,$reQaJson,$encBom)
$reQa = $reQaJson | ConvertFrom-Json

# ============================================================
# HARD FINAL QUALITY GATE
# ============================================================

$finalPass =
    ([int]$reQa.critical -eq 0) -and
    ([int]$reQa.major -eq 0) -and
    ([int]$reQa.score -ge $FastPassScore)

# revision cycle의 최신 결과를 candidate로 기록
$final = $revision

[System.IO.File]::WriteAllText(
    $finalPath,
    $final,
    $encBom
)

Write-Host ""
Write-Host "Initial QA Score :" $qa.score
Write-Host "Final QA Score   :" $reQa.score
Write-Host "Critical         :" $reQa.critical
Write-Host "Major            :" $reQa.major
Write-Host "Minor            :" $reQa.minor
Write-Host "Model Decision   :" $reQa.decision
Write-Host "Skills           :" (@($reQa.skills_used) -join ", ")
Write-Host "Final Log        :" $finalPath

if (
    -not [string]::IsNullOrWhiteSpace($pdfOut) `
    -and (Test-Path -LiteralPath $pdfOut)
) {
    $artifactInfo = Get-Item -LiteralPath $pdfOut

    Write-Host "Final Artifact   :" $artifactInfo.FullName
    Write-Host "Artifact Size    :" $artifactInfo.Length "bytes"
}

if ($finalPass) {

    $finalHtmlArtifact = $null

    if (
        $htmlDeliverableRequested `
        -and -not [string]::IsNullOrWhiteSpace($htmlPath) `
        -and (Test-Path -LiteralPath $htmlPath -PathType Leaf)
    ) {

        $finalHtmlArtifact =
            Publish-JungleHtmlArtifact `
                -HtmlPath $htmlPath `
                -OutputDir $outputDir `
                -Stamp $stamp
    }

    if ($null -ne $finalHtmlArtifact) {
        Write-Host "Final Artifact   :" $finalHtmlArtifact.FullName
        Write-Host "Artifact Size    :" $finalHtmlArtifact.Length "bytes"
    }

    Write-Host ""
    Write-Host "Decision         : ACCEPT_REVISION"
    Write-Host "[■■■■■■■■■■] 100%  COMPLETE"
    Write-Host "STATUS: JUNGLE_FAST_VERIFIED_COMPLETE"
}
else {

    Write-Host ""
    Write-Host "Decision         : NEEDS_REVIEW"
    Write-Host "[■■■■■■■■■□] 90%  QA GATE NOT PASSED"

    Write-Host (
        "Required         : score >= {0}, critical=0, major=0" `
        -f $FastPassScore
    )

    Write-Host (
        "Actual           : score={0}, critical={1}, major={2}" `
        -f $reQa.score,
           $reQa.critical,
           $reQa.major
    )

    Write-Host ""
    Write-Host "Latest artifact exists, but Jungle does NOT mark it complete."
    Write-Host "STATUS: JUNGLE_FAST_VERIFIED_NEEDS_REVIEW"
}

