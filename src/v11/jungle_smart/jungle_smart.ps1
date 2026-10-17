param(
    [Parameter(Mandatory=$true)]
    [string]$Request,

    [switch]$DryRun
)

$ErrorActionPreference =
    "Stop"


$root =
    Split-Path `
        -Parent `
        $MyInvocation.MyCommand.Path


$configPath =
    Join-Path `
        $root `
        "config\jungle.smart.json"


if (-not (
    Test-Path `
        -LiteralPath $configPath `
        -PathType Leaf
)) {

    throw "Jungle Smart config missing."
}


$config =
    Get-Content `
        -LiteralPath $configPath `
        -Raw |
    ConvertFrom-Json


. (Join-Path $root "modules\router.ps1")
. (Join-Path $root "modules\state.ps1")
. (Join-Path $root "modules\artifact.ps1")
. (Join-Path $root "modules\verifier.ps1")
. (Join-Path $root "modules\watchdog.ps1")
. (Join-Path $root "modules\worker.ps1")
. (Join-Path $root "modules\deterministic_g1.ps1")
. (Join-Path $root "modules\changed_scope.ps1")
. (Join-Path $root "modules\selective_visual_qa.ps1")
. (Join-Path $root "modules\blind_pairwise_qa.ps1")
. (Join-Path $root "providers\claude.ps1")
. (Join-Path $root "providers\codex.ps1")


$context =
    New-JungleRunContext `
        -Root $root `
        -Request $Request


$route =
    Get-JungleRoute `
        -Request $Request `
        -Config $config


Write-Host ""
Write-Host "========================================"
Write-Host " JUNGLE SMART"
Write-Host "========================================"
Write-Host ""

Write-Host "VERSION   : $($config.version)"
Write-Host "PHASE     : $($config.phase)"
Write-Host "RUN ID    : $($context.run_id)"
Write-Host ""

Write-Host "ROUTE     : $($route.route)"
Write-Host "PROVIDER  : $($route.provider)"
Write-Host "MODE      : $($route.mode)"
Write-Host "REASON    : $($route.reason)"
Write-Host ""


if (
    $DryRun `
    -or
    -not $config.feature_flags.execution_enabled
) {

    Write-Host "EXECUTION : DISABLED"
    Write-Host "AI CALLS  : 0"
    Write-Host ""
    Write-Host "STATUS : JUNGLE_SMART_DRY_RUN_COMPLETE"
    Write-Host "========================================"

    return
}


switch ($route.route) {

        "PPT_REFLOW" {

            # JUNGLE_SMART_REFLOW_MAIN_ISOLATION_V1
            #
            # Completely separate from PPT_SAFE.

            if (-not $config.feature_flags.reflow_v1) {

                throw (
                    "PPT_REFLOW route detected, but REFLOW V1 feature flag is OFF."
                )
            }


            $reflowBridge =
                Join-Path `
                    $root `
                    "executors\ppt_reflow_bridge.ps1"


            if (-not (
                Test-Path `
                    -LiteralPath $reflowBridge `
                    -PathType Leaf
            )) {

                throw (
                    "REFLOW bridge missing: " +
                    $reflowBridge
                )
            }


            & $reflowBridge `
                -Request $Request `
                -ConfigPath $configPath


            return
        }

    "PPT_SAFE" {

        # --------------------------------------------------------
        # Resolve source PPT path.
        # --------------------------------------------------------

        $sourceMatch =
            [regex]::Match(
                $Request,
                '(?i)(?:"(?<p>[A-Z]:\\[^"]+?\.pptx?)"|''(?<p>[A-Z]:\\[^'']+?\.pptx?)''|(?<p>[A-Z]:\\.*?\.pptx?)(?=\s|$))'
            )


        if (-not $sourceMatch.Success) {

            throw (
                "Jungle Smart could not resolve source PPT path."
            )
        }


        $sourcePath =
            [string]$sourceMatch.Groups["p"].Value


        $sourcePath =
            $sourcePath.Trim().Trim('"').Trim("'")


        if (-not (
            Test-Path `
                -LiteralPath $sourcePath `
                -PathType Leaf
        )) {

            throw "Source PPT not found: $sourcePath"
        }


        $sourceHashBefore =
            (
                Get-FileHash `
                    -LiteralPath $sourcePath `
                    -Algorithm SHA256
            ).Hash
        # --------------------------------------------------------
        # JUNGLE SMART NATIVE ARTIFACT SNAPSHOT V2
        #
        # Snapshot paths BEFORE invoking the isolated worker.
        # --------------------------------------------------------

        $nativeOutputDirectory =
            "C:\AI\work\output"


        $nativeArtifactSnapshot =
            @(
                Get-JungleNativePptArtifactSnapshot `
                    -SourcePath $sourcePath `
                    -OutputDirectory $nativeOutputDirectory
            )


        # --------------------------------------------------------
        # Existing v3.2 SAFE Production Stable bridge.
        #
        # v0.7:
        # Worker Protocol ON -> isolated worker process
        #                       supervised by Watchdog.
        #
        # Worker Protocol OFF -> direct legacy fallback.
        # --------------------------------------------------------

        if ($config.feature_flags.worker_protocol) {

            $workerRunDirectory =
                Join-Path `
                    $context.run_dir `
                    "ppt_safe_worker"


            $workerTimeoutSeconds =
                [int]$config.watchdog.default_timeout_seconds


            $workerPollMilliseconds =
                [int]$config.watchdog.poll_milliseconds


            $workerResult =
                Invoke-JunglePptSafeWorker `
                    -Root $root `
                    -Request $Request `
                    -ConfigPath $configPath `
                    -RunDirectory $workerRunDirectory `
                    -TimeoutSeconds $workerTimeoutSeconds `
                    -PollMilliseconds $workerPollMilliseconds


            Write-Host ""
            Write-Host "----------------------------------------"
            Write-Host " JUNGLE ISOLATED WORKER"
            Write-Host "----------------------------------------"
            Write-Host ""

            Write-Host "Process ID      : $($workerResult.process_id)"
            Write-Host "Completed       : $($workerResult.completed)"
            Write-Host "Timed out       : $($workerResult.timed_out)"
            Write-Host "Killed          : $($workerResult.killed)"
            Write-Host "Exit code       : $($workerResult.exit_code)"
            Write-Host "Duration ms     : $($workerResult.duration_ms)"
            Write-Host ""


            if ($workerResult.timed_out) {

                throw (
                    "PPT SAFE worker timed out and was terminated by Watchdog."
                )
            }


            if (-not $workerResult.completed) {

                throw (
                    "PPT SAFE worker did not complete."
                )
            }


            # ----------------------------------------------------
            # JUNGLE WORKER EXPLICIT RESULT GUARD
            #
            # Worker process exit/completion alone is not success.
            # Worker Protocol V2 result JSON is authoritative.
            # ----------------------------------------------------

            $explicitWorkerResult =
                Assert-JungleWorkerAcceptedResult `
                    -ResultPath $workerResult.result_path


            Write-Host ""
            Write-Host "----------------------------------------"
            Write-Host " JUNGLE WORKER EXPLICIT RESULT"
            Write-Host "----------------------------------------"
            Write-Host ""

            Write-Host "Protocol          : $($explicitWorkerResult.protocol)"
            Write-Host "Status            : $($explicitWorkerResult.status)"
            Write-Host "Pipeline complete : $($explicitWorkerResult.pipeline_complete)"
            Write-Host "Artifact          : $($explicitWorkerResult.artifact_path)"
            Write-Host ""

            Write-Host "STATUS : JUNGLE_WORKER_EXPLICIT_RESULT_PASS"
            Write-Host ""

            if ($workerResult.exit_code -ne 0) {

                $workerError =
                    [string]$workerResult.stderr


                if ([string]::IsNullOrWhiteSpace(
                    $workerError
                )) {

                    $workerError =
                        "No worker stderr was captured."
                }


                throw (
                    "PPT SAFE worker failed. " +
                    $workerError
                )
            }


            $bridgeLog =
                [string]$workerResult.stdout


            if ([string]::IsNullOrWhiteSpace(
                $bridgeLog
            )) {

                throw (
                    "PPT SAFE worker completed but returned no stdout."
                )
            }


            Write-Host "STATUS : JUNGLE_WORKER_EXECUTION_PASS"
            Write-Host ""

            Write-Host $bridgeLog
        }
        else {

            $bridgeLog =
                (
                    & (
                        Join-Path `
                            $root `
                            "executors\ppt_safe_bridge.ps1"
                    ) `
                        -Request $Request `
                        -ConfigPath $configPath *>&1 |
                    Out-String
                )


            Write-Host $bridgeLog
        }

        # --------------------------------------------------------
        # Resolve native artifact.
        #
        # V2:
        # 1. Flexible log parsing
        # 2. Pre/post path snapshot fallback
        #
        # No LastWriteTime dependency.
        # --------------------------------------------------------

        $artifactResolution =
            Resolve-JungleNativePptArtifact `
                -BridgeLog $bridgeLog `
                -SourcePath $sourcePath `
                -OutputDirectory $nativeOutputDirectory `
                -ExistingPaths ([string[]]$nativeArtifactSnapshot)


        $artifactPath =
            [string]$artifactResolution.path


        Write-Host ""
        Write-Host "----------------------------------------"
        Write-Host " JUNGLE NATIVE ARTIFACT RESOLUTION"
        Write-Host "----------------------------------------"
        Write-Host ""

        Write-Host "Method   : $($artifactResolution.method)"
        Write-Host "Artifact : $artifactPath"
        Write-Host ""

        Write-Host "STATUS : JUNGLE_NATIVE_ARTIFACT_RESOLUTION_PASS"
        Write-Host ""

        # --------------------------------------------------------
        # Deterministic G1.
        # --------------------------------------------------------

        if ($config.feature_flags.deterministic_g1) {

            $g1 =
                Test-JunglePptSafeG1 `
                    -SourcePath $sourcePath `
                    -ArtifactPath $artifactPath `
                    -ExpectedSourceHash $sourceHashBefore


            Write-Host ""
            Write-Host "----------------------------------------"
            Write-Host " JUNGLE DETERMINISTIC G1"
            Write-Host "----------------------------------------"
            Write-Host ""

            Write-Host "Source preserved : $($g1.source_preserved)"
            Write-Host "Slide count same : $($g1.slide_count_same)"
            Write-Host "Text same        : $($g1.text_same)"
            Write-Host "Structure same   : $($g1.structure_same)"
            Write-Host "Media count same : $($g1.media_count_same)"
            Write-Host "G1 PASS          : $($g1.pass)"
            Write-Host ""


            if (-not $g1.pass) {

                throw (
                    "Jungle Deterministic G1 failed."
                )
            }


            Write-Host "STATUS : JUNGLE_DETERMINISTIC_G1_PASS"
            Write-Host ""
        }


        # --------------------------------------------------------
        # Changed Scope Detection.
        # --------------------------------------------------------

        if ($config.feature_flags.changed_scope) {

            $scopeDir =
                Join-Path `
                    $context.run_dir `
                    "changed_scope"


            $scope =
                Get-JungleChangedScope `
                    -SourcePath $sourcePath `
                    -ArtifactPath $artifactPath `
                    -WorkDirectory $scopeDir


            Write-Host ""
            Write-Host "----------------------------------------"
            Write-Host " JUNGLE CHANGED SCOPE"
            Write-Host "----------------------------------------"
            Write-Host ""

            Write-Host "Byte identical   : $($scope.byte_identical)"
            Write-Host "Rendered         : $($scope.rendered)"
            Write-Host "Slide count same : $($scope.slide_count_same)"
            Write-Host "Changed count    : $($scope.changed_count)"


            if ($scope.changed_count -gt 0) {

                Write-Host (
                    "Changed slides   : " +
                    ($scope.changed_slides -join ", ")
                )
            }
            else {

                Write-Host "Changed slides   : NONE"
            }


            Write-Host "Scope PASS       : $($scope.pass)"
            Write-Host ""


            if (-not $scope.pass) {

                throw (
                    "Jungle Changed Scope detection failed."
                )
            }


            $scopeRecord =
                [ordered]@{

                    source =
                        $sourcePath

                    artifact =
                        $artifactPath

                    byte_identical =
                        $scope.byte_identical

                    rendered =
                        $scope.rendered

                    slide_count_same =
                        $scope.slide_count_same

                    changed_count =
                        $scope.changed_count

                    changed_slides =
                        @($scope.changed_slides)

                    created_at =
                        (Get-Date).ToString("s")
                }


            $scopeJson =
                $scopeRecord |
                ConvertTo-Json -Depth 5


            $scopeJsonPath =
                Join-Path `
                    $context.run_dir `
                    "changed_scope.json"


            $jsonEncoding =
                New-Object System.Text.UTF8Encoding(
                    $false
                )


            [System.IO.File]::WriteAllText(
                $scopeJsonPath,
                $scopeJson,
                $jsonEncoding
            )


            Write-Host "Scope record     : $scopeJsonPath"
            Write-Host ""
            Write-Host "STATUS : JUNGLE_CHANGED_SCOPE_PASS"

            # ----------------------------------------------------
            # Selective Visual QA Gate
            #
            # No live AI call in v0.6.
            # This stage decides whether visual AI QA is needed
            # and prepares a changed-slide-only request record.
            # ----------------------------------------------------

            if ($config.feature_flags.blind_pairwise_qa) {

                $maxSlidesPerCall =
                    [int]$config.visual_qa.max_slides_per_call


                $qaDecision =
                    Get-JungleSelectiveVisualQaDecision `
                        -ChangedSlides @($scope.changed_slides) `
                        -ByteIdentical ([bool]$scope.byte_identical) `
                        -MaxSlidesPerCall $maxSlidesPerCall


                $qaRequestPath =
                    Join-Path `
                        $context.run_dir `
                        "visual_qa_request.json"


                Write-JungleSelectiveVisualQaRequest `
                    -Decision $qaDecision `
                    -SourcePath $sourcePath `
                    -ArtifactPath $artifactPath `
                    -OutputPath $qaRequestPath |
                Out-Null


                Write-Host ""
                Write-Host "----------------------------------------"
                Write-Host " JUNGLE SELECTIVE VISUAL QA GATE"
                Write-Host "----------------------------------------"
                Write-Host ""

                Write-Host "Visual QA required : $($qaDecision.visual_qa_required)"
                Write-Host "Action             : $($qaDecision.action)"
                Write-Host "Provider           : $($qaDecision.provider)"
                Write-Host "Changed count      : $($qaDecision.changed_count)"


                if ($qaDecision.changed_count -gt 0) {

                    Write-Host (
                        "Changed slides     : " +
                        ($qaDecision.changed_slides -join ", ")
                    )
                }
                else {

                    Write-Host "Changed slides     : NONE"
                }


                Write-Host "Estimated AI calls : $($qaDecision.estimated_ai_calls)"
                Write-Host "Live execution     : $($qaDecision.live_execution)"
                Write-Host "Request record     : $qaRequestPath"
                Write-Host ""


                if ($qaDecision.visual_qa_required) {

                    Write-Host "STATUS : JUNGLE_SELECTIVE_VISUAL_QA_REQUEST_READY"


                    if ($config.visual_qa.live_execution) {

                        [int[]]$liveChangedSlides =
                            @(
                                $scope.changed_slides |
                                ForEach-Object {
                                    [int]$_
                                } |
                                Sort-Object -Unique
                            )


                        $liveQaDirectory =
                            Join-Path `
                                $context.run_dir `
                                "blind_pairwise_qa"


                        $liveQa =
                            Invoke-JungleBlindPairwiseVisualQa `
                                -Root $root `
                                -ChangedSlides $liveChangedSlides `
                                -ScopeDirectory $scopeDir `
                                -ArtifactPath $artifactPath `
                                -RunDirectory $liveQaDirectory `
                                -TimeoutSeconds ([int]$config.visual_qa.timeout_seconds) `
                                -PollMilliseconds ([int]$config.watchdog.poll_milliseconds)


                        Write-Host ""
                        Write-Host "----------------------------------------"
                        Write-Host " JUNGLE LIVE BLIND PAIRWISE QA"
                        Write-Host "----------------------------------------"
                        Write-Host ""


                        foreach (
                            $pairResult in
                            @(
                                $liveQa.results |
                                Sort-Object slide
                            )
                        ) {

                            Write-Host (
                                "Slide " +
                                $pairResult.slide +
                                " : " +
                                $pairResult.comparison +
                                " / " +
                                $pairResult.severity +
                                " / confidence " +
                                $pairResult.confidence
                            )
                        }


                        Write-Host ""
                        Write-Host "Claude calls        : $($liveQa.claude_calls)"
                        Write-Host "Before better       : $($liveQa.before_better_count)"
                        Write-Host "Critical regression : $($liveQa.critical_regressions)"
                        Write-Host "Major regression    : $($liveQa.major_regressions)"
                        Write-Host "Minor regression    : $($liveQa.minor_regressions)"
                        Write-Host "Visual Gate PASS    : $($liveQa.pass)"
                        Write-Host ""


                        if (-not $liveQa.pass) {

                            throw (
                                "Jungle Blind Pairwise Visual Regression Gate failed."
                            )
                        }


                        Write-Host "STATUS : JUNGLE_LIVE_BLIND_PAIRWISE_QA_PASS"
                        Write-Host ""
                    }
                }
                else {

                    Write-Host "STATUS : JUNGLE_SELECTIVE_VISUAL_QA_SKIPPED"
                }


                Write-Host ""
            }
            Write-Host ""
        }
    }


    default {

        throw (
            "Execution route not activated yet: " +
            $route.route
        )
    }
}