function Get-JungleCodexProvider {

    return [pscustomobject]@{

        name =
            "CODEX"

        role =
            "ANALYSIS_AND_QA"

        execution =
            "EXTERNAL"

        enabled =
            $true
    }
}