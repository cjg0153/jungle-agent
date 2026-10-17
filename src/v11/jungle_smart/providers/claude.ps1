function Get-JungleClaudeProvider {

    return [pscustomobject]@{

        name =
            "CLAUDE"

        role =
            "PPT_VISUAL_SPECIALIST"

        execution =
            "EXTERNAL"

        enabled =
            $true
    }
}