# JUNGLE_SMART_REFLOW_CONTRACT_V1
#
# Jungle defines boundaries.
# Claude owns visual judgment.
#
# REFLOW V1:
# - text content locked
# - slide count locked
# - slide order locked
# - source overwrite forbidden
# - group/ungroup forbidden
# - geometry changes allowed
# - z-order changes allowed
# - decorative changes allowed
# - Hard Gate unchanged

function Get-JungleReflowContract {

    return [pscustomobject]@{

        protocol =
            "JUNGLE_SMART_REFLOW_CONTRACT_V1"

        mode =
            "REFLOW"

        provider =
            "CLAUDE"

        text_content_locked =
            $true

        slide_count_locked =
            $true

        slide_order_locked =
            $true

        source_overwrite_allowed =
            $false

        group_ungroup_allowed =
            $false

        geometry_change_allowed =
            $true

        z_order_change_allowed =
            $true

        decorative_change_allowed =
            $true

        hard_gate_required =
            $true

        hard_gate_semantics =
            "UNCHANGED"
    }
}