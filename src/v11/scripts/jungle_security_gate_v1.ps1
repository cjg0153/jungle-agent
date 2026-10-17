function Invoke-JungleSecurityGate {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Prompt)

    $router = $Prompt
    $findings = New-Object System.Collections.Generic.List[string]
    $block = $false

    # Minimize local path exposure to Jev.
    $router = [regex]::Replace(
        $router,
        '(?i)[A-Z]:\\[^"''\r\n]+?\.(pptx?|docx?|xlsx?|xlsm|xls|pdf|html?|csv|txt|md|hwp|hwpx)',
        {
            param($m)
            $ext = [System.IO.Path]::GetExtension($m.Value).ToLowerInvariant()
            $findings.Add("LOCAL_PATH_MINIMIZED:$ext")
            return "[LOCAL_FILE:$ext]"
        }
    )

    # Never route explicit credentials/private keys.
    if ($router -match '(?s)-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----') {
        $findings.Add("PRIVATE_KEY")
        $block = $true
    }

    if ($router -match '(?i)\b(password|passwd|pwd|client_secret|api[_-]?key|access[_-]?token|refresh[_-]?token)\s*[:=]\s*\S{4,}') {
        $findings.Add("EXPLICIT_SECRET")
        $block = $true
    }

    if ($router -match '(?i)\bBearer\s+[A-Za-z0-9._~+/\-=]{16,}') {
        $findings.Add("BEARER_TOKEN")
        $block = $true
    }

    $tokenPatterns = @(
        '(?i)\bsk-[A-Za-z0-9_-]{16,}\b',
        '(?i)\bgh[pousr]_[A-Za-z0-9]{20,}\b',
        '(?i)\bAIza[0-9A-Za-z\-_]{20,}\b'
    )

    foreach ($pat in $tokenPatterns) {
        if ($router -match $pat) {
            $findings.Add("TOKEN_REDACTED")
            $router = [regex]::Replace($router, $pat, '[REDACTED_SECRET]')
        }
    }

    if ($router -match '\b\d{6}-[1-4]\d{6}\b') {
        $findings.Add("RRN_REDACTED")
        $router = [regex]::Replace($router, '\b\d{6}-[1-4]\d{6}\b', '[REDACTED_RRN]')
    }

    if ($router -match '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b') {
        $findings.Add("EMAIL_REDACTED")
        $router = [regex]::Replace(
            $router,
            '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b',
            '[REDACTED_EMAIL]'
        )
    }

    if ($router -match '\b01[016789][-\s]?\d{3,4}[-\s]?\d{4}\b') {
        $findings.Add("PHONE_REDACTED")
        $router = [regex]::Replace(
            $router,
            '\b01[016789][-\s]?\d{3,4}[-\s]?\d{4}\b',
            '[REDACTED_PHONE]'
        )
    }

    if ($router -match '(?<!\d)\d{10,16}(?!\d)') {
        $findings.Add("LONG_ID_REDACTED")
        $router = [regex]::Replace(
            $router,
            '(?<!\d)\d{10,16}(?!\d)',
            '[REDACTED_LONG_ID]'
        )
    }

    # Jev needs routing intent, not a huge prompt body.
    if ($router.Length -gt 1600) {
        $router = $router.Substring(0,1600) + " [TRUNCATED_FOR_ROUTING]"
        $findings.Add("ROUTER_PROMPT_TRUNCATED")
    }

    [PSCustomObject]@{
        Block        = $block
        RouterPrompt = $router
        FindingCount = $findings.Count
        Findings     = @($findings)
        Version      = "SECURITY_GATE_V1"
    }
}
