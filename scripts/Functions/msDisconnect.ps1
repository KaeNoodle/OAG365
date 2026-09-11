function msDisconnect {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Disconnects from Microsoft Graph and Exchange Online. Replaces the two separate
    disconnect functions in the original scripts.

    LOGIC
    Attempts each service independently inside its own try/catch, so a failure to
      disconnect one does not leave the other session open.
    Checks a session exists before disconnecting, so calling this with nothing
      connected is harmless.

    PARAMETERS
    -service (optional, one of: Graph, ExchangeOnline, All) defaults to All

    RUNNING CONTEXT
    Called by  : runMe.ps1 in its finally block
    Calls      : logWrite

    CMLETS/PERMISSIONS/SCOPES
    Disconnect-MgGraph
    Disconnect-ExchangeOnline

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [ValidateSet('Graph', 'ExchangeOnline', 'All')][string[]]$service = 'All'
    )

    if ($service -contains 'All' -or $service -contains 'Graph') {
        try {
            # A process that only ever connected to Exchange Online has no Graph cmdlets
            # loaded at all, so check the cmdlet exists before calling it.
            if ((Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue) -and
                (Get-MgContext -ErrorAction SilentlyContinue)) {
                Disconnect-MgGraph -ErrorAction Stop | Out-Null
                logWrite "Disconnected from Microsoft Graph" -level Detail -indent 1
            }
        } catch {
            logWrite "Failed disconnecting from Graph: $($_.Exception.Message)" -level Warning -indent 1
        }
    }

    if ($service -contains 'All' -or $service -contains 'ExchangeOnline') {
        try {
            if (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue) {
                # Same broken lookup as in the connect function: if it throws, fall back to
                # the generated cmdlet module, whose presence means there is a session to
                # close.
                $hasSession = $false
                try   { $hasSession = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue) }
                catch { $hasSession = [bool](Get-Module -Name 'tmpEXO_*') }

                if ($hasSession) {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                    logWrite "Disconnected from Exchange Online" -level Detail -indent 1
                }
            }
        } catch {
            logWrite "Failed disconnecting from Exchange Online: $($_.Exception.Message)" -level Warning -indent 1
        }
    }
}
