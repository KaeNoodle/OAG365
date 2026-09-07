function Disconnect-OagM365Service {
    <#
    .SYNOPSIS
    Disconnects from Microsoft Graph and / or Exchange Online.

    .DESCRIPTION
    Replaces the separate Disconnect-Oag365-FromGraph and Disconnect-Oag365Dfo-FromExchangeOnline
    functions. Disconnecting is always attempted for every connected service so a failure in one
    does not leave the other session open.

    .PARAMETER Service
    Which service(s) to disconnect. Defaults to All.

    .PARAMETER Context
    Run context object, used for logging.

    .EXAMPLE
    Disconnect-OagM365Service -Context $Context

    .NOTES
        NAME: Disconnect-OagM365Service
        VERSION: 2.0
    #>
    [CmdletBinding()]
    param (
        [ValidateSet('Graph', 'ExchangeOnline', 'All')]
        [string[]]$Service = 'All',

        [object]$Context
    )

    if ($Service -contains 'All' -or $Service -contains 'Graph') {
        try {
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -ErrorAction Stop | Out-Null
                Write-OagM365Log "Disconnected from Microsoft Graph" -Level Detail -Indent 1 -Context $Context
            }
        } catch {
            Write-OagM365Log "Failed disconnecting from Microsoft Graph: $($_.Exception.Message)" -Level Warning -Indent 1 -Context $Context
        }
    }

    if ($Service -contains 'All' -or $Service -contains 'ExchangeOnline') {
        try {
            if (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue) {
                if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                    Write-OagM365Log "Disconnected from Exchange Online" -Level Detail -Indent 1 -Context $Context
                }
            }
        } catch {
            Write-OagM365Log "Failed disconnecting from Exchange Online: $($_.Exception.Message)" -Level Warning -Indent 1 -Context $Context
        }
    }
}
