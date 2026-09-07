function Export-OagM365ThreatHuntReport {
    <#
    .SYNOPSIS
    Exports Microsoft Defender advanced hunting reports: vulnerabilities, agent health and
    device software inventory.

    .DESCRIPTION
    Runs three KQL advanced hunting queries against v1.0/security/runHuntingQuery.

    Known permission constraint: ThreatHunting.Read.All requires tenant administrator consent.
    Where consent has not been granted this report produces nothing, and the failure is recorded
    explicitly as a permission gap rather than a generic error so it can be documented.

    .PARAMETER Context
    Run context object from Initialize-OagM365Run.

    .PARAMETER Connection
    Existing connection result from Connect-OagM365Graph.

    .EXAMPLE
    Export-OagM365ThreatHuntReport -Context $Context -Connection $Connection

    .NOTES
        NAME: Export-OagM365ThreatHuntReport
        VERSION: 4.0.0

        REQUIRED ENTRA ROLE:
           Security Reader (Global Reader alone does not cover Defender hunting)

        GRAPH SCOPES:
           ThreatHunting.Read.All (requires admin consent)
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$Connection
    )

    $folder = Initialize-OagM365ReportPath -Context $Context -Report 'ThreatHunt'

    $status = [PSCustomObject]@{
        Report    = 'ThreatHunt'
        Started   = Get-Date
        Completed = $null
        Succeeded = $false
        Folder    = $folder
    }

    try {
        # Fail early with a clear message rather than running three queries that will all be
        # rejected for the same reason.
        if ($Connection.ScopesMissing -contains 'ThreatHunting.Read.All') {
            Write-OagM365Log "ThreatHunting.Read.All was requested but not granted. Advanced hunting queries will fail until an administrator consents to this scope." -Level Warning -Indent 1 -Context $Context
        }

        $definitions = Get-OagM365ThreatHuntQueryDefinition
        foreach ($definition in $definitions) {
            Invoke-OagM365HuntingQuery -Definition $definition -Context $Context | Out-Null
        }

        $status.Succeeded = $true

    } catch {
        $exception = Format-OagM365Exception -message "Threat hunting report failed" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
    } finally {
        $status.Completed = Get-Date
        $Context.ReportStatus.Add($status)
    }

    return $status
}
