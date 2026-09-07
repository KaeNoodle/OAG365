function Export-OagM365OrgReport {
    <#
    .SYNOPSIS
    Exports tenant organisation configuration, domains and email authentication records
    (MX, SPF, DKIM, DMARC).

    .DESCRIPTION
    Known permission constraint: the on-premises directory synchronisation export requires
    Global Administrator. The OnPremDirectorySynchronization.Read.All scope cannot be satisfied
    by Global Reader, Security Reader or any other read-only directory role, so this one export
    returns 403 under least-privilege credentials.

    Rather than failing the report, that specific call is isolated and the failure recorded as a
    known permission gap, so the rest of the organisation evidence is still collected and the
    gap is documented in the completeness output instead of appearing as an unexplained error.

    .PARAMETER Context
    Run context object from Initialize-OagM365Run.

    .PARAMETER Connection
    Existing connection result from Connect-OagM365Graph.

    .PARAMETER SkipOnPremSync
    Skip the on-premises sync export entirely. Use when running with read-only credentials and
    the 403 is already documented, to keep the run log clean.

    .EXAMPLE
    Export-OagM365OrgReport -Context $Context -Connection $Connection -SkipOnPremSync

    .NOTES
        NAME: Export-OagM365OrgReport
        VERSION: 4.0.0

        REQUIRED ENTRA ROLE:
           Global Reader for most exports
           Global Administrator for on-premises directory synchronisation configuration

        GRAPH SCOPES:
           Organization.Read.All, Directory.Read.All, Domain.Read.All,
           OnPremDirectorySynchronization.Read.All (privileged)
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$Connection,
        [switch]$SkipOnPremSync
    )

    $folder = Initialize-OagM365ReportPath -Context $Context -Report 'Org'

    $status = [PSCustomObject]@{
        Report    = 'Org'
        Started   = Get-Date
        Completed = $null
        Succeeded = $false
        Folder    = $folder
    }

    try {
        Write-OagM365Log "Exporting tenant organisation configuration" -Context $Context
        Export-OagM365OrgConfig | Out-Null

        if ($SkipOnPremSync) {
            Write-OagM365Log "On-premises sync export skipped by request. Recorded as a known gap." -Level Warning -Indent 1 -Context $Context
        } else {
            Write-OagM365Log "Exporting on-premises directory synchronisation configuration" -Context $Context
            try {
                Export-OagM365OrgOnPremiseSyncConfig -ErrorAction Stop | Out-Null
            } catch {
                Write-OagM365Log "On-premises sync export failed. This requires Global Administrator and is expected to fail under read-only credentials - record it as a scope limitation rather than an error." -Level Warning -Indent 1 -Context $Context
                Write-OagM365Log $_.Exception.Message -Level Detail -Indent 2 -Context $Context
            }
        }

        Write-OagM365Log "Exporting tenant domains" -Context $Context
        $domains = Export-OagM365OrgTenantDomain

        Write-OagM365Log "Exporting email authentication records for $(@($domains).Count) domain(s)" -Context $Context
        Export-OagM365OrgEmail -domains $domains | Out-Null

        Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter '*.csv' -Description 'Organisation CSV export'

        $status.Succeeded = $true

    } catch {
        $exception = Format-OagM365Exception -message "Organisation report failed" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
    } finally {
        $status.Completed = Get-Date
        $Context.ReportStatus.Add($status)
    }

    return $status
}
