function orgReportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports tenant organisation configuration, domains and email authentication records
    (MX, SPF, DKIM and DMARC).

    LOGIC
    Makes sure a run exists, then creates the Organisation folder and sets export targets.
    Isolates the on-premises sync export in its own try/catch. That single export needs
      Global Administrator and returns 403 under read-only credentials, so isolating it
      means the rest of the organisation evidence is still collected and the gap is
      recorded as a known permission limitation rather than an unexplained error.
    Exports domains, then passes them to the email record lookups.

    The MX, SPF, DKIM and DMARC lookups are public DNS queries and need no tenant
    permission at all.

    PARAMETERS
    -skipOnPremSync (optional) skip the export that needs Global Administrator, to keep
      the run log clean where the gap is already documented

    RUNNING CONTEXT
    Called by  : OAG-MainRunFile.ps1 -report ORG, or run directly
    Calls      : runEnsure, reportPathInitialize, orgConfigExport, orgOnPremSyncExport,
                 orgTenantDomainExport, orgEmailExport, exportRegister, logWrite,
                 exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Role  : Global Reader, except on-premises sync which needs Global Administrator
    Scopes: Organization.Read.All, Directory.Read.All, Domain.Read.All,
            OnPremDirectorySynchronization.Read.All (privileged)

    No read-only role satisfies OnPremDirectorySynchronization.Read.All. See
    Documentation\Permissions.md for the two options.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [switch]$skipOnPremSync
    )

    runEnsure -report 'ORG'
    $folder = reportPathInitialize -report 'ORG'
    $started = Get-Date

    try {
        logWrite "Exporting tenant organisation configuration"
        orgConfigExport | Out-Null

        if ($skipOnPremSync) {
            logWrite "On-premises sync export skipped by request. Recorded as a known gap." -level Warning -indent 1
        } else {
            logWrite "Exporting on-premises directory synchronisation configuration"
            try {
                orgOnPremSyncExport -ErrorAction Stop | Out-Null
            } catch {
                logWrite "On-premises sync export failed. This needs Global Administrator and is expected to fail under read-only credentials - record it as a scope limitation, not an error." -level Warning -indent 1
                logWrite $_.Exception.Message -level Detail -indent 2
            }
        }

        logWrite "Exporting tenant domains"
        $domains = orgTenantDomainExport

        logWrite "Exporting email authentication records for $(@($domains).Count) domain(s)"
        orgEmailExport -domains $domains | Out-Null

        exportRegister -folder $folder -filter '*.csv' -description 'Organisation CSV export'

        $script:run.status += [PSCustomObject]@{ report = 'ORG'; started = $started; completed = Get-Date; succeeded = $true }
        return $true

    } catch {
        logWrite (exceptionFormat -message "Organisation report failed" -exception $_) -level Error
        $script:run.status += [PSCustomObject]@{ report = 'ORG'; started = $started; completed = Get-Date; succeeded = $false }
        return $false
    }
}
