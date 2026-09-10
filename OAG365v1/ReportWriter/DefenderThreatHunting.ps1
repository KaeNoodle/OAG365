function threatHuntReportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Microsoft Defender advanced hunting evidence: software vulnerabilities,
    agent health and device software inventory.

    Runs three KQL queries against v1.0/security/runHuntingQuery.

    LOGIC
    Makes sure a run exists, then creates the Defender folder and sets export targets.
    Checks whether ThreatHunting.Read.All was actually granted and warns up front if not,
      rather than running three queries that will all be rejected for the same reason.
    Gets the query definitions and runs each one through threatHuntQueryRun.

    PARAMETERS
    None.

    RUNNING CONTEXT
    Called by  : runMe.ps1 -report TH, or run directly
    Calls      : runEnsure, reportPathInitialize, threatHuntQueryGet, threatHuntQueryRun,
                 logWrite, exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Role  : Security Reader. Global Reader alone does not cover Defender hunting - the
            two roles are not subsets of each other.
    Scope : ThreatHunting.Read.All, which requires tenant administrator consent

    Where consent has not been granted this report produces nothing. Consent is granted
    once per tenant, not per user, so it only needs doing once - but it needs doing
    before the audit, not during it.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    runEnsure -report 'TH'
    reportPathInitialize -report 'TH' | Out-Null
    $started = Get-Date

    try {
        $notGranted = $script:run.scopes |
                      Where-Object { $_.scope -eq 'ThreatHunting.Read.All' -and -not $_.granted }
        if ($notGranted) {
            logWrite "ThreatHunting.Read.All was requested but not granted. These queries will fail until an administrator consents." -level Warning -indent 1
        }

        foreach ($definition in (threatHuntQueryGet)) {
            threatHuntQueryRun -definition $definition | Out-Null
        }

        $script:run.status += [PSCustomObject]@{ report = 'TH'; started = $started; completed = Get-Date; succeeded = $true }
        return $true

    } catch {
        logWrite (exceptionFormat -message "Threat hunting report failed" -exception $_) -level Error
        $script:run.status += [PSCustomObject]@{ report = 'TH'; started = $started; completed = Get-Date; succeeded = $false }
        return $false
    }
}
