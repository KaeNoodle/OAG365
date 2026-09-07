function capReportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Entra conditional access policies, authentication method policies and
    authentication strength policies.

    Produces one HTML report per conditional access policy plus a CSV policy summary.

    LOGIC
    Makes sure a run exists, so this report can be called on its own without the main
      run file having been used.
    Creates the ConditionalAccess output folder and sets the export target paths.
    Runs the three exports in order.
    Registers the HTML and CSV files afterwards, because the ported functions write them
      directly rather than through exportWrite.
    Records success or failure on the run so the summary can report it.

    PARAMETERS
    None. Reads $script:run and $script:exportTarget.

    RUNNING CONTEXT
    Called by  : OAG-MainRunFile.ps1 -report CAP, or run directly
    Calls      : runEnsure, reportPathInitialize, capPolicyExport, capAuthMethodExport,
                 capAuthStrengthExport, exportRegister, logWrite, exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Role  : Global Reader
    Scopes: Policy.Read.All, Policy.Read.ConditionalAccess,
            Policy.Read.AuthenticationMethod, AuthenticationContext.Read.All,
            RoleManagement.Read.Directory, Application.Read.All, Directory.Read.All

    RoleManagement.Read.Directory and Application.Read.All are needed not for the
    policies themselves but to turn the GUIDs inside them into readable names. Without
    them the HTML renders raw GUIDs, which is complete but not usable as evidence.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    runEnsure -report 'CAP'
    $folder = reportPathInitialize -report 'CAP'
    $started = Get-Date

    try {
        logWrite "Exporting conditional access policies"
        capPolicyExport -OutputFolder $folder | Out-Null

        logWrite "Exporting authentication method policies"
        capAuthMethodExport | Out-Null

        logWrite "Exporting authentication strength policies"
        capAuthStrengthExport | Out-Null

        exportRegister -folder $folder -filter 'CAP_*.html' -description 'Conditional access policy report'
        exportRegister -folder $folder -filter '*.csv'      -description 'Conditional access CSV export'

        $script:run.status += [PSCustomObject]@{ report = 'CAP'; started = $started; completed = Get-Date; succeeded = $true }
        return $true

    } catch {
        logWrite (exceptionFormat -message "Conditional access report failed" -exception $_) -level Error
        $script:run.status += [PSCustomObject]@{ report = 'CAP'; started = $started; completed = Get-Date; succeeded = $false }
        return $false
    }
}
