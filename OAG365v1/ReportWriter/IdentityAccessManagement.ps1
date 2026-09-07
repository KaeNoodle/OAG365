function iamReportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Entra identity and access management evidence: users, devices, service
    principals, authentication methods, administrative units and PIM role assignments.

    The heaviest of the five reports.

    LOGIC
    Makes sure a run exists, then creates the IAM folder and sets export targets.
    Runs in three stages, matching the original script, so a partial failure is
      attributable to a stage rather than the whole report.
    Stage 3 has an ordering dependency: iamPimRoleGet populates
      $script:pimMgmtPolicyRules as a side effect, so it must run before those rules are
      written out. This is inherited from the original and is called out because it is
      not obvious from reading the calls.

    PARAMETERS
    None. Reads $script:run and $script:exportTarget.

    RUNNING CONTEXT
    Called by  : OAG-MainRunFile.ps1 -report IAM, or run directly
    Calls      : runEnsure, reportPathInitialize, iamUserExport, iamDeviceExport,
                 iamServicePrincipalExport, iamAuthMethodExport, iamAdminUnitExport,
                 iamPimRoleGet, exportWrite, exportRegister, logWrite, exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Roles : Global Reader plus Reports Reader
    Scopes: Directory.Read.All, DeviceManagementApps.Read.All, Device.Read.All,
            Application.Read.All, AuditLog.Read.All, RoleManagement.Read.Directory,
            PrivilegedEligibilitySchedule.Read.AzureADGroup,
            PrivilegedAssignmentSchedule.Read.AzureADGroup

    AuditLog.Read.All needs Reports Reader on top of Global Reader. Without it the
    sign-in activity columns come back blank rather than the call failing, so it is a
    silent gap. Check completeness.csv if user activity data is missing.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    runEnsure -report 'IAM'
    $folder = reportPathInitialize -report 'IAM'
    $started = Get-Date

    try {
        logWrite "STAGE 1 of 3: users, devices and service principals"
        iamUserExport | Out-Null
        iamDeviceExport | Out-Null
        iamServicePrincipalExport | Out-Null

        logWrite "STAGE 2 of 3: user authentication methods"
        iamAuthMethodExport | Out-Null

        logWrite "STAGE 3 of 3: administrative units and PIM role membership"
        iamAdminUnitExport | Out-Null

        # Must run before pimMgmtPolicyRules is written - see LOGIC above.
        iamPimRoleGet | exportWrite -path $script:exportTarget.builtInRoleMembers `
                                    -description 'Built-in role assignments'

        $script:pimMgmtPolicyRules | exportWrite -path $script:exportTarget.pimMgmtPolicyRules `
                                                 -description 'PIM role management policy rules'

        exportRegister -folder $folder -filter '*.csv' -description 'IAM CSV export'

        $script:run.status += [PSCustomObject]@{ report = 'IAM'; started = $started; completed = Get-Date; succeeded = $true }
        return $true

    } catch {
        logWrite (exceptionFormat -message "IAM report failed" -exception $_) -level Error
        $script:run.status += [PSCustomObject]@{ report = 'IAM'; started = $started; completed = Get-Date; succeeded = $false }
        return $false
    }
}
