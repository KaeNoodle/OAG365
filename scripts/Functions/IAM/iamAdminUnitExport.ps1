function iamAdminUnitExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export any administrative units used to manage role allocations

    LOGIC
    Retrieves data using Get-MgDirectoryAdministrativeUnit.
    Writes the result to CSV at the path held in $script:exportTarget.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.
    Returns $true on success and $false on failure.

    PARAMETERS
    None. Reads its inputs from module scope.

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDirectoryAdministrativeUnit

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "Entra Administrative Units"

        # Get all Entra role definitions to resolve IDs
        $adminUnits =  Get-MgDirectoryAdministrativeUnit -All
        Write-Host " - Found $(($adminUnits | Measure-Object).count) admin units"

        # Export to CSV
        $adminUnits | Select-Object -Property Id, DisplayName, Description, IsMemberManagementRestricted, MembershipRule, MembershipRuleProcessingState, MembershipType, Visibility, Extensions, @{l='AdditionalProperties'; e={$_.AdditionalProperties | ConvertTo-Json }} | Export-CSV $script:exportTarget.roleAdminUnits -NoTypeInformation
        Write-Host " - Exported M365 admin units to CSV $($script:exportTarget.roleAdminUnits)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving admin units" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
