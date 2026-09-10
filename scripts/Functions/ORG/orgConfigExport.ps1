function orgConfigExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Get M365 tenant settings

    LOGIC
    Retrieves data using Get-MgOrganization.
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
    Get-MgOrganization

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nM365 tenant organization settings"

        # Get organization details
        $getMgOrganization = Get-MgOrganization -All
        $getMgOrganizationCount = $getMgOrganizationCount.count
        Write-Host " - Found M365 organizations data"

        # Export to CSV
        $getMgOrganization | Select-Object -Property DisplayName, CountryLetterCode, CreatedDateTime, Id, MobileDeviceManagementAuthority, OnPremisesLastSyncDateTime, OnPremisesSyncEnabled, TenantType, @{l='TechnicalNotificationMails'; e={$_.TechnicalNotificationMails -join '; ' }}  | Export-CSV $script:exportTarget.organisations -NoTypeInformation
        Write-Host " - Exported M365 organizations to CSV $($script:exportTarget.organisations)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed getting M365 organization details" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function orgOnPremSyncExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Get details of Entra ID sync with on-premises AD

    LOGIC
    Runs 1 loop over the returned collection.
    Retrieves data using Get-MgDirectoryOnPremiseSynchronization.
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
    Get-MgDirectoryOnPremiseSynchronization

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nM365 tenant onPremises AD sync settings"

        # Get onPremise sync status
        $getMgDirectoryOnPremiseSync = Get-MgDirectoryOnPremiseSynchronization 
        $getMgDirectoryOnPremiseSyncCount = $getMgDirectoryOnPremiseSync.count
        Write-Host " - Found $getMgDirectoryOnPremiseSyncCount M365 onPremises sync settings"

        # Export to CSV
        $getMgDirectoryOnPremiseSync | ForEach-Object {
            $onPremisesSync = $_.Features
            $onPremisesSync | Add-Member -MemberType NoteProperty -Name Id -Value $_.Id
            $onPremisesSync
        } | Select-Object -Property * -ExcludeProperty AdditionalProperties | Export-CSV $script:exportTarget.onPremisesSync -NoTypeInformation
        Write-Host " - Exported M365 onPremises sync settings to CSV $($script:exportTarget.onPremisesSync)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed getting M365 directory sync config" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
