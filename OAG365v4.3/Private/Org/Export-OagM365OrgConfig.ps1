function Export-OagM365OrgConfig {
    <#
    .SYNOPSIS
    Get M365 tenant settings 

    .DESCRIPTION    
    Get M365 tenant settings 

    .EXAMPLE
    Export-OagM365OrgConfig

    .NOTES
        NAME: Export-OagM365OrgConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgOrganization

        CHANGELOG:    
    #>

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
        $exception = Format-OagM365Exception -message "Failed getting M365 organization details" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
