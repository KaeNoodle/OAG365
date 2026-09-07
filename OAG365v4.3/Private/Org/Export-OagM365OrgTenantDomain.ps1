function Export-OagM365OrgTenantDomain {
    <#
    .SYNOPSIS
    Get domains configured within M365 tenant 

    .DESCRIPTION    
    Get domains configured within M365 tenant 

    .EXAMPLE
    Export-OagM365OrgTenantDomain

    .NOTES
        NAME: Export-OagM365OrgTenantDomain
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDomain

        CHANGELOG:    
    #>

    try {
        Write-Host "`nM365 tenant domains"

        # Get organization domains
        $getMgDomain = Get-MgDomain -All
        $getMgDomainCount = $getMgDomain.count
        Write-Host " - Found $getMgDomainCount M365 domains"

        $getMgDomain | Select-Object -Property Id, IsAdminManaged, IsDefault, IsInitial, IsRoot, IsVerified, AuthenticationType, AvailabilityStatus, DomainNameReferences, FederationConfiguration, PasswordNotificationWindowInDays, PasswordValidityPeriodInDays | Export-CSV $script:exportTarget.domains -NoTypeInformation
        Write-Host " - Exported M365 domains to CSV $($script:exportTarget.domains)" -ForegroundColor Green

        return $getMgDomain
    } catch {
        $exception = Format-OagM365Exception -message "Failed getting M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
