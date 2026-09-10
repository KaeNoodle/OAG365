function orgEmailExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Get email authentication DNS records for M365 tenant domains

    LOGIC
    Writes the result to CSV at the path held in $script:exportTarget.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.
    Returns $true on success and $false on failure.

    PARAMETERS
    -domains (required)

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : exceptionFormat, orgDomainDkimGet, orgDomainDmarcGet, orgDomainMxGet, orgDomainSpfGet

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$domains
    )

    try {
        Write-Host "`nEmail authentication DNS settings for M365 domains"

        # Get MX records to confirm a mail server is available
        $getMxDnsRecord = orgDomainMxGet -getMgDomains $domains

        # Get SPF DNS record for each M365 domain
        Write-Host "`nRetrieving SPF records for M365 domains"
        $getSpfDnsRecord = orgDomainSpfGet -getMgDomains $domains
        $getSpfDnsRecordCount = $getSpfDnsRecord.count
        Write-Verbose " - SPF records for $getSpfDnsRecordCount M365 domains"

        (@($getMxDnsRecord) + @($getSpfDnsRecord)) | Select-Object -Property * | Export-CSV $script:exportTarget.domainSpf -NoTypeInformation
        Write-Host " - Exported SPF records for M365 domains to CSV $($script:exportTarget.domainSpf)" -ForegroundColor Green

        # Get DKIM DNS Record for each M365 domain
        Write-Host "`nRetrieving DKIM records for M365 domains"
        $getDkimSigningRecord = orgDomainDkimGet -getMgDomains $domains 
        $getDkimSigningRecordCount = $getDkimSigningRecord.count
        Write-Verbose " - Found DKIM config for $getDkimSigningRecordCount domains"

        (@($getMxDnsRecord) + @($getDkimSigningRecord)) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDkim -NoTypeInformation
        Write-Host " - Exported M365 DKIM records for M365 domains to CSV $($script:exportTarget.domainDkim)" -ForegroundColor Green

        # Get DMARC DNS Record for each M365 domain
        Write-Host "`nRetrieving DMARC records for M365 domains"
        $getDmarcDnsRecord += orgDomainDmarcGet -getMgDomains $domains
        $getDmarcDnsRecordCount = $getDmarcDnsRecord.count
        Write-Verbose " - DMARC records for $getDmarcDnsRecordCount M365 domains"

        (@($getMxDnsRecord) + @($getDmarcDnsRecord)) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDmarc -NoTypeInformation
        Write-Host " - Exported DMARC records for M365 domains to CSV $($script:exportTarget.domainDmarc)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed getting email authentication DNS records for M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function orgTenantDomainExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Get domains configured within M365 tenant

    LOGIC
    Retrieves data using Get-MgDomain.
    Writes the result to CSV at the path held in $script:exportTarget.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    None. Reads its inputs from module scope.

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDomain

    --------------------------------------------------------------------------------#>


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
        $exception = exceptionFormat -message "Failed getting M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
