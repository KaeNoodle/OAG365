function Export-OagM365OrgEmail {
    <#
    .SYNOPSIS
    Get email authentication DNS records for M365 tenant domains 

    .DESCRIPTION    
    Get email authentication DNS records for M365 tenant domains 

    .PARAMETER domains
    Iterable object containing domains retrieved from Get-MgDomains

    .EXAMPLE
    Export-OagM365OrgEmail -domains $object

    .NOTES
        NAME: Export-OagM365OrgEmail
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDomain
           Get-OagM365OrgDomainMxRecord
           Export-OagM365OrgTenantDomainSpf
           Get-OagM365OrgDomainDkimRecord
           Get-OagM365OrgDomainDmarcRecord

        CHANGELOG:    
    #>
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$domains
    )

    try {
        Write-Host "`nEmail authentication DNS settings for M365 domains"

        # Get MX records to confirm a mail server is available
        $getMxDnsRecord = Get-OagM365OrgDomainMxRecord -getMgDomains $domains

        # Get SPF DNS record for each M365 domain
        Write-Host "`nRetrieving SPF records for M365 domains"
        $getSpfDnsRecord = Export-OagM365OrgTenantDomainSpf -getMgDomains $domains
        $getSpfDnsRecordCount = $getSpfDnsRecord.count
        Write-Verbose " - SPF records for $getSpfDnsRecordCount M365 domains"

        ($getMxDnsRecord + $getSpfDnsRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainSpf -NoTypeInformation
        Write-Host " - Exported SPF records for M365 domains to CSV $($script:exportTarget.domainSpf)" -ForegroundColor Green

        # Get DKIM DNS Record for each M365 domain
        Write-Host "`nRetrieving DKIM records for M365 domains"
        $getDkimSigningRecord = Get-OagM365OrgDomainDkimRecord -getMgDomains $domains 
        $getDkimSigningRecordCount = $getDkimSigningRecord.count
        Write-Verbose " - Found DKIM config for $getDkimSigningRecordCount domains"

        ($getMxDnsRecord + $getDkimSigningRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDkim -NoTypeInformation
        Write-Host " - Exported M365 DKIM records for M365 domains to CSV $($script:exportTarget.domainDkim)" -ForegroundColor Green

        # Get DMARC DNS Record for each M365 domain
        Write-Host "`nRetrieving DMARC records for M365 domains"
        $getDmarcDnsRecord += Get-OagM365OrgDomainDmarcRecord -getMgDomains $domains
        $getDmarcDnsRecordCount = $getDmarcDnsRecord.count
        Write-Verbose " - DMARC records for $getDmarcDnsRecordCount M365 domains"

        ($getMxDnsRecord + $getDmarcDnsRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDmarc -NoTypeInformation
        Write-Host " - Exported DMARC records for M365 domains to CSV $($script:exportTarget.domainDmarc)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed getting email authentication DNS records for M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
