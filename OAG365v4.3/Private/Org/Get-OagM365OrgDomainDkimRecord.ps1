function Get-OagM365OrgDomainDkimRecord {
    <#
    .SYNOPSIS
    Returns DKIM DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns DKIM DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    Get-OagM365OrgDomainDkimRecord -getMgDomains $object

    .NOTES
        NAME: Get-OagM365OrgDomainDkimRecord
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>
    
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $DkimSelectors = @(
        "selector1",
        "selector2"
    )

    $getMgDomains | ForEach-Object { 
        try {
            foreach ($DkimSelector in $DkimSelectors) {
                $dnsRecordName = "$($DkimSelector)._domainkey.$($_.Id)"
                $records = Resolve-DnsName -Type TXT -Name $dnsRecordName -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=DKIM1"; } 
                if ($records) {
                    Write-Host " - $($dnsRecordName) found" -ForegroundColor Green
                    foreach ($record in $records) {
                        [PSCustomObject]@{
                            "Domain"  = $record.Name
                            "Type"    = $record.Type
                            "TTL"     = $record.TTL
                            "Section" = $record.Section
                            "Strings" = $record.Strings -join ';'
                        }
                    }
                } else {
                    Write-Host " - $($dnsRecordName) not found" -ForegroundColor DarkYellow
                    [PSCustomObject]@{
                        "Domain"  = $dnsRecordName
                        "Type"    = "TXT"
                        "TTL"     = ""
                        "Section" = ""
                        "Strings"   = "No DKIM record found"
                    }
                }
            }
        }
        catch {
            Write-Host (Format-OagM365Exception -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $dnsRecordName
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving DKIM record; $($_.Exception.message)"
            }
        }
    } 
}
