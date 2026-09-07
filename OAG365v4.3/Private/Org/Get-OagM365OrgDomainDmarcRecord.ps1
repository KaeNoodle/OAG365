function Get-OagM365OrgDomainDmarcRecord {
    <#
    .SYNOPSIS
    Returns DMARC DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns DMARC DNS records for domains obtained from Get-MgDomains. Uses selector1 and selector2 DMARC record types.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    Get-OagM365OrgDomainDmarcRecord -getMgDomains $object

    .NOTES
        NAME: Get-OagM365OrgDomainDmarcRecord
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>

    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $getMgDomains | ForEach-Object { 
        try {
            $dnsRecordName = "_dmarc." + $_.Id
            $records = Resolve-DnsName -Type TXT -Name $dnsRecordName -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=DMARC1"; } 
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
                    "Strings"  = "No DMARC record found"
                }
            }
        }
        catch {
            Write-Host (Format-OagM365Exception -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain" = $dnsRecordName
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Value"  = "Error retrieving DMARC record; $($_.Exception.message)"
            }
        }
    } 
}
