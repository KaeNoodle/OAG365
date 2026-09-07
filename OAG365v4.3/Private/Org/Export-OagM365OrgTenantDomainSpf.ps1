function Export-OagM365OrgTenantDomainSpf {
    <#
    .SYNOPSIS
    Returns SPF DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns SPF DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    Export-OagM365OrgTenantDomainSpf -getMgDomains $object

    .NOTES
        NAME: Export-OagM365OrgTenantDomainSpf
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
            $records = Resolve-DnsName -Type TXT -Name $_.Id -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=spf1"; }
            if ($records) {
                foreach ($record in $records) {
                    Write-Host " - $($_.Id) SPF found" -ForegroundColor Green
                    [PSCustomObject]@{
                        "Domain"  = $record.Name
                        "Type"    = $record.Type
                        "TTL"     = $record.TTL
                        "Section" = $record.Section
                        "Strings" = $record.Strings -join ';'
                    }
                }
            } else {
                Write-Host " - $($_.Id) SPF not found" -ForegroundColor DarkYellow
                [PSCustomObject]@{
                    "Domain"  = $_.Id
                    "Type"    = "TXT"
                    "TTL"     = ""
                    "Section" = ""
                    "Strings"   = "No SPF record found"
                }
            }
        }
        catch {
            Write-Host (Format-OagM365Exception -message " - $($_.Id) SPF error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $_.Id
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving SPF record; $($_.Exception.message)"
            }
        }
    } 
}
