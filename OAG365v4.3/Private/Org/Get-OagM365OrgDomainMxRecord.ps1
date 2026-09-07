function Get-OagM365OrgDomainMxRecord {
    <#
    .SYNOPSIS
    Returns MX DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns MX DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    Get-OagM365OrgDomainMxRecord -getMgDomains $object

    .NOTES
        NAME: Get-OagM365OrgDomainMxRecord
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
            $records = Resolve-DnsName -Type MX -Name $_.Id -ErrorAction SilentlyContinue | Where-Object { $_.NameExchange -ne ""; }
            if ($records) {
                foreach ($record in $records) {
                    Write-Host " - $($_.Id) MX found" -ForegroundColor Green

                    if ([string]::IsNullOrEmpty($record.NameExchange) -eq $false) { 
                        $strings = $record.NameExchange -join '; '
                    } elseif ([string]::IsNullOrEmpty($record.IPAddress) -eq $false) {
                        $strings = $record.IPAddress -join '; '
                    } else {
                        $strings = ""
                    }

                    [PSCustomObject]@{
                        "Domain"  = $record.Name
                        "Type"    = $record.Type
                        "TTL"     = $record.TTL
                        "Section" = $record.Section
                        "Strings" = $strings
                    }
                }
            } else {
                Write-Host " - $($_.Id) MX not found" -ForegroundColor DarkYellow
                [PSCustomObject]@{
                    "Domain"  = $_.Id
                    "Type"    = "MX"
                    "TTL"     = ""
                    "Section" = ""
                    "Strings"   = "No MX record found"
                }
            }
        }
        catch {
            Write-Host (Format-OagM365Exception -message " - $($_.Id) MX error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $_.Id
                "Type"    = "MX"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving MX record; $($_.Exception.message)"
            }
        }
    } 
}
