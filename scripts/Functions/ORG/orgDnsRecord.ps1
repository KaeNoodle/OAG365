function orgDomainDkimGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns DKIM DNS records for domains obtained from Get-MgDomains

    LOGIC
    Runs 3 loops over the returned collection.
    Queries public DNS with Resolve-DnsName. Needs no tenant permission.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -getMgDomains (required)

    RUNNING CONTEXT
    Called by  : orgEmailExport
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Resolve-DnsName

    --------------------------------------------------------------------------------#>

    
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
            Write-Host (exceptionFormat -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
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


function orgDomainDmarcGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns DMARC DNS records for domains obtained from Get-MgDomains

    LOGIC
    Runs 2 loops over the returned collection.
    Queries public DNS with Resolve-DnsName. Needs no tenant permission.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -getMgDomains (required)

    RUNNING CONTEXT
    Called by  : orgEmailExport
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Resolve-DnsName

    --------------------------------------------------------------------------------#>


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
            Write-Host (exceptionFormat -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
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


function orgDomainMxGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns MX DNS records for domains obtained from Get-MgDomains

    LOGIC
    Runs 2 loops over the returned collection.
    Retrieves data using Get-MgDomains.
    Queries public DNS with Resolve-DnsName. Needs no tenant permission.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -getMgDomains (required)

    RUNNING CONTEXT
    Called by  : orgEmailExport
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDomains
    Resolve-DnsName

    --------------------------------------------------------------------------------#>

    <#
    .SYNOPSIS
    Returns MX DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns MX DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    orgDomainMxGet -getMgDomains $object

    .NOTES
        NAME: orgDomainMxGet
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
            $records = @(Resolve-DnsName -Type MX -Name $_.Id -ErrorAction SilentlyContinue | Where-Object { $_.NameExchange -ne ""; })
            if ($records) {
                # One line per domain, not one per record. A domain routing to Exchange
                # Online resolves to several MX hosts, and printing each one made a
                # single-domain tenant look like nine domains in the transcript.
                Write-Host " - $($_.Id) MX found ($($records.Count) record(s))" -ForegroundColor Green

                foreach ($record in $records) {
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
            Write-Host (exceptionFormat -message " - $($_.Id) MX error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
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


function orgDomainSpfGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns SPF DNS records for domains obtained from Get-MgDomains

    LOGIC
    Runs 2 loops over the returned collection.
    Queries public DNS with Resolve-DnsName. Needs no tenant permission.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -getMgDomains (required)

    RUNNING CONTEXT
    Called by  : orgEmailExport
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Resolve-DnsName

    --------------------------------------------------------------------------------#>

    
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
            Write-Host (exceptionFormat -message " - $($_.Id) SPF error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
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
