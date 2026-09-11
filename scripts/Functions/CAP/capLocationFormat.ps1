function capLocationToHtml {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Outputs HTML table formatted rows containing location data from a conditional
    access policy

    LOGIC
    Runs 2 loops over the returned collection.

    PARAMETERS
    -locations (required)
    -type (required, one of: IncludeLocation, ExcludeLocation)

    RUNNING CONTEXT
    Called by  : capPolicyToHtml
    Calls      : capNamedLocationResolve

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>


    [CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][AllowNull()]$locations,
		[Parameter(Mandatory = $true)][ValidateSet('IncludeLocation', 'ExcludeLocation')]$type
    )

    if ($null -eq $locations) { 
        return ''; 
    }

    foreach ($location in $locations) {
        $thisLocation =  capNamedLocationResolve -InputObject $location

        $out_displayName = "(no name found)"
        if ($thisLocation.DisplayName) {
            $out_displayName = $thisLocation.DisplayName
        }

        $out_cidrAddress = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('cidrAddress')) {
            $out_cidrAddress = ($thisLocation.AdditionalProperties['cidrAddress'] | Sort-Object | Out-String)
        }

        $out_ipRanges = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('ipRanges')) {
            $out_ipRanges = ($thisLocation.AdditionalProperties['ipRanges'] | ForEach-Object{ $_.cidrAddress } | Sort-Object) -join $separator
        }

        $out_countriesAndRegions = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('countriesAndRegions')) {
            $out_countriesAndRegions = ($thisLocation.AdditionalProperties['countriesAndRegions'] | Sort-Object | Out-String )
        }

        $out_includeUnknownCountriesAndRegions = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('includeUnknownCountriesAndRegions')) {
            $out_cidrAddress = ($thisLocation.AdditionalProperties['includeUnknownCountriesAndRegions'] | Sort-Object | Out-String)
        }

        '<tr><td>{0}<br/><nobr>{1}<nobr></td><td>{2}</td><td><p>CIDR Addresses: {3} </p><p>IP Ranges: {4}</p><p>Countries and Regions: {5} </p><p>Include Unknown Countries and Regions: {6}</p></td></tr>' -f $out_displayName, $location, $type, $out_cidrAddress, $out_ipRanges, $out_countriesAndRegions, $out_includeUnknownCountriesAndRegions
    }
}


function capNamedLocationResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolve a Microsoft Conditional Access Policy named location GUID to an object

    LOGIC
    Sets up 1 lookup cache to avoid repeat Graph calls for the same ID.
    Runs 1 loop over the returned collection.
    Retrieves data using Get-MgIdentityConditionalAccessNamedLocation.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : capLocationToHtml
    Calls      : guidVerify

    CMLETS/PERMISSIONS/SCOPES
    Get-MgIdentityConditionalAccessNamedLocation

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (guidVerify -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheNamedLocations.ContainsKey($InputObject) -eq $false) {
                    $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
                    $namedLocations | ForEach-Object { $script:cacheNamedLocations[$_.Id] = $_ }
                    Write-Host "Found $($script:cacheNamedLocations.Count) named locations, adding to cache"
                }

                Write-Verbose "Loading cached named location '$InputObject' $(($script:cacheNamedLocations[$InputObject]).DisplayName)"
                return $script:cacheNamedLocations[$InputObject]
            } catch {
                Write-Warning "Unable to resolve named location with ID $InputObject!"
            }
        }
        return $InputObject
    }
}

$script:cacheAdminUnits = @{}
