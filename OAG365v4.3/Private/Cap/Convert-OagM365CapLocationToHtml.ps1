function Convert-OagM365CapLocationToHtml {
    <#
    .SYNOPSIS
    Outputs HTML table formatted rows containing location data from a conditional access policy

    .DESCRIPTION
    Outputs HTML table formatted rows containing location data from a conditional access policy

    .PARAMETER locations
    List of locations to iterate through and format.

    .PARAMETER type
    IncludeLocations or ExcludeLocations. Determins how row containing location data will be described in HTML report.

    .EXAMPLE
    Convert-OagM365CapLocationToHtml -locations $object -type 'IncludeLocation'

    .NOTES
        NAME: Convert-OagM365CapLocationToHtml
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-OagM365CapNamedLocation
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up named locations
    #>

    [CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][AllowNull()]$locations,
		[Parameter(Mandatory = $true)][ValidateSet('IncludeLocation', 'ExcludeLocation')]$type
    )

    if ($null -eq $locations) { 
        return ''; 
    }

    foreach ($location in $locations) {
        $thisLocation =  Resolve-OagM365CapNamedLocation -InputObject $location

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
