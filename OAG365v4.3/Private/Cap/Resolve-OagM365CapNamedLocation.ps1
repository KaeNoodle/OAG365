function Resolve-OagM365CapNamedLocation {
    <#
    .SYNOPSIS
    Resolve a Microsoft Conditional Access Policy named location GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Conditional Access Policy named location GUID to an object containing Get-MgIdentityConditionalAccessNamedLocation properties including display name and CIDR IP address details

    .PARAMETER InputObject
    The GUID(s) of the named location to resolve

    .EXAMPLE
    Resolve-OagM365CapNamedLocation -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-OagM365CapNamedLocation
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityConditionalAccessNamedLocation

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-OagM365Guid -InputObject $InputObject) {
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
