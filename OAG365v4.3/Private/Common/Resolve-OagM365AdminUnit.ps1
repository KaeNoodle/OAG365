function Resolve-OagM365AdminUnit {
    <#
    .SYNOPSIS
    Resolve an Entra Administrative Unit ID to an object

    .DESCRIPTION
    Resolve an Entra Administrative Unit ID to an object

    .PARAMETER InputObject
    The GUID(s) of the admin unit to resolve

    .EXAMPLE
    Resolve-OagM365AdminUnit -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-OagM365AdminUnit
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryAdministrativeUnit

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
                if ($script:cacheAdminUnits.ContainsKey($InputObject) -eq $false) {
                    $adminUnits = Get-MgDirectoryAdministrativeUnit -All -ErrorAction Stop
                    $adminUnits | ForEach-Object { $script:cacheAdminUnits[$_.Id] = $_ }
                    Write-Host "Found $($script:cacheAdminUnits.Count) admin units, adding to cache"
                }

                Write-Verbose "Loading cached admin unit '$InputObject' $(($script:cacheAdminUnits[$InputObject]).DisplayName)"
                return $script:cacheAdminUnits[$InputObject]
            } catch {
                Write-Warning "Unable to resolve admin unit with ID $InputObject!"
            }
        }
        return $InputObject
    }
}

$script:cacheServicePrincipals = @{}
