function Resolve-OagM365ServicePrincipal {
    <#
    .SYNOPSIS
    Resolve a service principal appId

    .DESCRIPTION
    Resolve a service principal appId to an object

    .PARAMETER InputObject
    The appId of the service principal to resolve

    .EXAMPLE
    Resolve-OagM365ServicePrincipal -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-OagM365ServicePrincipal
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgServicePrincipal

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
                if ($script:cacheServicePrincipals.ContainsKey($InputObject) -eq $false) {
                    $script:cacheServicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop | Group-Object -Property AppId -AsHashTable
                    Write-Host "Found $($script:cacheServicePrincipals.Count) service principals, adding to cache"
                }

                Write-Verbose "Loading cached service principal '$InputObject' $(($script:cacheServicePrincipals[$InputObject]).DisplayName)"
                return $script:cacheServicePrincipals[$InputObject]
            } catch {
                Write-Warning "Unable to resolve service principal with ID $InputObject!"
            }
        }
        return $InputObject
    }
}


$script:cacheDirectoryRoleTemplates = @{}
