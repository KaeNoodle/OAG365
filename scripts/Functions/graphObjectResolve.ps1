function adminUnitResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves an administrative unit ID to its display name.

    LOGIC
    Sets up 1 lookup cache to avoid repeat Graph calls for the same ID.
    Runs 1 loop over the returned collection.
    Retrieves data using Get-MgDirectoryAdministrativeUnit.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : ?.ps1
    Calls      : guidVerify.ps1

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDirectoryAdministrativeUnit

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


function directoryObjectListResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a list of directory object IDs in one pass.

    LOGIC
    Runs 1 loop over the returned collection.

    PARAMETERS
    -srcObject (required)
    -getAttrib (required, one of: displayName,userPrincipalName)

    RUNNING CONTEXT
    Called by  : capPolicyToHtml.ps1
    Calls      : directoryObjectResolve.ps1, guidVerify.ps1

    CMLETS/PERMISSIONS/SCOPES
    None

    --------------------------------------------------------------------------------#>


    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$srcObject,
        [Parameter(Mandatory = $true)][ValidateSet('displayName','userPrincipalName')]$getAttrib
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        if (guidVerify -InputObject $PSItem) {
            $object = directoryObjectResolve -InputObject $PSItem
            switch ($getAttrib) {
                "userPrincipalName" {
                    $list.Add(($object.AdditionalProperties.userPrincipalName)) | Out-Null                    
                }
                "displayName" {
                    $list.Add(($object.AdditionalProperties.displayName)) | Out-Null                    
                }
            }
        } else {
            Write-Verbose " - $($PSItem)"
            $list.Add($PSItem) | Out-Null
        }
    }

    return $list
}


function directoryObjectResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a directory object ID (user, group, role or device) to its display name.

    LOGIC
    Sets up 1 lookup cache to avoid repeat Graph calls for the same ID.
    Retrieves data using Get-MgDirectoryObject, Get-MgServicePrincipal.
    Wrapped in 2 try/catch blocks; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : capPolicyToHtml, directoryObjectListResolve
    Calls      : exceptionFormat, guidVerify

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDirectoryObject
    Get-MgServicePrincipal

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
                if ($script:cacheMgObject.ContainsKey($InputObject)) {
                    Write-Debug "Cached display name for `"$InputObject`""
                    Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    return $script:cacheMgObject[$InputObject]
                } else {
                    try {
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $InputObject -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]
                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    } catch {
                        $exception = exceptionFormat -message "Unable to find object using Get_MgDirectoryObject using provided GUID" -exception $_
                        Write-Host $exception -ForegroundColor DarkYellow

                        $directoryObjectId = (Get-MgServicePrincipal -filter "appId eq '$($InputObject)'").id
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $directoryObjectId -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]

                        Write-Host " - Don't panic. I found the service principal '$($displayName)' using the provided string as an AppID instead."

                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup from ServicePrincipal AppId: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    }
                    return $directoryObject
                }
            }
            catch {
                Write-Warning "Unable to resolve directory object with ID $InputObject, might have been deleted!"
            }
        }
        return $InputObject
    }
}


$script:cacheNamedLocations = @{}


function directoryRoleListResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a list of directory role template IDs in one pass.

    LOGIC
    Runs 1 loop over the returned collection.

    PARAMETERS
    -srcObject (required)

    RUNNING CONTEXT
    Called by  : capPolicyToHtml
    Calls      : directoryRoleTemplateResolve

    CMLETS/PERMISSIONS/SCOPES
    None
    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((directoryRoleTemplateResolve -InputObject $PSItem)) | Out-Null
    }

    return $list
}


function directoryRoleTemplateResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a directory role template ID to the role display name.

    LOGIC
    Retrieves data using Get-MgDirectoryRoleTemplate.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : directoryRoleListResolve
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    Get-MgDirectoryRoleTemplate

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        try {
            # use hashtable as cache to limit API calls
            if (($script:cacheDirectoryRoleTemplates.ContainsKey($InputObject) -eq $false)) {
                $script:cacheDirectoryRoleTemplates = Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Group-Object -Property Id -AsHashTable
                Write-Host "Found $($script:cacheDirectoryRoleTemplates.Count) directory role templates, adding to cache"
            }

            Write-Verbose "Loading cached directory role template '$InputObject' $(($script:cacheDirectoryRoleTemplates[$InputObject]).DisplayName)"
            return $script:cacheDirectoryRoleTemplates[$InputObject].DisplayName
        } catch {
            Write-Warning "Unable to resolve directory role template with ID '$InputObject'"
        }

        return $InputObject
    }
}


function servicePrincipalListResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a list of service principal IDs in one pass.

    LOGIC
    Runs 1 loop over the returned collection.

    PARAMETERS
    -srcObject (required)

    RUNNING CONTEXT
    Called by  : capPolicyToHtml
    Calls      : servicePrincipalResolve

    CMLETS/PERMISSIONS/SCOPES
    None
    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((servicePrincipalResolve -InputObject $PSItem)) | Out-Null
    }

    return $list
}


function servicePrincipalResolve {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Resolves a service principal object ID to its application display name.

    LOGIC
    Sets up 1 lookup cache to avoid repeat Graph calls for the same ID.
    Retrieves data using Get-MgServicePrincipal.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : servicePrincipalListResolve
    Calls      : guidVerify

    CMLETS/PERMISSIONS/SCOPES
    Get-MgServicePrincipal

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
