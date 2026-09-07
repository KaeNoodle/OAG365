function Export-OagM365IamServicePrincipal {
    <#
    .SYNOPSIS
    Exports Microsoft Entra service principals and their assigned rights

    .DESCRIPTION
    Exports Microsoft Entra service principals and their assigned rights from any application role assignments or oAuth grants.

    .EXAMPLE
    Export-OagM365IamServicePrincipal

    .NOTES
        NAME: Export-OagM365IamServicePrincipal
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgServicePrincipal
           Get-MgServicePrincipalAppRoleAssignment
           Get-MgServicePrincipalOauth2PermissionGrant
           Get-MgDirectoryObject
           Find-MgGraphPermission

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID Service Principals"

        $pCache = @{}
        $scopeCache = @{}
        $spCache = @{}
        $returnServicePrincpals = @()
        $servicePrincipals = Get-MgServicePrincipal -All
        $spCount = $servicePrincipals.Count
        $spProcessed = 0
        Write-Host " - Found $($spCount) service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))"

        $servicePrincipals | ForEach-Object {
            $sp = $_
        
            Write-Verbose " - Processing assigned app roles and permission grants for $($sp.DisplayName) ($($sp.Id))"

            $spProcessed++
            $completed = [math]::Ceiling(($spProcessed / $spCount) * 100)
            Write-Progress -Activity "Processing $spProcessed of $spCount '$($sp.DisplayName)'" -PercentComplete $completed 

            # Collect roles assigned to this service principal
            $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -All -ServicePrincipalId $_.Id
            $appRolesCount = $appRoleAssignments.Count
            $appRolesProcessed = 0
            if ($appRolesCount -ge 1) {
                Write-Host "    + Found $($appRolesCount) application role assignments for $($($sp.DisplayName))"
            }

            # Add roles to assigned permissions list using custom object
            foreach ($appRoleAssignment in $appRoleAssignments) {
                $appRolesProcessed++
                $completedAppRoles = [math]::Ceiling(($appRolesProcessed / $appRolesCount) * 100)
                Write-Progress -Activity "Processing $($appRoleAssignment.PrincipalDisplayName) <-> $($appRoleAssignment.ResourceDisplayName): $appRolesProcessed of $appRolesCount" -PercentComplete $completedAppRoles 

                Write-Verbose "    + Documenting app role assignments for $appRoleAssignment "

                if (!$spCache.ContainsKey($appRoleAssignment.ResourceId)) { 
                    $spCache[$appRoleAssignment.ResourceId] = Get-MgServicePrincipal -ServicePrincipalId $appRoleAssignment.ResourceId
                    Write-Verbose "    + Retrieving service principal details"
                } else { 
                    Write-Verbose "    + Using cache for service principal details"
                }

                $role = $spCache[$appRoleAssignment.ResourceId].AppRoles | Where-Object Id -eq $appRoleAssignment.AppRoleId
                $item = [pscustomobject]@{
                        'ServicePrincipalId' = $sp.Id
                        'ServicePrincipalDisplayName' = $sp.DisplayName
                        'TypeOfRights' = 'App Role Assignment'
                        'ServicePrincipalType' = $sp.ServicePrincipalType
                        'ServicePrincipalAccountEnabled' = $sp.AccountEnabled
                        'ServicePrincipalDeletedDateTime' = $sp.DeletedDateTime

                        'AppId' = $sp.AppId
                        'AppDisplayName' = $sp.AppDisplayName

                        'ResourceId' = $appRoleAssignment.ResourceId
                        'ResourceDisplayName' = $appRoleAssignment.ResourceDisplayName

                        'ConsentType' = $appRoleAssignment.ConsentType

                        'PrincipalId' = $appRoleAssignment.PrincipalId
                        'PrincipalDisplayName' = $appRoleAssignment.PrincipalDisplayName
                        'PrincipalType' = $appRoleAssignment.PrincipalType

                        'Permission' = $role.Value
                        'PermissionDesc' = "$($role.DisplayName) - $($role.Description)"
                        'PermissionEnabled' = $role.IsEnabled

                        'PermissionType' = 'AppRole'
                        'PermissionType Identifier' = $appRoleAssignment.AppRoleId
                        'PermissionType DeletedDateTime' = $appRoleAssignment.DeletedDateTime
                        'PermissionType CreatedDateTime' = $appRoleAssignment.CreatedDateTime
                    }

                $returnServicePrincpals += $item
            }

            # Get oAuth2 permission grants for this service principal
            $grants = Get-MgServicePrincipalOauth2PermissionGrant -All -ServicePrincipalId $_.Id
            $grantsCount = $grants.Count
            $grantsProcessed = 0
            if ($grantsCount -ge 1) {
                Write-Host "    + Found $($grantsCount) permission grants for $($($sp.DisplayName))"
            }

            # Add grants to assigned permissions list using custom object
            foreach ($grant in $grants) {
                $grantsProcessed++
                $completedGrants = [math]::Ceiling(($grantsProcessed / $grantsCount) * 100)
                Write-Progress -Activity "Processing $($sp.AppDisplayName) <-> $($resource.DisplayName): $grantsProcessed of $grantsCount" -PercentComplete $completedGrants 

                if (!$spCache.ContainsKey($grant.ResourceId)) { 
                    $spCache[$grant.ResourceId] = Get-MgServicePrincipal -ServicePrincipalId $grant.ResourceId
                }

                if ($grant.PrincipalId) {
                    if (!$pCache.ContainsKey($grant.PrincipalId)) { 
                        $pCache[$grant.PrincipalId] = Get-MgDirectoryObject -DirectoryObjectId $grant.PrincipalId | 
                            Select-Object -Property Id, @{l='displayName'; e={$_.AdditionalProperties["displayName"]}}, @{l='userPrincipalName'; e={$_.AdditionalProperties["userPrincipalName"]}}, @{l='principalType'; e={$_.AdditionalProperties["@odata.type"]}}
                    }

                    $principal = $pCache[$grant.PrincipalId]
                } else {
                    $principal = @{}
                }

                $resource = $spCache[$grant.ResourceId]
                foreach ($scope in $grant.Scope.split(' ')) {
                    if ($scope -ne '') {                
                        $role = $resource.AppRoles | Where-Object Value -eq $scope
                        Write-Verbose "    + Documenting Oauth2 permission grants for $role in $scope "

                        if (!$scopeCache.ContainsKey($scope)) {
                            $scopeCache[$scope] = Find-MgGraphPermission -ExactMatch $scope -PermissionType Delegated -ErrorAction SilentlyContinue
                            Write-Verbose "    + Retrieving permission details for $scope"
                        } else { 
                            Write-Verbose "    + Using cache for permission details for $scope "
                        }
                        $permissions = $scopeCache[$scope]                

                        $item = [pscustomobject]@{
                                'ServicePrincipalId' = $sp.Id
                                'ServicePrincipalDisplayName' = $sp.DisplayName
                                'TypeOfRights' = 'Oauth2 Permission Grant'
                                'ServicePrincipalType' = $sp.ServicePrincipalType
                                'ServicePrincipalAccountEnabled' = $sp.AccountEnabled
                                'ServicePrincipalDeletedDateTime' = $sp.DeletedDateTime
        
                                'AppId' = $sp.AppId
                                'AppDisplayName' = $sp.AppDisplayName
            
                                'ResourceId' = $grant.ResourceId
                                'ResourceDisplayName' = $resource.DisplayName
        
                                'ConsentType' = $grant.ConsentType
        
                                'PrincipalId' = $grant.PrincipalId
                                'PrincipalDisplayName' = $principal.displayName
                                'PrincipalType' = $principal.principalType
        
                                'Permission' = $scope
                                'PermissionDesc' = $permissions.Description
                                'PermissionEnabled' = 'N/A'

                                'PermissionType' = $permissions.PermissionType    
                                'PermissionType Identifier' = $permissions.Id
                                'PermissionType DeletedDateTime' = $grant.DeletedDateTime
                                'PermissionType CreatedDateTime' = 'N/A'
                            }
                        $returnServicePrincpals += $item    
                    }
                }
            }
        }

        Write-Progress -Completed -Activity "Processing complete"

        # Export collected rights to CSV
        $returnServicePrincpals | Export-CSV $script:exportTarget.servicePrincipals -NoTypeInformation
        Write-Host " - Exported M365 service principals to CSV $($script:exportTarget.servicePrincipals)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving M365 service principals" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
