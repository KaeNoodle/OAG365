function Set-OagM365IamPimRoleMgmtPolicy {
    <#
    .SYNOPSIS
    Set the Entra PIM management policies for either a role or group

    .DESCRIPTION
    Adds the provided Entra PIM management policies for either a role or group to the PIM role management policies cache object.

    .PARAMETER roleId
    The ID of the role to lookup management policies for

    .PARAMETER roleName
    The display name of the role - for ease of log reading.

    .PARAMETER groupId
    The ID of the group to lookup management policies for

    .PARAMETER groupName
    The display name of the group - for ease of log reading.

    .EXAMPLE
    Set-OagM365IamPimRoleMgmtPolicy -roleId $roleId -roleName $roleDisplayName

    .EXAMPLE
    Set-OagM365IamPimRoleMgmtPolicy -groupId $groupId -groupName $groupDisplayName

    .NOTES
        NAME: Set-OagM365IamPimRoleMgmtPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-OagM365IamPimRoleMgmtPolicyRule
           Format-OagM365IamPimRoleMgmtPolicy
           Get-MgPolicyRoleManagementPolicyAssignment
           Get-MgPolicyRoleManagementPolicyRule

        CHANGELOG:    
    #>
    [Cmdletbinding(DefaultParameterSetName = 'ByRole')]param (
        [Parameter(Mandatory = $true, ParameterSetName='ByRole')][object]$roleId,
        [Parameter(Mandatory = $true, ParameterSetName='ByRole')][object]$roleName,
        [Parameter(Mandatory = $true, ParameterSetName='ByGroup')][object]$groupId,
        [Parameter(Mandatory = $true, ParameterSetName='ByGroup')][object]$groupName
    )

    switch ($PSCmdlet.ParameterSetName) {
        'ByRole' {
            Write-Host " - PIM role management policy for $roleName ($roleId) "

            if ($script:pimMgmtPolicyRolesCache[$roleId]) {
                Write-Host "     - Returning cached role management policy results "
                return $script:pimMgmtPolicyRolesCache[$roleId]
            }

            try {
                $assignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '$($roleId)'"  -ExpandProperty "policy" -All -ErrorAction Stop
                Write-Host "     - Found $(($assignments | Measure-Object).Count) DirectoryRole policy assignments "
            } catch {
                $exception = Format-OagM365Exception -message "Failed retrieving PIM role management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM role management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = Get-OagM365IamPimRoleMgmtPolicyRule -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) DirectoryRole policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += Format-OagM365IamPimRoleMgmtPolicy -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
                }
            }

            if ($policyRules) {
                $script:pimMgmtPolicyRules += $policyRules
                $script:pimMgmtPolicyRolesCache[$roleId] = $policyRules
            } else {
                Write-Host "     - No role management policy assignments or rules found "
            }

            return $script:pimMgmtPolicyRolesCache[$roleId]
        }
        'ByGroup' {
            Write-Host " - PIM group management policy for $groupName ($groupId) "

            if ($script:pimMgmtPolicyGroupsCache[$groupId]) {
                Write-Host "     - Returning cached group management policy "
                return $script:pimMgmtPolicyGroupsCache[$groupId]
            } 

            try {
                $assignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '$($groupId)' and scopeType eq 'Group'" -ExpandProperty "policy" -All -ErrorAction Stop
                Write-Host "     - Found $(($assignments | Measure-Object).Count) group policy assignments "
            } catch {
                $exception = Format-OagM365Exception -message "Failed retrieving PIM group management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM group management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = Get-OagM365IamPimRoleMgmtPolicyRule -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) Group policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += Format-OagM365IamPimRoleMgmtPolicy -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
                }
            }

            if ($policyRules) {
                $script:pimMgmtPolicyRules += $policyRules
                $script:pimMgmtPolicyGroupsCache[$groupId] = $policyRules
            } else {
                Write-Host "     - No group management policy assignments or rules found "
            }

            return $script:pimMgmtPolicyGroupsCache[$groupId]
        }
    } 
}
