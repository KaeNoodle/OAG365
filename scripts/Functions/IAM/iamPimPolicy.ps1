function iamPimPolicyFormat {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns a formatted object with the Entra PIM management policies for output

    LOGIC
    Runs 1 loop over the returned collection.

    PARAMETERS
    -mgmtPolicyAssignment (required)
    -mgmtPolicyRules (required)

    RUNNING CONTEXT
    Called by  : iamPimPolicySet
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true)][object]$mgmtPolicyAssignment,
        [Parameter(Mandatory = $true)][object]$mgmtPolicyRules
    )

    $policyRules = @()

    $mgmtPolicyRules | ForEach-Object {
        $rule = $_

        $policyRules += [PsCustomObject]@{
            "Id" = $mgmtPolicyAssignment.Id
            "PolicyId" = $mgmtPolicyAssignment.PolicyId
            "PolicyDescription" = $mgmtPolicyAssignment.Policy.Description
            "PolicyDisplayName" = $mgmtPolicyAssignment.Policy.DisplayName
            "PolicyIsOrganizationDefault" = $mgmtPolicyAssignment.Policy.isOrganizationDefault
            "PolicyLastModifiedByName" = $mgmtPolicyAssignment.Policy.LastModifiedBy.DisplayName
            "PolicyLastModifiedById" = $mgmtPolicyAssignment.Policy.LastModifiedBy.Id
            "PolicyLastModifiedDateTime" = $mgmtPolicyAssignment.Policy.LastModifiedDateTime
            "PolicyAdditionalProperties" = $mgmtPolicyAssignment.AdditionalProperties | ConvertTo-Json -Depth 50   # Complex
            "RoleDefinitionId" = $mgmtPolicyAssignment.RoleDefinitionId
            "ScopeId" = $mgmtPolicyAssignment.ScopeId
            "ScopeType" = $mgmtPolicyAssignment.ScopeType
            "RuleId" = $rule.Id
            "RuleDataType" = $rule.AdditionalProperties.'@odata.type'
            "RuleTargetCaller" = $rule.Target.Caller
            "RuleTargetLevel" = $rule.Target.Level
            "RuleTargetOperations" = $rule.Target.Operations | ConvertTo-Json -Depth 50     #complex
            "RuleTargetObjects" = $rule.Target.TargetObjects
            "RuleTargetEnforcedSettings" = $rule.Target.EnforcedSettings | ConvertTo-Json -Depth 50   #Complex
            "RuleTargetInheritableSettings" = $rule.Target.InheritableSettings | ConvertTo-Json -Depth 50    #Complex
            "RuleApEnabledRules" = $rule.AdditionalProperties.enabledRules | ConvertTo-Json -Depth 50   # Complex 
            "RuleApIsEnabled" = $rule.AdditionalProperties.isEnabled
            "RuleApIsExpirationRequired" = $rule.AdditionalProperties.isExpirationRequired
            "RuleApMaxDuration" = $rule.AdditionalProperties.maximumDuration   
            "RuleApNotificationType" = $rule.AdditionalProperties.notificationType
            "RuleApReceipientType" = $rule.AdditionalProperties.receipientType
            "RuleApNotificationLevel" = $rule.AdditionalProperties.notificationLevel
            "RuleApIsDefaultRecipientsEnabled" = $rule.AdditionalProperties.isDefaultRecipientsEnabled
            "RuleApNotificationRecipients" = $rule.AdditionalProperties.notificationRecipients | ConvertTo-Json -Depth 50  # Complex 
            "RuleApSettingIsApprovalRequired" = $rule.AdditionalProperties.setting.isApprovalRequired
            "RuleApSettingIsApprovalRequiredForExtension" = $rule.AdditionalProperties.setting.isApprovalRequiredForExtension
            "RuleApSettingIsRequestorJustificationRequired" = $rule.AdditionalProperties.setting.isRequestorJustificationRequired
            "RuleApSettingApprovalMode" = $rule.AdditionalProperties.setting.approvalMode
            "RuleApSettingApprovalStages" = $rule.AdditionalProperties.setting.approvalStages | ConvertTo-Json -Depth 50   #complex
        }
    }

    return $policyRules
}


function iamPimPolicyRuleGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Return the requested Entra PIM management policy assignments

    LOGIC
    Runs 1 loop over the returned collection.
    Retrieves data using Get-MgPolicyRoleManagementPolicyRule.

    PARAMETERS
    -rolePolicyAssignment (required)

    RUNNING CONTEXT
    Called by  : iamPimPolicySet
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    Get-MgPolicyRoleManagementPolicyRule

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true)][object]$rolePolicyAssignment
    )

    $rules = @()
    $rolePolicyAssignment | ForEach-Object {
        $rule = Get-MgPolicyRoleManagementPolicyRule -UnifiedRoleManagementPolicyId $_.PolicyId -All
        $rules += $rule
    }

    return $rules
}


function iamPimPolicySet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Set the Entra PIM management policies for either a role or group

    LOGIC
    Runs 2 loops over the returned collection.
    Retrieves data using Get-MgPolicyRoleManagementPolicyAssignment.
    Wrapped in 2 try/catch blocks; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -roleId (required)
    -roleName (required)
    -groupId (required)
    -groupName (required)

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : exceptionFormat, iamPimPolicyFormat, iamPimPolicyRuleGet

    CMLETS/PERMISSIONS/SCOPES
    Get-MgPolicyRoleManagementPolicyAssignment

    --------------------------------------------------------------------------------#>

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
                $exception = exceptionFormat -message "Failed retrieving PIM role management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM role management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = iamPimPolicyRuleGet -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) DirectoryRole policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += iamPimPolicyFormat -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
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
                $exception = exceptionFormat -message "Failed retrieving PIM group management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM group management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = iamPimPolicyRuleGet -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) Group policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += iamPimPolicyFormat -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
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
