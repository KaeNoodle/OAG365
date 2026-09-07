function Format-OagM365IamPimRoleMgmtPolicy {
    <#
    .SYNOPSIS
    Returns a formatted object with the Entra PIM management policies for output

    .DESCRIPTION
    Format Entra PIM management policies obtained in Set-OagM365IamPimRoleMgmtPolicy for output

    .PARAMETER mgmtPolicyAssignment
    The management policy assignment obtained from Get-MgPolicyRoleManagementPolicyAssignment

    .PARAMETER mgmtPolicyRules
    The management policy rules obtained from Get-OagM365IamPimRoleMgmtPolicyRule for this assignment

    .EXAMPLE
    Format-OagM365IamPimRoleMgmtPolicy -mgmtPolicyRules $mgmtPolicyRules -mgmtPolicyAssignments $mgmtPolicyAssignments

    .NOTES
        NAME: Format-OagM365IamPimRoleMgmtPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Format-OagM365IamPimRoleMgmtPolicy

        CHANGELOG:    
    #>
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
