function Get-OagM365IamPimRoleMgmtPolicyRule {
    <#
    .SYNOPSIS
    Return the requested Entra PIM management policy assignments 

    .DESCRIPTION
    Return the requested Entra PIM management policy assignments 

    .PARAMETER rolePolicyAssignment
    The results of the policy look up in Get-OagM365IamPimRoleMgmtPolicyAssignment.

    .EXAMPLE
    Get-OagM365IamPimRoleMgmtPolicyRule -rolePolicyAssignment $rolePolicyAssignment

    .NOTES
        NAME: Get-OagM365IamPimRoleMgmtPolicyRule
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyRoleManagementPolicyRule

        CHANGELOG:    
    #>
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
