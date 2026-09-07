function Get-OagM365IamPimRole {
    <#
    .SYNOPSIS
    Obtain permanent, eligible and active built-in role assignments. Return a custom object containing the assignment details. 

    .DESCRIPTION
    Obtain permanent, eligible and active built-in role assignments. Return a custom object containing the assignment details. 
       Permanently assigned to a built-in role
       Permanently assigned via a group within a built-in role
       Eligible through PIM
       Eligible and currently activated
       Eligible through a PIM group

    .EXAMPLE
    Get-OagM365IamPimRole

    .NOTES
        NAME: Get-OagM365IamPimRole
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            Get-MgRoleManagementDirectoryRoleAssignment
            ExportM365-Pim-RoleAssignments
            Get-OagM365IamPimRoleEligibility
            Get-MgGroupTransitiveMember
            Get-OagM365IamPimGroupAssignment
            Get-OagM365IamPimGroupEligibility
            Get-OagM365IamPimRoleMember
           
        CHANGELOG:
          2025-12-04: Rewrite of PIM related functions used in Stage 4 (built-in role memberships)
    #>
	[Cmdletbinding()]param ()

    Write-Host "`nEntra membership of built-in and privileged roles"

	# Data retrieval: Get all directory roles
	$pimRoleMembers = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty principal -All
	$rolesDefinition = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty roleDefinition -All 
	$assignmentSchedules = Get-OagM365IamPimRoleAssignment -type 'schedule'

	# Data retrieval: For each directory role, add the role definition and assignment schedule properties
	foreach ($role in $pimRoleMembers) {
		$newRoleDefinition = ($rolesDefinition | Where-Object { $_.id -eq $role.id }).roleDefinition
		$role.RoleDefinition = $newRoleDefinition

        if (-not $script:pimMgmtPolicyRolesCache[$role.roleDefinitionId]) {
            $roleMgmtPolicies = Set-OagM365IamPimRoleMgmtPolicy -roleId $role.roleDefinitionId -roleName $newRoleDefinition.DisplayName
        }
		
		if ($null -ne $assignmentSchedules) {
			$assignmentSchedule = $assignmentSchedules | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.principalId -eq $role.PrincipalId) }
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $assignmentSchedule
		} else {
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $null
		}

        Add-Member -InputObject $role -MemberType NoteProperty -Name "OagRightsSummary" -Value "Role-Direct-ActiveMember"        
	}

	# Data retrieval: Add eligible PIM roles to roles array
	$eligibleRoles = Get-OagM365IamPimRoleEligibility
	if ($null -ne $eligibleRoles) {
		$pimRoleMembers += $eligibleRoles
	}
    Write-Host " - Found $(($pimRoleMembers | Measure-Object).count) direct members of roles"
    

	# Data retrieval: Identify duplicates in roles array based on active and eligible
	$pimRoleActivations = Get-OagM365IamPimRoleAssignment -type 'activated'

	if ($null -ne $pimRoleActivations) {
        if ($null -ne $pimRoleMembers) {
            foreach ($activation in $pimRoleActivations) {
                $pimRoleMembers | Where-Object { ($null -ne $_) -and ($_.id -eq $activation.RoleAssignmentOriginId) -and ($_.Id -eq $activation.Id) } | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "duplicate" -Value $true }
            }
        }
	} else {
        Write-Host " - No PIM role activations found"
    }


    Write-Host " - Identifying inherited permission through the membership of groups assigned to roles"

	# Data retrieval: For each role assigned to a group, build arrays of active and eligible group members
	$pimRoleMembers | Where-Object { $_.Principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } | ForEach-Object {	
        $inheritedActiveRoleMembers = @()
		$inheritedEligibleRoleMembers = @()

        if (-not $script:pimMgmtPolicyGroupsCache[$_.principalId]) {
            $groupMgmtPolicies = Set-OagM365IamPimRoleMgmtPolicy -groupId $_.principalId -groupName $_.Principal.AdditionalProperties.displayName
        }

		$activeGroupMembers = Get-MgGroupTransitiveMember -GroupId $_.principalId -All -Property id, displayName, userPrincipalName
		
		$assignmentSchedules = Get-OagM365IamPimGroupAssignment -groupId $_.principalId		
		$eligibleRoleMembers = Get-OagM365IamPimGroupEligibility -groupId $_.principalId

        Write-Host " + Resolving members of $($_.Principal.AdditionalProperties.displayName) group ($($_.principalId))"

		# Add active role members from group, include active schedule and eligibilty info 
        if ($null -ne $activeGroupMembers) {
            foreach ($member in $activeGroupMembers) {
                Write-Host "      Adding active member $($member.AdditionalProperties.displayName)"
                
                $assignmentSchedule = $null
                $eligibilitySchedule = $null
                $assignmentSchedule = $assignmentSchedules | Where-Object { $_.PrincipalId -eq $member.Id }
                $eligibilitySchedule = $eligibleRoleMembers | Where-Object { $_.PrincipalId -eq $member.Id }
                
                $inheritedActiveRoleMembers += ([PSCustomObject][Ordered]@{
                    "PrincipalType" = $member.AdditionalProperties.'@odata.type'
                    "PrincipalId" = $member.Id
                    "PrincipalDisplayName" = $member.AdditionalProperties.displayName
                    "PrincipalUpn" = $member.AdditionalProperties.userPrincipalName
                    "AssignmentType" = "Active via Group"
                    "AssignmentStartDateTime" = $assignmentSchedule.ScheduleInfo.StartDateTime
                    "AssignmentEndDateTime" = $assignmentSchedule.ScheduleInfo.Expiration.EndDateTime
                    "GroupName" = $_.Principal.AdditionalProperties.displayName
                    "GroupId" = $_.PrincipalId
                    "Status" = $assignmentSchedule.Status
                    "OagRightsSummary" = "Role-Group-ActiveMember"
                    "EligibleStartDateTime" = $eligibilitySchedule.ScheduleInfo.StartDateTime
                    "EligibleEndDateTime" = $eligibilitySchedule.ScheduleInfo.Expiration.EndDateTime
                    "ScheduleExpireType" = $assignmentSchedule.ScheduleInfo.Expiration.Type
                    "ScheduleExpireDuration" = $assignmentSchedule.ScheduleInfo.Expiration.Duration
                    "ScheduleRecurrence" = ($assignmentSchedule.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
                })
            }
        }

		
        # Add eligible role members from PIM role eligibility schedules
        if ($null -ne $eligibleRoleMembers) {
            foreach ($member in $eligibleRoleMembers) {
                Write-Host "      Adding eligible role member $($member.principal.AdditionalProperties.displayName)"

                $inheritedEligibleRoleMembers += ([PSCustomObject][Ordered]@{
                    "PrincipalType" = $member.principal.AdditionalProperties.'@odata.type'
                    "PrincipalId" = $member.principal.Id
                    "PrincipalDisplayName" = $member.principal.AdditionalProperties.displayName
                    "PrincipalUpn" = $member.principal.AdditionalProperties.userPrincipalName
                    "AssignmentType" = "Eligible via Group"
                    "AssignmentStartDateTime" = ""
                    "AssignmentEndDateTime" = ""
                    "GroupName" = $_.Principal.AdditionalProperties.displayName
                    "GroupId" = $_.principalId
                    "Status" = $member.Status
                    "OagRightsSummary" = "Role-Group-EligibleMember"
                    "EligibleStartDateTime" = $member.ScheduleInfo.StartDateTime
                    "EligibleEndDateTime" = $member.ScheduleInfo.Expiration.EndDateTime
                    "ScheduleExpireType" = $member.ScheduleInfo.Expiration.Type
                    "ScheduleExpireDuration" = $member.ScheduleInfo.Expiration.Duration
                    "ScheduleRecurrence" = ($member.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
                })
            }
        }
            
		$_ | Add-Member -MemberType NoteProperty -Name "activeGroupMembers" -Value $inheritedActiveRoleMembers
		$_ | Add-Member -MemberType NoteProperty -Name "eligibleRoleMembers" -Value $inheritedEligibleRoleMembers
	}
	
	$privilegedUsers = Get-OagM365IamPimRoleMember -pimRoleMembers $pimRoleMembers -pimRoleActivations $pimRoleActivations
	return $privilegedUsers
}
