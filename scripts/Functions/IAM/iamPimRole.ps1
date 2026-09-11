function iamPimRoleGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Obtain permanent, eligible and active built-in role assignments. Return a custom
    object containing the assignment details.

    LOGIC
    Runs 6 loops over the returned collection.
    Retrieves data using Get-MgGroupTransitiveMember,
      Get-MgRoleManagementDirectoryRoleAssignment.

    PARAMETERS
    None. Reads its inputs from module scope.

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : iamPimGroupAssignmentGet, iamPimGroupEligibilityGet, iamPimPolicySet, iamPimRoleAssignmentGet, iamPimRoleEligibilityGet, iamPimRoleMemberGet

    CMLETS/PERMISSIONS/SCOPES
    Get-MgGroupTransitiveMember
    Get-MgRoleManagementDirectoryRoleAssignment

    --------------------------------------------------------------------------------#>

	[Cmdletbinding()]param ()

    Write-Host "`nEntra membership of built-in and privileged roles"

	# Data retrieval: Get all directory roles
	$pimRoleMembers = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty principal -All
	$rolesDefinition = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty roleDefinition -All 
	$assignmentSchedules = iamPimRoleAssignmentGet -type 'schedule'

	# Data retrieval: For each directory role, add the role definition and assignment schedule properties
	foreach ($role in $pimRoleMembers) {
		$newRoleDefinition = ($rolesDefinition | Where-Object { $_.id -eq $role.id }).roleDefinition
		$role.RoleDefinition = $newRoleDefinition

        if (-not $script:pimMgmtPolicyRolesCache[$role.roleDefinitionId]) {
            $roleMgmtPolicies = iamPimPolicySet -roleId $role.roleDefinitionId -roleName $newRoleDefinition.DisplayName
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
	$eligibleRoles = iamPimRoleEligibilityGet
	if ($null -ne $eligibleRoles) {
		$pimRoleMembers += $eligibleRoles
	}
    Write-Host " - Found $(($pimRoleMembers | Measure-Object).count) direct members of roles"
    

	# Data retrieval: Identify duplicates in roles array based on active and eligible
	$pimRoleActivations = iamPimRoleAssignmentGet -type 'activated'

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
            $groupMgmtPolicies = iamPimPolicySet -groupId $_.principalId -groupName $_.Principal.AdditionalProperties.displayName
        }

		$activeGroupMembers = Get-MgGroupTransitiveMember -GroupId $_.principalId -All -Property id, displayName, userPrincipalName
		
		$assignmentSchedules = iamPimGroupAssignmentGet -groupId $_.principalId		
		$eligibleRoleMembers = iamPimGroupEligibilityGet -groupId $_.principalId

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
	
	$privilegedUsers = iamPimRoleMemberGet -pimRoleMembers $pimRoleMembers -pimRoleActivations $pimRoleActivations
	return $privilegedUsers
}


function iamPimRoleMemberGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Build an array of privileged role members from data provided by
    iamPimRoleGet

    LOGIC
    Runs 5 loops over the returned collection.

    PARAMETERS
    -pimRoleMembers (required)
    -pimRoleActivations (required)

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>

	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [object]$pimRoleMembers,
        [Parameter(Mandatory = $true)][AllowNull()]
        [object]$pimRoleActivations
	)
	
	Write-Host "`nBuilding privileged role members array"
	$privilegedUsers = @()

	foreach ($role in $pimRoleMembers) {	
		if (!$role.status) {
			# Role is permanent. No start/end dates to obtain
			$role | Add-Member -MemberType NoteProperty -Name "scriptedStartDateTime" -Value "Permanent"
			$role | Add-Member -MemberType NoteProperty -Name "scriptedEndDateTime" -Value "Permanent"

			if ($role.assignmentSchedule.AssignmentType) {
				$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value $role.assignmentSchedule.AssignmentType			
			} else {
				$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value "Permanent"			
			}

			if ($role.assignmentSchedule.Status) {
				$role | Add-Member -MemberType NoteProperty -Name "Status" -Value $role.assignmentSchedule.Status
			} else {
				$role | Add-Member -MemberType NoteProperty -Name "Status" -Value "Permanent (non PIM)"
			}
		} else {
			# Role requires activation. Use activation data to determine start/end dates
			if ($role.Principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
				$roleActivation = $pimRoleActivations | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($role.activeGroupMembers.ContainsKey($_.principalId)) -and ($_.MemberType -eq "Group") }
			} else {
				$roleActivation = $pimRoleActivations | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.PrincipalId -eq $role.PrincipalId) }
			}
			
            if ($roleActivation.startDateTime) { 
                $tmpStartDateTime = $roleActivation.startDateTime | Select-Object -Unique | Sort-Object | Select-Object -First 1 
                $tmpAssignmentType = "Eligible (Active)" 
            } else { 
                $tmpStartDateTime = $null 
                $tmpAssignmentType = "Eligible" 
            }

            if ($roleActivation.endDateTime) { 
                $tmpEndDateTime = $roleActivation.endDateTime | Select-Object -Unique | Sort-Object -Descending | Select-Object -First 1 
            } else { 
                $tmpEndDateTime = $null 
            }

			$role | Add-Member -MemberType NoteProperty -Name "scriptedStartDateTime" -Value $tmpStartDateTime
			$role | Add-Member -MemberType NoteProperty -Name "scriptedEndDateTime" -Value $tmpEndDateTime
			$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value $tmpAssignmentType
		}
		
		switch ($role.principal.AdditionalProperties.'@odata.type') {
			'#microsoft.graph.servicePrincipal' {
				$thisPrincipalId = $role.principal.AdditionalProperties.appId
			}
			
			default {
				$thisPrincipalId = $role.PrincipalId
			}
		}
		
		Write-Host " - $($role.RoleDefinition.displayName): $($role.principal.AdditionalProperties.'@odata.type') $($role.principal.AdditionalProperties.displayName)"

		# Add PIM role assignment data to privileged users array
		$privilegedUsers += [PSCustomObject]@{		
			"PrincipalType" = $role.principal.AdditionalProperties.'@odata.type'
			"PrincipalId" = $thisPrincipalId
			"PrincipalDisplayName" = $role.principal.AdditionalProperties.displayName
			"PrincipalUpn" = $role.principal.AdditionalProperties.userPrincipalName
			"RoleDisplayName" = $role.RoleDefinition.displayName
			"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
			"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
			"OagRightsFrom" = ""
            "OagRightsSummary" = $role.OagRightsSummary
			"AssignmentType" = $role.scriptedAssignmentType
			"AssignmentStart" = $role.scriptedStartDateTime
			"AssignmentEnd" = $role.scriptedEndDateTime
			"EligibleStartDateTime" = ""
			"EligibleEndDateTime" = ""
			"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
			"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
			"RoleTemplateId" = $role.RoleDefinition.templateId
			"AssignmentStatus" = $role.Status
			"AssignmentAppScope" = $role.AppScope.DisplayName
			"AssignmentAppScopeId" = $role.AppScope.Id
			"AssignmentDirectoryScope" = $role.directoryScopeId
			"ScheduleExpireType" = $role.assignmentSchedule.ScheduleInfo.Expiration.Type
			"ScheduleExpireDuration" = $role.assignmentSchedule.ScheduleInfo.Expiration.Duration
			"ScheduleRecurrence" = ($role.assignmentSchedule.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
		}

		# If active PIM group members have been defined for this role, add to privileged users array.
		$role.activeGroupMembers | ForEach-Object {
			foreach ($activeMember in $_) {
				Write-Host "     + Active PIM role member $($activeMember.PrincipalDisplayName)"
		
				if ($role.AppScope.DisplayName) { $appScope = "$($role.AppScope.DisplayName) *" } else { $appScope = "" }			
				if ($role.AppScope.Id) { $appScopeId = "$($role.AppScope.Id) *" } else { $appScopeId = "" }
				if ($role.directoryScopeId) { $dirScope = "$($role.directoryScopeId) *" } else { $dirScope = "" }			
			
				if ($activeMember.Status) { $status = $activeMember.Status } else { $status = $role.Status }
					
				$privilegedUsers += [PSCustomObject]@{
					"PrincipalType" = $activeMember.PrincipalType
					"PrincipalId" = $activeMember.PrincipalId
					"PrincipalDisplayName" = $activeMember.PrincipalDisplayName
					"PrincipalUpn" = $activeMember.PrincipalUpn
					"RoleDisplayName" = $role.RoleDefinition.displayName
					"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
					"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
					"OagRightsFrom" = $activeMember.GroupName
                    "OagRightsSummary" = $activeMember.OagRightsSummary
					"AssignmentType" = $activeMember.AssignmentType
					"AssignmentStart" = $activeMember.AssignmentStartDateTime
					"AssignmentEnd" = $activeMember.AssignmentEndDateTime
					"EligibleStartDateTime" = ""
					"EligibleEndDateTime" = ""
					"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
					"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
					"RoleTemplateId" = $role.RoleDefinition.templateId
					"AssignmentStatus" = $status
					"AssignmentAppScope" = $appScope
					"AssignmentAppScopeId" = $appScopeId
					"AssignmentDirectoryScope" = $dirScope
					"ScheduleExpireType" = $activeMember.ScheduleExpireType
					"ScheduleExpireDuration" = $activeMember.ScheduleExpireDuration
					"ScheduleRecurrence" = $activeMember.ScheduleRecurrence
				}
			}
		}
			
		# If eligible PIM group members have been defined for this role, add to privileged users array.
		$role.eligibleRoleMembers | ForEach-Object {
			foreach ($eligibleMember in $_) {
				Write-Host "     + Eligible PIM Role Member $($eligibleMember.PrincipalDisplayName)"
				
				if ($role.AppScope.DisplayName) { $appScope = "$($role.AppScope.DisplayName) *" } else { $appScope = "" }			
				if ($role.AppScope.Id) { $appScopeId = "$($role.AppScope.Id) *" } else { $appScopeId = "" }
				if ($role.directoryScopeId) { $dirScope = "$($role.directoryScopeId) *" } else { $dirScope = "" }			
				
				$privilegedUsers += [PSCustomObject]@{
					"PrincipalType" = $eligibleMember.PrincipalType
					"PrincipalId" = $eligibleMember.PrincipalId
					"PrincipalDisplayName" = $eligibleMember.PrincipalDisplayName
					"PrincipalUpn" = $eligibleMember.PrincipalUpn
					"RoleDisplayName" = $role.RoleDefinition.displayName
					"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
					"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
					"OagRightsFrom" = $eligibleMember.GroupName
                    "OagRightsSummary" = $eligibleMember.OagRightsSummary
					"AssignmentType" = $eligibleMember.AssignmentType
					"AssignmentStart" = ""
					"AssignmentEnd" = ""
					"EligibleStartDateTime" = $eligibleMember.EligibleStartDateTime
					"EligibleEndDateTime" = $eligibleMember.EligibleEndDateTime
					"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
					"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
					"RoleTemplateId" = $role.RoleDefinition.templateId
					"AssignmentStatus" = $eligibleMember.Status
					"AssignmentAppScope" = $appScope
					"AssignmentAppScopeId" = $appScopeId
					"AssignmentDirectoryScope" = $dirScope
					"ScheduleExpireType" = $eligibleMember.ScheduleExpireType
					"ScheduleExpireDuration" = $eligibleMember.ScheduleExpireDuration
					"ScheduleRecurrence" = $eligibleMember.ScheduleRecurrence
				}
			}
		}
	}
	
	return $privilegedUsers
}
