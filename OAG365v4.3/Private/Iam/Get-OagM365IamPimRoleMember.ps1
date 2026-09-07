function Get-OagM365IamPimRoleMember {
    <#
    .SYNOPSIS
    Build an array of privileged role members from data provided by Get-OagM365IamPimRole

    .DESCRIPTION
    Build an array of privileged role members from data provided by Get-OagM365IamPimRole.

    In addition to the information obtained through the Graph API, add the OAG summary fields:
         OagRightsFrom: Specifies the name of the group the user gets their role membership from 
                        if they are not a direct member of the role.
      OagRightsSummary: Summaries how the user got their rights and if they are active or eligible

    Users granted a role directly (not using PIM) will only be listed once in that role. However, users granted
    a role using PIM will be listed as 'Eligible' and, if activated, a second time as 'Assigned'.
       
    .EXAMPLE
    Get-OagM365IamPimRoleMember

    .NOTES
        NAME: Get-OagM365IamPimRoleMember
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            Get-OagM365IamPimRole (called by, requires data from)
           
        CHANGELOG:
    #>
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
