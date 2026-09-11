function iamPimGroupAssignmentGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Return the requested Entra PIM group assignment data

    LOGIC
    Retrieves data using
      Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -groupId (required)

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule

    --------------------------------------------------------------------------------#>

	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [string]$groupId
    )

	process {
		try {
			$assignments = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule -All -Filter "groupId eq '$($groupId)'" -ErrorAction Stop
			Write-Host " - Lookup: Found $(($assignments | Measure-Object).Count) PIM group assignment schedules"
            return $assignments
		} catch {
            $exception = exceptionFormat -message "Failed retrieving PIM group assignment schedules"  -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM group eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}


function iamPimGroupEligibilityGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Return the requested Entra PIM group eligibilitiy data

    LOGIC
    Retrieves data using
      Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -groupId (required)

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule

    --------------------------------------------------------------------------------#>

    <#
    .SYNOPSIS
    Return the requested Entra PIM group eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM group eligibilitiy data

    .PARAMETER groupId
    The ID of the group to return eligibilitiy for

    .EXAMPLE
    iamPimGroupEligibilityGet -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: iamPimGroupEligibilityGet
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule

        CHANGELOG:    
    #>       
	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [string]$groupId
    )
	
	process {
		try {
			$eligible = Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule -All -Filter "groupId eq '$($groupId)'" -ExpandProperty principal -ErrorAction Stop
			Write-Host " - Lookup: Found $(($eligible | Measure-Object).Count) PIM group eligibility schedules"
			return $eligible
		} catch {            
            $exception = exceptionFormat -message "Failed retrieving PIM group eligibility schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM group eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}


function iamPimRoleAssignmentGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Return the requested Entra PIM role assignment data

    LOGIC
    Retrieves data using Get-MgRoleManagementDirectoryRoleAssignmentSchedule,
      Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -type (required, one of: schedule, activated)

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgRoleManagementDirectoryRoleAssignmentSchedule
    Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance

    --------------------------------------------------------------------------------#>


	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('schedule', 'activated')]
        [string]$type
    )

	process {
		try {
			switch ($type) {
				'schedule' {
					$assignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -All -ErrorAction Stop
				}
				'activated' {
					$assignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -Filter "AssignmentType eq 'Activated'" -ErrorAction Stop
				}
			}
			
			Write-Host " - Lookup: Found $(($assignments | Measure-Object).Count) PIM role assignment schedules"
			return $assignments	
		} catch {
            $exception = exceptionFormat -message "Failed retrieving PIM role assignment schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM role assignment schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}


function iamPimRoleEligibilityGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Return the requested Entra PIM role eligibilitiy data

    LOGIC
    Retrieves data using Get-MgRoleManagementDirectoryRoleEligibilitySchedule.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    None. Reads its inputs from module scope.

    RUNNING CONTEXT
    Called by  : iamPimRoleGet
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgRoleManagementDirectoryRoleEligibilitySchedule

    --------------------------------------------------------------------------------#>

    <#
    .SYNOPSIS
    Return the requested Entra PIM role eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM role eligibilitiy data

    .EXAMPLE
    iamPimRoleEligibilityGet

    .NOTES
        NAME: iamPimRoleEligibilityGet
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRoleManagementDirectoryRoleEligibilitySchedule

        CHANGELOG:    
    #>    
	[Cmdletbinding()]param()
	
	process {
		try {
			$eligible = (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty principal,roleDefinition -All -ErrorAction Stop | Select-Object id, principalId, directoryScopeId, roleDefinitionId, status, @{Name = "OagRightsSummary"; Expression = { "Role-Direct-EligibleMember" }}, principal, roleDefinition)
			Write-Host " - Lookup: Found $(($eligible | Measure-Object).Count) PIM role eligibility schedules"
			return $eligible
		} catch {
            $exception = exceptionFormat -message "Failed retrieving PIM role eligibility schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM role eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}

		return $null
	}
}
