function Get-OagM365IamPimRoleEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM role eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM role eligibilitiy data

    .EXAMPLE
    Get-OagM365IamPimRoleEligibility

    .NOTES
        NAME: Get-OagM365IamPimRoleEligibility
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
            $exception = Format-OagM365Exception -message "Failed retrieving PIM role eligibility schedules" -exception $_
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
