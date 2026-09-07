function Get-OagM365IamPimRoleAssignment{
    <#
    .SYNOPSIS
    Return the requested Entra PIM role assignment data

    .DESCRIPTION
    Return the requested Entra PIM role assignment data

    .PARAMETER type
    The type of object to return. Either activated or schedule

    .EXAMPLE
    Get-OagM365IamPimRoleAssignment -type "activated"

    .NOTES
        NAME: ExportM365-Pim-RoleAssignments
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRoleManagementDirectoryRoleAssignmentSchedule
           Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance

        CHANGELOG:    
    #>

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
            $exception = Format-OagM365Exception -message "Failed retrieving PIM role assignment schedules" -exception $_
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
