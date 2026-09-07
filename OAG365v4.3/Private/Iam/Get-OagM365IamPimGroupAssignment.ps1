function Get-OagM365IamPimGroupAssignment {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group assignment data

    .DESCRIPTION
    Return the requested Entra PIM group assignment data

    .PARAMETER groupId
    The ID of the group to return assignments for

    .EXAMPLE
    Get-OagM365IamPimGroupAssignment -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Get-OagM365IamPimGroupAssignment
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule

        CHANGELOG:    
    #>
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
            $exception = Format-OagM365Exception -message "Failed retrieving PIM group assignment schedules"  -exception $_
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
