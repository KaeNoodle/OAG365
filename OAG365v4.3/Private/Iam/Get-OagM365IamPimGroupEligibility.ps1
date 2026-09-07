function Get-OagM365IamPimGroupEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM group eligibilitiy data

    .PARAMETER groupId
    The ID of the group to return eligibilitiy for

    .EXAMPLE
    Get-OagM365IamPimGroupEligibility -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Get-OagM365IamPimGroupEligibility
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
            $exception = Format-OagM365Exception -message "Failed retrieving PIM group eligibility schedules" -exception $_
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
