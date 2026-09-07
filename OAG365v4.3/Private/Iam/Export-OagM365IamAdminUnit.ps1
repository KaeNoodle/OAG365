function Export-OagM365IamAdminUnit {
    <#
    .SYNOPSIS
    Export any administrative units used to manage role allocations

    .DESCRIPTION
    Export any administrative units used to manage role allocations

    .EXAMPLE
    Export-OagM365IamAdminUnit

    .NOTES
        NAME: Export-OagM365IamAdminUnit
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           
        CHANGELOG:
    #>

    try {
        Write-Host "Entra Administrative Units"

        # Get all Entra role definitions to resolve IDs
        $adminUnits =  Get-MgDirectoryAdministrativeUnit -All
        Write-Host " - Found $(($adminUnits | Measure-Object).count) admin units"

        # Export to CSV
        $adminUnits | Select-Object -Property Id, DisplayName, Description, IsMemberManagementRestricted, MembershipRule, MembershipRuleProcessingState, MembershipType, Visibility, Extensions, @{l='AdditionalProperties'; e={$_.AdditionalProperties | ConvertTo-Json }} | Export-CSV $script:exportTarget.roleAdminUnits -NoTypeInformation
        Write-Host " - Exported M365 admin units to CSV $($script:exportTarget.roleAdminUnits)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving admin units" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
