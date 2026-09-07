function Export-OagM365CapAuthStrengthPolicy {
    <#
    .SYNOPSIS
    Exports Microsoft Entra authentication strength policies

    .DESCRIPTION
    Exports Microsoft Entra authentication strength policies

    .EXAMPLE
    Export-OagM365CapAuthStrengthPolicy

    .NOTES
        NAME: Export-OagM365CapAuthStrengthPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyAuthenticationStrengthPolicy

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID tenancy authentication strength policies"

        # Get defined authentication strength policies
        $getMgPolicyAuthStrengthPolicy = Get-MgPolicyAuthenticationStrengthPolicy -All
        $getMgPolicyAuthStrengthPolicyCount = $getMgPolicyAuthStrengthPolicy.count
        Write-Host " - Found $getMgPolicyAuthStrengthPolicyCount M365 authentication policies"

        $getMgPolicyAuthStrengthPolicy | Select-Object -Property Id, DisplayName, Description, PolicyType, RequirementsSatisfied, @{l='AllowedCombinations'; e={$_.AllowedCombinations | Out-String }}, @{l='CombinationConfigurations'; e={$_.CombinationConfigurations | Out-String }}, CreatedDateTime, ModifiedDateTime | Export-CSV $script:exportTarget.authStrengthPolicy -NoTypeInformation
        Write-Host " - Exported M365 authentication strength policies to CSV $($script:exportTarget.authStrengthPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving Entra authentication strength policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
