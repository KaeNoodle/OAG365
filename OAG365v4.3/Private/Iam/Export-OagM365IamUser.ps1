function Export-OagM365IamUser {
    <#
    .SYNOPSIS
    Exports Microsoft Entra users

    .DESCRIPTION
    Exports Microsoft Entra users

    .EXAMPLE
    Export-OagM365IamUser

    .NOTES
        NAME: Export-OagM365IamUser
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgUser

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID users"

        # Get Entra users
        $users = Get-MgUser -All -Filter 'accountEnabled eq true' -ConsistencyLevel eventual -Property @("DisplayName", "UserPrincipalName", "SignInActivity", "UserType", "Mail", "AssignedLicenses", "AccountEnabled", "EmployeeId", "UsageLocation", "OnPremisesLastSyncDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "LastPasswordChangeDateTime", "PasswordPolicies", "CreatedDateTime", "CreationType", "DeletedDateTime", "id")
        $userCount = $users.count
        Write-Host " - Found $userCount M365 users"

        # Export to CSV
        $users | Select-Object $script:userExportProperties | Export-CSV $script:exportTarget.users -NoTypeInformation
        Write-Host " - Exported M365 users to CSV $($script:exportTarget.users)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving M365 users" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
