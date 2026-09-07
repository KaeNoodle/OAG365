function Export-OagM365IamUserAuthMethod {
    <#
    .SYNOPSIS
    Returns Microsoft Entra MFA settings for all users

    .DESCRIPTION
    Creates a Microsoft Entra user registration details report for each user including and SSPR, MFA settings and preferences. Returns details for all users.

    .EXAMPLE
    Export-OagM365IamUserAuthMethod

    .NOTES
        NAME: Export-OagM365IamUserAuthMethod
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgReportAuthenticationMethodUserRegistrationDetail

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID user MFA registration report"

        # Get user MFA regsistration report
        $UserMFA = Get-MgReportAuthenticationMethodUserRegistrationDetail -All:$true -Property * |
            Select-Object UserDisplayName, UserPrincipalName, UserType, `
                IsAdmin, IsMfaCapable, IsMfaRegistered, IsPasswordlessCapable, IsSsprCapable, IsSsprEnabled, IsSsprRegistered, IsSystemPreferredAuthenticationMethodEnabled, `
                @{n='MethodsRegistered';e={($_.MethodsRegistered -join ", ")}}, `
                @{n='SystemPreferredAuthenticationMethods';e={($_.SystemPreferredAuthenticationMethods -join ", ")}}, @{n='UserPreferredMethodForSecondaryAuthentication';e={($_.UserPreferredMethodForSecondaryAuthentication -join ", ")}}, `
                LastUpdatedDateTime, Id
        Write-Host "Found MFA settings for $(($UserMFA | Measure-Object).Count) users"

        # Export to CSV
        $UserMFA | Export-Csv -Path $script:exportTarget.userMfa -NoTypeInformation -Append
        Write-Host "Exported MFA settings to $($script:exportTarget.userMfa)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving Entra user MFA report" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
