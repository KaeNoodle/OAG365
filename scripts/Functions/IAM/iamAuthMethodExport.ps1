function iamAuthMethodExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns Microsoft Entra MFA settings for all users

    LOGIC
    Retrieves data using Get-MgReportAuthenticationMethodUserRegistrationDetail.
    Writes the result to CSV at the path held in $script:exportTarget.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.
    Returns $true on success and $false on failure.

    PARAMETERS
    None. Reads its inputs from module scope.

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Get-MgReportAuthenticationMethodUserRegistrationDetail

    --------------------------------------------------------------------------------#>


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
        $exception = exceptionFormat -message "Failed retrieving Entra user MFA report" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
