function Export-OagM365CapAuthMethodPolicy {
    <#
    .SYNOPSIS
    Exports Microsoft Entra authentication method policies

    .DESCRIPTION
    Exports Microsoft Entra authentication method policies

    .EXAMPLE
    ExportM365-AuthPolicy

    .NOTES
        NAME: ExportM365-AuthPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyAuthenticationMethodPolicy

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID tenancy authentication method policies"

        # Get defined authentication method policies
        $getMgPolicyAuthMethodPolicy = Get-MgPolicyAuthenticationMethodPolicy
        $getMgPolicyAuthMethodPolicyCount = $getMgPolicyAuthMethodPolicy.count
        Write-Host " - Found $getMgPolicyAuthMethodPolicyCount M365 authentication policies"
        
        # Export as CSV
        $getMgPolicyAuthMethodPolicy | ForEach-Object { 
            $authMethodPolicy = $_
            ForEach ($item in $_.AuthenticationMethodConfigurations) {
                $item.AdditionalProperties | ForEach-Object {
                    foreach ($key in $_.Keys) {
                        [PSCustomObject]@{
                            "Id"                     = $authMethodPolicy.Id
                            "DisplayName"            = $authMethodPolicy.DisplayName
                            "LastModifiedDateTime"   = $authMethodPolicy.LastModifiedDateTime
                            "PolicyMigrationState"   = $authMethodPolicy.PolicyMigrationState
                            "PolicyVersion"          = $authMethodPolicy.PolicyVersion
                            "AuthMethodId"           = $item.Id
                            "AuthMethodState"        = $item.State
                            "AUthMethodSetting"      = $key
                            "AuthMethodSettingValue" = ($_[$key] | ConvertTo-Json)
                        }
                    }
                }
            }
        } | Export-CSV $script:exportTarget.authMethodPolicy -NoTypeInformation
        Write-Host " - Exported M365 authentication methods policies to CSV $($script:exportTarget.authMethodPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving Entra authentication method policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
