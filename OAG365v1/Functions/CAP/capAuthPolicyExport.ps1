function capAuthMethodExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Microsoft Entra authentication method policies

    LOGIC
    Runs 4 loops over the returned collection.
    Retrieves data using Get-MgPolicyAuthenticationMethodPolicy.
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
    Get-MgPolicyAuthenticationMethodPolicy

    --------------------------------------------------------------------------------#>


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
        $exception = exceptionFormat -message "Failed retrieving Entra authentication method policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function capAuthStrengthExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Microsoft Entra authentication strength policies

    LOGIC
    Retrieves data using Get-MgPolicyAuthenticationStrengthPolicy.
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
    Get-MgPolicyAuthenticationStrengthPolicy

    --------------------------------------------------------------------------------#>


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
        $exception = exceptionFormat -message "Failed retrieving Entra authentication strength policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
