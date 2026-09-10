function dfoAntiMalwareExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export anti-malware policy and rule threat protection policies

    LOGIC
    Retrieves data using Get-MalwareFilterPolicy, Get-MalwareFilterRule.
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
    Get-MalwareFilterPolicy
    Get-MalwareFilterRule

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nExchange Online Protection anti-malware protection settings"

        # Get EOP malware policy settings
        $getMalwareFilterPolicy = Get-MalwareFilterPolicy 
        $getMalwareFilterPolicy | Export-CSV $script:exportTarget.eopMalwarePolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection policy settings to $($script:exportTarget.eopMalwarePolicy)" -ForegroundColor Green

        $getMalwareFilterRule = Get-MalwareFilterRule 
        $getMalwareFilterRule | Export-CSV $script:exportTarget.eopMalwareRule -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection rules to $($script:exportTarget.eopMalwareRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving EOP anti-malware policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function dfoAntiPhishingExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export anti-malware policy and rule threat protection policies

    LOGIC
    Retrieves data using Get-AntiPhishPolicy, Get-AntiPhishRule.
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
    Get-AntiPhishPolicy
    Get-AntiPhishRule

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nExchange Online Protection anti-phishing protection settings"

        # Get EOP phishing policy settings
        $getAntiPhishPolicy = Get-AntiPhishPolicy 
        $getAntiPhishPolicy | Export-CSV $script:exportTarget.eopPhishPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection policy settings to $($script:exportTarget.eopPhishPolicy)" -ForegroundColor Green

        # Export to CSV
        $getAntiPhishRule = Get-AntiPhishRule 
        $getAntiPhishRule | Export-CSV $script:exportTarget.eopPhishRule -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection rule to $($script:exportTarget.eopPhishRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving EOP anti-phishing policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function dfoAntiSpamExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export anti-spam policy threat protection policies

    LOGIC
    Retrieves data using Get-HostedConnectionFilterPolicy.
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
    Get-HostedConnectionFilterPolicy

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nExchange Online Protection anti-spam protection settings"

        # Get EOP spam policy settings
        $getHostedConnectionFilterPolicy = Get-HostedConnectionFilterPolicy 
        $getHostedConnectionFilterPolicy | Export-CSV $script:exportTarget.eopSpamPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-spam protection policy settings to $($script:exportTarget.eopSpamPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving EOP anti-spam policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
