function dfoSafeAttachmentsExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export safe attachment threat protection policies

    LOGIC
    Retrieves data using Get-SafeAttachmentPolicy, Get-SafeAttachmentRule.
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
    Get-SafeAttachmentPolicy
    Get-SafeAttachmentRule

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nDefender Threat Protection safe attachments policy settings"

        # Get DTP safe attachments settings
        $getSafeAttachmentPolicy = Get-SafeAttachmentPolicy
        $getSafeAttachmentPolicy | Export-CSV $script:exportTarget.dtppSafeAttachmentPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe attachments policy settings to $($script:exportTarget.dtppSafeAttachmentPolicy)" -ForegroundColor Green

        # Export as CSV
        $getSafeAttachmentRule = Get-SafeAttachmentRule
        $getSafeAttachmentRule | Export-CSV $script:exportTarget.dtppSafeAttachmentRule -NoTypeInformation
        Write-Host " - Exported DTP safe attachments rule to $($script:exportTarget.dtppSafeAttachmentRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving DTP safe attachments policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function dfoSafeLinksExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export safe links policy threat protection policies

    LOGIC
    Retrieves data using Get-SafeLinksPolicy, Get-SafeLinksRule.
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
    Get-SafeLinksPolicy
    Get-SafeLinksRule

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nDefender Threat Protection safe links policy settings"

        # Get DTP safe links settings
        $getSafeLinksPolicy = Get-SafeLinksPolicy 
        $getSafeLinksPolicy | Export-CSV $script:exportTarget.dtppSafeLinksPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe links policy settings to $($script:exportTarget.dtppSafeLinksPolicy)" -ForegroundColor Green

        $getSafeLinksRule = Get-SafeLinksRule 
        $getSafeLinksRule | Export-CSV $script:exportTarget.dtppSafeLinksRule -NoTypeInformation
        Write-Host " - Exported DTP safe links rule settings to $($script:exportTarget.dtppSafeLinksRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving DTP safe links policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
