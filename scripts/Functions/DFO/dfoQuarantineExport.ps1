function dfoQuarantineExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Export threat protection quarantine policies

    LOGIC
    Retrieves data using Get-QuarantinePolicy.
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
    Get-QuarantinePolicy

    --------------------------------------------------------------------------------#>


    try {
        Write-Host "`nExchange Online quarantine policy settings"

        $getQuarantinePolicy = Get-QuarantinePolicy | Select-Object -Property Id, IsValid, Name, EndUserQuarantinePermissions, ESNEnabled, QuarantinePolicyType, QuarantineRetentionDays, IncludeMessagesFromBlockedSenderAddress, DistinguishedName, ObjectCategory, ObjectClass, WhenChanged, WhenCreated, WhenChangedUTC, WhenCreatedUTC, ExchangeObjectId, OrganizationalUnitRoot, OrganizationId, Guid, OriginatingServer, ObjectState, ExchangeVersion
        $getQuarantinePolicy | Export-CSV $script:exportTarget.dfoQuarantinePolicy -NoTypeInformation
        Write-Host " - Exported EOP quarantine policy settings to $($script:exportTarget.dfoQuarantinePolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = exceptionFormat -message "Failed retrieving EOP quarantine policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
