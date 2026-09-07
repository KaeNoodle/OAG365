function Export-OagM365DfoQuarantinePolicy {
    <#
    .SYNOPSIS
    Export threat protection quarantine policies

    .DESCRIPTION
    Export quarantine policy settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-OagM365DfoQuarantinePolicy

    .NOTES
        NAME: Export-OagM365DfoQuarantinePolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-QuarantinePolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online quarantine policy settings"

        $getQuarantinePolicy = Get-QuarantinePolicy | Select-Object -Property Id, IsValid, Name, EndUserQuarantinePermissions, ESNEnabled, QuarantinePolicyType, QuarantineRetentionDays, IncludeMessagesFromBlockedSenderAddress, DistinguishedName, ObjectCategory, ObjectClass, WhenChanged, WhenCreated, WhenChangedUTC, WhenCreatedUTC, ExchangeObjectId, OrganizationalUnitRoot, OrganizationId, Guid, OriginatingServer, ObjectState, ExchangeVersion
        $getQuarantinePolicy | Export-CSV $script:exportTarget.dfoQuarantinePolicy -NoTypeInformation
        Write-Host " - Exported EOP quarantine policy settings to $($script:exportTarget.dfoQuarantinePolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving EOP quarantine policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
