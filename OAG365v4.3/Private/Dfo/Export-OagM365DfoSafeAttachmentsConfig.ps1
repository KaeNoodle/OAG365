function Export-OagM365DfoSafeAttachmentsConfig {
    <#
    .SYNOPSIS
    Export safe attachment threat protection policies

    .DESCRIPTION
    Export safe attachments policy settings from Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-OagM365DfoSafeAttachmentsConfig

    .NOTES
        NAME: Export-OagM365DfoSafeAttachmentsConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeAttachmentPolicy
           Get-SafeAttachmentRule

        CHANGELOG:
    #>

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
        $exception = Format-OagM365Exception -message "Failed retrieving DTP safe attachments policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
