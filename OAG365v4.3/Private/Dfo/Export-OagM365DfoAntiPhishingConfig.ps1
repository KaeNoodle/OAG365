function Export-OagM365DfoAntiPhishingConfig {
    <#
    .SYNOPSIS
    Export anti-malware policy and rule threat protection policies

    .DESCRIPTION
    Export anti-malware policy and rule settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-OagM365DfoAntiPhishingConfig

    .NOTES
        NAME: Export-OagM365DfoAntiPhishingConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-AntiPhishPolicy
           Get-AntiPhishRule

        CHANGELOG:
    #>

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
        $exception = Format-OagM365Exception -message "Failed retrieving EOP anti-phishing policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
