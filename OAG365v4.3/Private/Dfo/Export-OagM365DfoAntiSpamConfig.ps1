function Export-OagM365DfoAntiSpamConfig {
    <#
    .SYNOPSIS
    Export anti-spam policy threat protection policies

    .DESCRIPTION
    Export anti-spam policy settings from Exchange Online / Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    Export-OagM365DfoAntiSpamConfig

    .NOTES
        NAME: Export-OagM365DfoAntiSpamConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-HostedConnectionFilterPolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-spam protection settings"

        # Get EOP spam policy settings
        $getHostedConnectionFilterPolicy = Get-HostedConnectionFilterPolicy 
        $getHostedConnectionFilterPolicy | Export-CSV $script:exportTarget.eopSpamPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-spam protection policy settings to $($script:exportTarget.eopSpamPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving EOP anti-spam policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
