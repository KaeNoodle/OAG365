function Export-OagM365DfoSafeLinksConfig {
    <#
    .SYNOPSIS
    Export safe links policy threat protection policies

    .DESCRIPTION
    Export safe links policy settings from Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    Export-OagM365DfoSafeLinksConfig

    .NOTES
        NAME: Export-OagM365DfoSafeLinksConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeLinksPolicy
           Get-SafeLinksRule

        CHANGELOG:
    #>

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
        $exception = Format-OagM365Exception -message "Failed retrieving DTP safe links policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
