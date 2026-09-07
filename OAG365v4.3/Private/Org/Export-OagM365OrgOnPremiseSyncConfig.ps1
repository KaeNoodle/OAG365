function Export-OagM365OrgOnPremiseSyncConfig {
    <#
    .SYNOPSIS
    Get details of Entra ID sync with on-premises AD

    .DESCRIPTION    
    Get details of Entra ID sync with on-premises AD

    .EXAMPLE
    Export-OagM365OrgOnPremiseSyncConfig

    .NOTES
        NAME: Export-OagM365OrgOnPremiseSyncConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryOnPremiseSynchronization

        CHANGELOG:    
    #>

    try {
        Write-Host "`nM365 tenant onPremises AD sync settings"

        # Get onPremise sync status
        $getMgDirectoryOnPremiseSync = Get-MgDirectoryOnPremiseSynchronization 
        $getMgDirectoryOnPremiseSyncCount = $getMgDirectoryOnPremiseSync.count
        Write-Host " - Found $getMgDirectoryOnPremiseSyncCount M365 onPremises sync settings"

        # Export to CSV
        $getMgDirectoryOnPremiseSync | ForEach-Object {
            $onPremisesSync = $_.Features
            $onPremisesSync | Add-Member -MemberType NoteProperty -Name Id -Value $_.Id
            $onPremisesSync
        } | Select-Object -Property * -ExcludeProperty AdditionalProperties | Export-CSV $script:exportTarget.onPremisesSync -NoTypeInformation
        Write-Host " - Exported M365 onPremises sync settings to CSV $($script:exportTarget.onPremisesSync)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed getting M365 directory sync config" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}
