function Export-OagM365IamDevice {
    <#
    .SYNOPSIS
    Exports Microsoft Entra managed devices

    .DESCRIPTION
    Exports Microsoft Entra managed devices

    .EXAMPLE
    Export-OagM365IamDevice

    .NOTES
        NAME: Export-OagM365IamDevice
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDevice

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID devices"

        # Get Entra devices
        $getMgDevice = Get-MgDevice -All
        $getMgDeviceCount = $getMgDevice.count
        Write-Host " - Found $getMgDeviceCount Entra device records"

        # Export to CSV
        $getMgDevice | Select-Object -Property $script:deviceExportProperties | Export-CSV $script:exportTarget.entraDevices -NoTypeInformation
        Write-Host " - Exported Entra devices to CSV $($script:exportTarget.entraDevices)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-OagM365Exception -message "Failed retrieving M365 devices" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }

}
