function Export-OagM365DfoReport {
    <#
    .SYNOPSIS
    Exports Exchange Online Protection and Defender for Office 365 threat protection policies.

    .DESCRIPTION
    Covers anti-malware, anti-phishing, anti-spam, Safe Links, Safe Attachments and quarantine
    policies.

    This is the only report that uses Exchange Online rather than Microsoft Graph, so it
    connects separately. See Connect-OagM365ExchangeOnline for the temporary cmdlet module scope
    fix that this report depends on.

    .PARAMETER Context
    Run context object from Initialize-OagM365Run.

    .PARAMETER Connection
    Existing Exchange Online connection result from Connect-OagM365ExchangeOnline.

    .EXAMPLE
    Export-OagM365DfoReport -Context $Context -Connection $ExoConnection

    .NOTES
        NAME: Export-OagM365DfoReport
        VERSION: 4.0.0

        REQUIRED ROLE:
           Exchange Online View-Only Organization Management, or Global Reader

        MODULE:
           ExchangeOnlineManagement 3.9.x (3.10.1 has a known regression)
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$Connection
    )

    $folder = Initialize-OagM365ReportPath -Context $Context -Report 'Dfo'

    $status = [PSCustomObject]@{
        Report    = 'Dfo'
        Started   = Get-Date
        Completed = $null
        Succeeded = $false
        Folder    = $folder
    }

    try {
        # Each export is run independently so one failing policy type does not prevent the rest
        # being collected. The original ran them in sequence with no isolation.
        $exports = [ordered]@{
            'Anti-malware policies'    = { Export-OagM365DfoAntiMalwareConfig }
            'Anti-phishing policies'   = { Export-OagM365DfoAntiPhishingConfig }
            'Anti-spam policies'       = { Export-OagM365DfoAntiSpamConfig }
            'Safe Links policies'      = { Export-OagM365DfoSafeLinksConfig }
            'Safe Attachments policies'= { Export-OagM365DfoSafeAttachmentsConfig }
            'Quarantine policies'      = { Export-OagM365DfoQuarantinePolicy }
        }

        foreach ($name in $exports.Keys) {
            Write-OagM365Log "Exporting $name" -Context $Context
            try {
                & $exports[$name] | Out-Null
            } catch {
                $exception = Format-OagM365Exception -message "Failed exporting $name" -exception $_
                Write-OagM365Log $exception -Level Error -Indent 1 -Context $Context
            }
        }

        Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter '*.csv' -Description 'Threat protection CSV export'

        $status.Succeeded = $true

    } catch {
        $exception = Format-OagM365Exception -message "Defender for Office report failed" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
    } finally {
        $status.Completed = Get-Date
        $Context.ReportStatus.Add($status)
    }

    return $status
}
