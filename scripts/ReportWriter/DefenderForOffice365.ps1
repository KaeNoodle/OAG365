function dfoReportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Exports Exchange Online Protection and Defender for Office 365 threat protection
    policies: anti-malware, anti-phishing, anti-spam, Safe Links, Safe Attachments and
    quarantine.

    The only report that uses Exchange Online rather than Microsoft Graph, so it
    authenticates separately. Running it alongside Graph reports produces two logins.

    LOGIC
    Makes sure a run exists, then creates the ThreatProtection folder.
    Connects to Exchange Online if not already connected, so this report works on its own.
      The connection result is assigned and tested properly. The original wrote
      if (Connect-... -eq $true), which PowerShell parses as passing two positional
      arguments to the function rather than comparing its return value, so the branch
      never tested what it appeared to.
    Runs each policy export inside its own try/catch, so one failing policy type does not
      stop the rest being collected. The original ran them in sequence with no isolation.

    PARAMETERS
    None.

    RUNNING CONTEXT
    Called by  : runMe.ps1 -report DFO, or run directly
    Calls      : runEnsure, reportPathInitialize, msExchangeOnlineConnect,
                 dfoAntiMalwareExport, dfoAntiPhishingExport, dfoAntiSpamExport,
                 dfoSafeLinksExport, dfoSafeAttachmentsExport, dfoQuarantineExport,
                 exportRegister, logWrite, exceptionFormat

    CMLETS/PERMISSIONS/SCOPES
    Role   : Exchange Online View-Only Organization Management, or Global Reader
    Module : ExchangeOnlineManagement 3.9.x - 3.10.1 has a known regression
    Scopes : none, this report does not use Graph

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    runEnsure -report 'DFO'
    $folder = reportPathInitialize -report 'DFO'
    $started = Get-Date

    try {
        # Connect only if a session is not already open, so running this report on its
        # own works and running it from the main file does not prompt twice.
        # Get-ConnectionInformation is the natural test, but on module versions where
        # Get-ConnectionContext is broken it throws outright rather than returning nothing,
        # and -ErrorAction cannot suppress a thrown exception. That ended this report
        # before it started, on a session that was working. Treat a throw as "cannot tell
        # from here" and fall back to whether the cmdlets this report needs are present,
        # which is the thing that actually matters.
        $connected = $false
        try {
            if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                $connected = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue)
            }
        } catch {
            $connected = $false
        }
        if (-not $connected) {
            $connected = [bool](Get-Command -Name 'Get-MalwareFilterPolicy' -ErrorAction SilentlyContinue)
        }
        if (-not $connected) {
            $connection = msExchangeOnlineConnect
            if (-not $connection.connected) {
                logWrite "No Exchange Online connection. This report cannot run." -level Error -indent 1
                $script:run.status += [PSCustomObject]@{ report = 'DFO'; started = $started; completed = Get-Date; succeeded = $false }
                return $false
            }
        }

        $exports = [ordered]@{
            'Anti-malware policies'     = { dfoAntiMalwareExport }
            'Anti-phishing policies'    = { dfoAntiPhishingExport }
            'Anti-spam policies'        = { dfoAntiSpamExport }
            'Safe Links policies'       = { dfoSafeLinksExport }
            'Safe Attachments policies' = { dfoSafeAttachmentsExport }
            'Quarantine policies'       = { dfoQuarantineExport }
        }

        foreach ($name in $exports.Keys) {
            logWrite "Exporting $name"
            try {
                & $exports[$name] | Out-Null
            } catch {
                logWrite (exceptionFormat -message "Failed exporting $name" -exception $_) -level Error -indent 1
            }
        }

        exportRegister -folder $folder -filter '*.csv' -description 'Threat protection CSV export'

        $script:run.status += [PSCustomObject]@{ report = 'DFO'; started = $started; completed = Get-Date; succeeded = $true }
        return $true

    } catch {
        logWrite (exceptionFormat -message "Defender for Office report failed" -exception $_) -level Error
        $script:run.status += [PSCustomObject]@{ report = 'DFO'; started = $started; completed = Get-Date; succeeded = $false }
        return $false
    }
}
