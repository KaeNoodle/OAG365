function runEnsure {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Makes sure a run exists before a report writer starts work.

    This is what lets a report be run on its own. Without it, every function would need
    a run context passed in as a parameter, which was noisy and made single-report use
    awkward. Instead each report writer calls this first: if the main run file already
    started a run, nothing happens; if not, one is started with sensible defaults.

    LOGIC
    Returns immediately if $script:run already exists.
    Otherwise calls runInitialize against the module folder and says so, so the
      operator can see a run was created rather than joined.

    PARAMETERS
    -report (optional) name recorded as the reason the run was created

    RUNNING CONTEXT
    Called by  : capReportWrite, iamReportWrite, orgReportWrite, dfoReportWrite,
                 threatHuntReportWrite
    Calls      : runInitialize, logWrite

    CMLETS/PERMISSIONS/SCOPES
    None.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [string]$report = 'Solo'
    )

    if ($null -ne $script:run) { return }

    $root = Split-Path -Path $PSScriptRoot -Parent
    runInitialize -output $root -reports @($report) -moduleRoot $root | Out-Null
    logWrite "No run was active, so one was started for this report." -level Detail -indent 1
}
