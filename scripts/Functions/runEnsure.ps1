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
    Otherwise calls runInitialize with the module folder as moduleRoot (for the code
      integrity check) but the folder next to it as the output root, so a solo report run
      writes its evidence beside the module rather than inside it - the same reasoning
      runMe.ps1 applies to its own default -output.

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

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $outputRoot = Split-Path -Path $moduleRoot -Parent
    runInitialize -output $outputRoot -reports @($report) -moduleRoot $moduleRoot | Out-Null
    logWrite "No run was active, so one was started for this report." -level Detail -indent 1
}
