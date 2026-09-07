function runInitialize {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Starts a run. Creates one run ID, one output folder tree and one transcript, and
    stores the run state in $script:run where every other function can read it.

    In the original scripts each of the five files generated its own timestamp and its
    own folder, so running all five produced five sibling folders with nothing linking
    them. One execution now produces one folder.

    LOGIC
    Resolves the output path and builds M365\<runId>\00_RunLog beneath it.
    Stores run state in $script:run so no function needs a context parameter.
    Starts the transcript inside the run folder rather than beside it.
    Records the environment: PowerShell version, language mode, module versions.
    Stops with an error if the session is in Constrained Language Mode, because the
      module cannot produce complete output there and should not pretend otherwise.
    Runs codeVerify to record which signed code is producing this evidence.

    Collections are plain arrays rather than generic lists. Generic list constructors
    are blocked in Constrained Language Mode, and at this scale the difference in
    speed does not matter.

    PARAMETERS
    -output (optional) root folder for evidence, defaults to the script folder
    -reports (optional) which reports this run covers, recorded in the metadata
    -authMode (optional) which authentication method was used
    -moduleRoot (optional) module path, used for the code integrity check

    RUNNING CONTEXT
    Called by  : OAG-MainRunFile.ps1, or by a report writer if run on its own
    Calls      : logWrite, codeVerify
    Sets       : $script:run

    CMLETS/PERMISSIONS/SCOPES
    None. Local filesystem only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        $output     = $PSScriptRoot,
        [string[]]$reports = @('Solo'),
        [string]$authMode  = 'Prompt',
        [string]$moduleRoot
    )

    if (-not (Test-Path -Path $output)) {
        New-Item -ItemType Directory -Force -Path $output -ErrorAction Stop | Out-Null
    }
    $outputRoot = (Resolve-Path -Path $output).Path

    $runId   = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $runRoot = Join-Path (Join-Path $outputRoot 'M365') $runId
    $logPath = Join-Path $runRoot '00_RunLog'

    New-Item -ItemType Directory -Force -Path $logPath -ErrorAction Stop | Out-Null

    $script:run = [PSCustomObject]@{
        id             = $runId
        rootPath       = $runRoot
        logPath        = $logPath
        outputRoot     = $outputRoot
        moduleRoot     = $moduleRoot
        startTime      = Get-Date
        endTime        = $null
        reports        = $reports
        authMode       = $authMode
        currentReport  = 'Startup'
        tenantId       = $null
        account        = $null
        events         = @()
        expected       = @()
        scopes         = @()
        status         = @()
        transcriptPath = Join-Path $logPath 'transcript.txt'
    }

    Start-Transcript -Path $script:run.transcriptPath -ErrorAction SilentlyContinue | Out-Null

    logWrite "OAG Microsoft 365 audit evidence export"
    logWrite ("=" * 70) -level Detail
    logWrite "Run ID    : $runId" -level Detail
    logWrite "Started   : $($script:run.startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -level Detail
    logWrite "Output    : $runRoot" -level Detail
    logWrite "Reports   : $($reports -join ', ')" -level Detail
    logWrite "Auth mode : $authMode" -level Detail
    logWrite ("=" * 70) -level Detail

    $module = Get-Module -Name 'OAG-FileInfo' | Select-Object -First 1

    [PSCustomObject]@{
        runId             = $runId
        startTime         = $script:run.startTime.ToString('o')
        computerName      = $env:COMPUTERNAME
        userName          = $env:USERNAME
        powerShellVersion = $PSVersionTable.PSVersion.ToString()
        powerShellEdition = $PSVersionTable.PSEdition
        operatingSystem   = $PSVersionTable.OS
        languageMode      = $ExecutionContext.SessionState.LanguageMode.ToString()
        moduleVersion     = if ($module) { $module.Version.ToString() } else { 'not imported' }
        reportsRequested  = $reports -join ', '
        authMode          = $authMode
    } | ConvertTo-Json -Depth 4 | Out-File (Join-Path $logPath 'environment.json') -Encoding utf8

    logWrite "PowerShell $($PSVersionTable.PSVersion) ($($ExecutionContext.SessionState.LanguageMode))" -level Detail -indent 1

    # Constrained Language Mode genuinely prevents this module working. The original
    # printed "Abort now" in red and then carried on regardless; this actually stops.
    if ($ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage') {
        throw "Cannot run in Constrained Language Mode. The module files must be signed with a certificate trusted by the WDAC policy on this machine. See Documentation\Verification.md."
    }

    Get-Module -ListAvailable -Name 'Microsoft.Graph*', 'ExchangeOnlineManagement' |
        Select-Object Name, @{n = 'Version'; e = { $_.Version.ToString() } }, Path |
        Sort-Object Name, Version |
        Export-Csv (Join-Path $logPath 'modules-available.csv') -NoTypeInformation -Encoding utf8

    if ($moduleRoot) {
        codeVerify -moduleRoot $moduleRoot |
            ConvertTo-Json -Depth 5 |
            Out-File (Join-Path $logPath 'code-integrity.json') -Encoding utf8
    }

    return $script:run
}
