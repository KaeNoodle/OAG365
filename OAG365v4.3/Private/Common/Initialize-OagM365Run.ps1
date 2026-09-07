function Initialize-OagM365Run {
    <#
    .SYNOPSIS
    Creates the run context: one run ID, one output folder tree, one transcript.

    .DESCRIPTION
    Called once by Export-M365.ps1 before any report executes.

    In the original scripts each of the five files independently called
    (Get-Date).ToString("yyyyMMdd_HHmmss") and created its own folder under \M365\. Running all
    five therefore produced five sibling folders with five different timestamps and nothing
    linking them. For audit purposes that is a problem: the evidence set could not be shown to
    come from a single engagement, and the transcripts were written beside the folders rather
    than inside them.

    This function generates the run ID once and returns a context object that every report
    writes through, so one execution produces one folder:

        <Output>\M365\<RunId>\
            00_RunLog\          transcript, parameters, environment, granted scopes, events
            ConditionalAccess\
            IAM\
            Organisation\
            ThreatProtection\
            Defender\

    .PARAMETER Output
    Root folder for the export. The M365\<RunId> tree is created beneath it.

    .PARAMETER Reports
    The reports selected for this run. Recorded in the run metadata.

    .PARAMETER AuthMode
    Which authentication parameter set was used. Recorded in the run metadata.

    .PARAMETER ModuleRoot
    Path to the module root, used for the integrity self-check.

    .EXAMPLE
    $Context = Initialize-OagM365Run -Output C:\Audit -Reports Cap,Iam -AuthMode Prompt

    .NOTES
        NAME: Initialize-OagM365Run
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Output,
        [Parameter(Mandatory = $true)][string[]]$Reports,
        [string]$AuthMode = 'Prompt',
        [string]$ModuleRoot
    )

    if (-not (Test-Path -Path $Output)) {
        New-Item -ItemType Directory -Force -Path $Output -ErrorAction Stop | Out-Null
    }
    $resolvedOutput = (Resolve-Path -Path $Output).Path

    $runId    = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $runRoot  = Join-Path -Path (Join-Path -Path $resolvedOutput -ChildPath 'M365') -ChildPath $runId
    $logPath  = Join-Path -Path $runRoot -ChildPath '00_RunLog'

    New-Item -ItemType Directory -Force -Path $logPath -ErrorAction Stop | Out-Null

    $Context = [PSCustomObject]@{
        RunId          = $runId
        RootPath       = $runRoot
        LogPath        = $logPath
        OutputRoot     = $resolvedOutput
        ModuleRoot     = $ModuleRoot
        StartTime      = Get-Date
        EndTime        = $null
        Reports        = $Reports
        AuthMode       = $AuthMode
        CurrentReport  = 'Startup'
        TenantId       = $null
        Account        = $null
        Events         = [System.Collections.Generic.List[object]]::new()
        Expected       = [System.Collections.Generic.List[object]]::new()
        ScopeLog       = [System.Collections.Generic.List[object]]::new()
        ReportStatus   = [System.Collections.Generic.List[object]]::new()
        TranscriptPath = Join-Path -Path $logPath -ChildPath 'transcript.txt'
    }

    Start-Transcript -Path $Context.TranscriptPath -ErrorAction SilentlyContinue | Out-Null

    Write-OagM365Log "OAG Microsoft 365 audit evidence export" -Context $Context
    Write-OagM365Log ("=" * 70) -Level Detail -Context $Context
    Write-OagM365Log "Run ID      : $runId" -Level Detail -Context $Context
    Write-OagM365Log "Started     : $($Context.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level Detail -Context $Context
    Write-OagM365Log "Output      : $runRoot" -Level Detail -Context $Context
    Write-OagM365Log "Reports     : $($Reports -join ', ')" -Level Detail -Context $Context
    Write-OagM365Log "Auth mode   : $AuthMode" -Level Detail -Context $Context
    Write-OagM365Log ("=" * 70) -Level Detail -Context $Context

    # ---------------------------------------------------------------------------------------
    # Environment capture. Recorded because the evidence is only interpretable if the reviewer
    # knows which module versions and which PowerShell produced it.
    # ---------------------------------------------------------------------------------------
    $environment = [PSCustomObject]@{
        RunId              = $runId
        StartTime          = $Context.StartTime.ToString('o')
        ComputerName       = $env:COMPUTERNAME
        UserName           = $env:USERNAME
        PowerShellVersion  = $PSVersionTable.PSVersion.ToString()
        PowerShellEdition  = $PSVersionTable.PSEdition
        OS                 = $PSVersionTable.OS
        LanguageMode       = $ExecutionContext.SessionState.LanguageMode.ToString()
        ModuleVersion      = $(
            # Null-safe: the module is normally imported by Export-M365.ps1, but the functions
            # can also be dot-sourced directly during development or testing. A missing version
            # should be recorded as unknown, not throw and abort the run.
            $m = Get-Module -Name 'OagExportM365' | Select-Object -First 1
            if ($m) { $m.Version.ToString() } else { 'Unknown (module not imported)' }
        )
        ReportsRequested   = $Reports -join ', '
        AuthMode           = $AuthMode
    }
    $environment | ConvertTo-Json -Depth 4 |
        Out-File -FilePath (Join-Path $logPath 'environment.json') -Encoding utf8

    Write-OagM365Log "PowerShell $($PSVersionTable.PSVersion) ($($ExecutionContext.SessionState.LanguageMode))" -Level Detail -Indent 1 -Context $Context

    # Constrained Language Mode is imposed by application control. The module genuinely cannot
    # run under it: [PSCustomObject], generic collections and .NET method calls are all blocked,
    # and those are load-bearing throughout both this module and the Microsoft Graph SDK.
    #
    # Stopping here is deliberate. Attempting to continue would produce a partial evidence set
    # with no clear indication that it was partial, which is the worst outcome for audit
    # evidence - worse than not running at all.
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
        $clmMessage = @"
PowerShell is running in $($ExecutionContext.SessionState.LanguageMode), not FullLanguage.

This module cannot run in this mode. Application control is restricting PowerShell, which
blocks [PSCustomObject], generic collections and .NET method calls. Those are used throughout
this module and the Microsoft Graph SDK, so this is not something the code can work around.

Diagnose the cause:
    .\Tools\Test-LanguageMode.ps1

The remedy depends on the cause and requires the application control administrator:
  - Run the module from a path approved by policy
  - Have the module signed by a certificate the policy trusts (this also satisfies the
    evidence integrity requirement - see Docs\Verification.md)
  - Request a time-limited exemption for the audit window

Do not attempt to bypass the restriction. Record it as a constraint on fieldwork if it
cannot be resolved in time.
"@
        throw $clmMessage
    }

    # Record the installed versions of every module the reports may use.
    $moduleInventory = Get-Module -ListAvailable -Name 'Microsoft.Graph*', 'ExchangeOnlineManagement' |
        Select-Object Name, @{n = 'Version'; e = { $_.Version.ToString() } }, Path |
        Sort-Object Name, Version
    $moduleInventory | Export-Csv -Path (Join-Path $logPath 'modules-available.csv') -NoTypeInformation -Encoding utf8
    Write-OagM365Log "Recorded $($moduleInventory.Count) available module version(s)" -Level Detail -Indent 1 -Context $Context

    # ---------------------------------------------------------------------------------------
    # Integrity self-check. Records which signed code produced this evidence.
    # ---------------------------------------------------------------------------------------
    if ($ModuleRoot) {
        $integrity = Test-OagM365Integrity -ModuleRoot $ModuleRoot -Context $Context
        $integrity | ConvertTo-Json -Depth 5 |
            Out-File -FilePath (Join-Path $logPath 'code-integrity.json') -Encoding utf8
    }

    return $Context
}
