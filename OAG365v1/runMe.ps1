<#--------------------------------------------------------------------------------------

    runMe.ps1

    DESCRIPTION
    Single entry point for the five M365 audit evidence reports. Replaces the five
    standalone scripts, which shared around 1,200 lines of duplicated code and each
    produced their own separately timestamped output folder.

    Reports remain independently runnable. That was deliberate in the original design and
    is kept, because they have genuinely different permission requirements: the
    organisation report needs Global Administrator for one export, the Defender for Office
    report uses Exchange Online rather than Graph, and threat hunting needs a scope that
    requires tenant administrator consent. Forcing them into one run would mean a single
    missing permission blocks everything.

    Any report can also be called directly without this file. Each one calls runEnsure,
    which starts a run if none exists, so no context has to be passed around.

    LOGIC
    Imports the module. RequiredModules in the manifest enforces the PowerShell and
      dependency versions, so a version problem stops here rather than surfacing as a
      confusing failure part way through.
    If -report was not explicitly passed (and -nonInteractive was not set), drops into a
      looping menu (Reports / Tools / Help / Exit) instead of silently defaulting to
      running every report once. Passing -report keeps the script fully scriptable and
      unattended, exactly as before - the menu never appears in that case.
    Otherwise resolves the report selection and works out which need Graph and which need
      Exchange, connects once with the union of scopes needed, and runs each selected
      report.
    In the interactive menu, the same run and the same Graph/EXO connections are reused
      across every report picked in the session - connecting again only for a service not
      already connected. Choosing Exit is what finalizes the run.
    In the finally block: disconnects, checks completeness, writes the run summary, closes
      the transcript and then generates the manifest.

    The order at the end matters. The transcript is closed before the manifest is
    generated. A transcript that is still recording grows while the manifest is written,
    so its hash would change the instant after being taken and it would fail verification
    on every run. Routine false failures train reviewers to ignore the check.

    PARAMETERS
    -report (optional, one of: CAP, IAM, ORG, DFO, TH, All) accepts multiple values. Omit
      it (without -nonInteractive) to get the interactive menu instead.
    -output (optional) evidence folder, defaults to this script's folder
    -appClientId (optional) application ID for enterprise application authentication
    -appTenantId (optional) tenant ID for enterprise application authentication
    -appCertThumbprint (optional) certificate thumbprint, must be in the machine store
    -appSecret (optional) application secret, plain text
    -skipOnPremSync (optional) skip the ORG export that needs Global Administrator
    -nonInteractive (optional) suppress confirmation pauses and the menu, for unattended runs

    RUNNING CONTEXT
    Called by  : the operator
    Calls      : runInitialize, msGraphConnect, msExchangeOnlineConnect, the five report
                 writers, msDisconnect, exportVerify, manifestWrite, logWrite

    EXAMPLES
    .\runMe.ps1
    .\runMe.ps1 -report CAP,IAM
    .\runMe.ps1 -report ORG -skipOnPremSync -output C:\Audit\2026
    .\runMe.ps1 -appClientId "d359..." -appTenantId "a1b2..." -appCertThumbprint "A1B2..."

    CMLETS/PERMISSIONS/SCOPES
    See Documentation\Permissions.md for the full role and scope mapping per report.

    VERSION: 4.1.0
    RELEASE APPROVAL: <UNAPPROVED>

    --------------------------------------------------------------------------------------#>

[CmdletBinding(DefaultParameterSetName = 'Prompt')]
Param(
    [ValidateSet('CAP', 'IAM', 'ORG', 'DFO', 'TH', 'All')][string[]]$report = 'All',
    $output = $PSScriptRoot,

    [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appTenantId,
    [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')][string]$appCertThumbprint,
    [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appSecret,

    [switch]$skipOnPremSync,
    [switch]$nonInteractive
)

$ErrorActionPreference = 'Stop'

# -report falls back to 'All' via ValidateSet's default, so $PSBoundParameters is the only
# reliable way to tell "the operator typed nothing" apart from "it resolved to All". An
# unattended invocation that forgot -report still needs to run rather than block on a menu
# with nobody at the keyboard, so -nonInteractive also suppresses it.
$reportExplicitlyBound = $PSBoundParameters.ContainsKey('report')
$showMenu = (-not $reportExplicitlyBound) -and (-not $nonInteractive)

# Files copied from a network location or downloaded carry Mark-of-the-Web, which
# silently blocks module import and would otherwise need each file unblocked by hand
# through a security prompt. Unblock-File only exists on Windows, so this is a no-op
# elsewhere, and never stops the run if it can't clear a file (e.g. ACL restrictions).
if (Get-Command Unblock-File -ErrorAction SilentlyContinue) {
    Get-ChildItem -Path $PSScriptRoot -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
}

$manifestPath = Join-Path $PSScriptRoot 'OAG-ModuleManifest.psd1'
if (-not (Test-Path $manifestPath)) {
    throw "Cannot find OAG-ModuleManifest.psd1 in $PSScriptRoot. Run this from the module folder."
}
Import-Module -Name $manifestPath -Force -ErrorAction Stop

$allReports = @('CAP', 'IAM', 'ORG', 'DFO', 'TH')

# ----------------------------------------------------------------------------------------
# Per-report requirements. Held here so the union of scopes can be requested in one login.
# ----------------------------------------------------------------------------------------
$requirements = @{
    'CAP' = @{ service = 'Graph'
        scopes  = @('Policy.Read.AuthenticationMethod','Policy.Read.All','AuthenticationContext.Read.All',
                    'Policy.Read.ConditionalAccess','RoleManagement.Read.Directory','Application.Read.All','Directory.Read.All')
        modules = @('Microsoft.Graph.Authentication','Microsoft.Graph.DirectoryObjects','Microsoft.Graph.Identity.SignIns',
                    'Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Applications') }
    'IAM' = @{ service = 'Graph'
        scopes  = @('Directory.Read.All','DeviceManagementApps.Read.All','Device.Read.All','Application.Read.All',
                    'AuditLog.Read.All','RoleManagement.Read.Directory',
                    'PrivilegedEligibilitySchedule.Read.AzureADGroup','PrivilegedAssignmentSchedule.Read.AzureADGroup')
        modules = @('Microsoft.Graph.Authentication','Microsoft.Graph.DirectoryObjects','Microsoft.Graph.Users',
                    'Microsoft.Graph.Groups','Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Applications',
                    'Microsoft.Graph.Reports','Microsoft.Graph.Identity.Governance') }
    'ORG' = @{ service = 'Graph'
        scopes  = @('Organization.Read.All','Directory.Read.All','Domain.Read.All','OnPremDirectorySynchronization.Read.All')
        modules = @('Microsoft.Graph.Authentication','Microsoft.Graph.DirectoryObjects','Microsoft.Graph.Identity.DirectoryManagement') }
    'DFO' = @{ service = 'ExchangeOnline'; scopes = @(); modules = @('ExchangeOnlineManagement') }
    'TH'  = @{ service = 'Graph'; scopes = @('ThreatHunting.Read.All'); modules = @('Microsoft.Graph.Authentication') }
}

# ----------------------------------------------------------------------------------------
# Session state for the interactive menu - lets a service already connected in this
# session be reused by a later report pick instead of reconnecting every loop iteration.
# ----------------------------------------------------------------------------------------
$script:graphConnected = $false
$script:exoConnected   = $false
$script:graph          = $null
$script:exo            = $null
$script:sessionReports = @()

# ----------------------------------------------------------------------------------------
# Connects to whatever service the given reports need (skipping a service already
# connected this session) and runs them. Shared by the scripted and menu paths so there is
# exactly one copy of the connect/dispatch logic.
# ----------------------------------------------------------------------------------------
function invokeReports {
    param([Parameter(Mandatory = $true)][string[]]$names)

    $graphNeeded = $names | Where-Object { $requirements[$_].service -eq 'Graph' }
    $exoNeeded   = $names | Where-Object { $requirements[$_].service -eq 'ExchangeOnline' }

    if ($graphNeeded -and -not $script:graphConnected) {
        # An interactive session can pick more reports later, so its first Graph login
        # asks for every Graph report's scopes up front rather than just what's picked so
        # far - one login covers whatever comes next. A scripted run keeps asking only for
        # what was actually selected, exactly as before the menu existed.
        $scopeSource = if ($showMenu) { $allReports | Where-Object { $requirements[$_].service -eq 'Graph' } } else { $graphNeeded }
        $scopes  = $scopeSource | ForEach-Object { $requirements[$_].scopes }  | Sort-Object -Unique
        $modules = $scopeSource | ForEach-Object { $requirements[$_].modules } | Sort-Object -Unique

        $connectArgs = @{ scopes = $scopes; modules = $modules; noPause = $nonInteractive }
        switch ($PSCmdlet.ParameterSetName) {
            'AppCertThumbprint' { $connectArgs += @{ appClientId = $appClientId; appTenantId = $appTenantId; appCertThumbprint = $appCertThumbprint } }
            'AppSecret'         { $connectArgs += @{ appClientId = $appClientId; appTenantId = $appTenantId; appSecret = $appSecret } }
        }

        $script:graph = msGraphConnect @connectArgs
        if (-not $script:graph.connected) { throw "Could not connect to Microsoft Graph. No Graph reports can run." }
        $script:graphConnected = $true
    } elseif ($graphNeeded) {
        logWrite "Reusing existing Microsoft Graph connection." -level Detail -indent 1
    }

    if ($exoNeeded -and -not $script:exoConnected) {
        $script:exo = msExchangeOnlineConnect -noPause:$nonInteractive
        if ($script:exo.connected) { $script:exoConnected = $true }
        else { logWrite "Could not connect to Exchange Online. The DFO report will be skipped." -level Error }
    } elseif ($exoNeeded) {
        logWrite "Reusing existing Exchange Online connection." -level Detail -indent 1
    }

    foreach ($name in $names) {
        logWrite ""
        logWrite ("-" * 70) -level Detail
        logWrite "REPORT: $name"
        logWrite ("-" * 70) -level Detail

        switch ($name) {
            'CAP' { capReportWrite | Out-Null }
            'IAM' { iamReportWrite | Out-Null }
            'ORG' { orgReportWrite -skipOnPremSync:$skipOnPremSync | Out-Null }
            'TH'  { threatHuntReportWrite | Out-Null }
            'DFO' {
                if ($script:exo -and $script:exo.connected) { dfoReportWrite | Out-Null }
                else { logWrite "Skipped - no Exchange Online connection." -level Warning -indent 1 }
            }
        }
    }
}

# ----------------------------------------------------------------------------------------
# Interactive menu helpers. Only reached when $showMenu is true.
# ----------------------------------------------------------------------------------------
function showMainMenu {
    Write-Host ""
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host " OAG M365 Audit Evidence Export" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host " 1) Reports"
    Write-Host " 2) Tools"
    Write-Host " 3) Help"
    Write-Host " 0) Exit"
    Write-Host ("=" * 50) -ForegroundColor Cyan

    switch (Read-Host "Choose an option") {
        '1' { 'Reports' }
        '2' { 'Tools' }
        '3' { 'Help' }
        '0' { 'Exit' }
        default { Write-Host "Not a valid option." -ForegroundColor Red; 'Retry' }
    }
}

function showReportsMenu {
    Write-Host ""
    Write-Host "Run which report(s)?" -ForegroundColor Cyan
    for ($i = 0; $i -lt $allReports.Count; $i++) {
        Write-Host " $($i + 1)) $($allReports[$i])"
    }
    Write-Host " $($allReports.Count + 1)) All"
    Write-Host " 0) Back"

    $choice = Read-Host "Choose one or more (comma-separated), or 0 to go back"
    if ([string]::IsNullOrWhiteSpace($choice) -or $choice.Trim() -eq '0') { return @() }

    $picked = @()
    foreach ($token in ($choice -split ',')) {
        $token = $token.Trim()
        if ($token -eq "$($allReports.Count + 1)") { return $allReports }
        $index = 0
        if ([int]::TryParse($token, [ref]$index) -and $index -ge 1 -and $index -le $allReports.Count) {
            $picked += $allReports[$index - 1]
        }
    }
    return ($picked | Sort-Object -Unique)
}

function showToolsMenu {
    Write-Host ""
    Write-Host "Tools" -ForegroundColor Cyan
    Write-Host " 1) Verify-Export - check a completed run's evidence integrity"
    Write-Host " 2) Debugger      - diagnose signing, execution policy and module issues"
    Write-Host " 0) Back"

    switch (Read-Host "Choose an option") {
        '1' {
            $m365Root = Join-Path $output 'M365'
            $default = $null
            if (Test-Path $m365Root) {
                $default = Get-ChildItem -Path $m365Root -Directory -ErrorAction SilentlyContinue |
                    Where-Object { Test-Path (Join-Path $_.FullName 'manifest.sha256') } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
            }
            $prompt = if ($default) { "Run folder to verify [$default]" } else { "Run folder to verify" }
            $folder = Read-Host $prompt
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = $default }
            if ($folder) {
                & (Join-Path $PSScriptRoot 'Tools\Verify-Export.ps1') -runFolder $folder
            } else {
                Write-Host "No completed run found to verify, and no folder given." -ForegroundColor Yellow
            }
        }
        '2' { & (Join-Path $PSScriptRoot 'Tools\Debugger.ps1') -ModulePath $PSScriptRoot }
    }
}

function showHelpScreen {
    Write-Host ""
    Write-Host "Help" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host "CAP  Conditional access, authentication method and strength policies"
    Write-Host "IAM  Users, devices, service principals, auth methods, admin units, PIM roles"
    Write-Host "ORG  Tenant configuration, on-premises sync, domains, DNS auth records"
    Write-Host "DFO  Defender for Office 365 anti-phishing, anti-spam, anti-malware, quarantine"
    Write-Host "TH   Defender advanced hunting - threat, health and software inventory reports"
    Write-Host ""
    Write-Host "Key parameters: -output, -skipOnPremSync, -nonInteractive, or the app-auth" -ForegroundColor Gray
    Write-Host "parameters (-appClientId/-appTenantId with -appCertThumbprint or -appSecret)." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Evidence for this session is written under: $(Join-Path $output 'M365')" -ForegroundColor Gray
    Write-Host "See Documentation\README.md for full details." -ForegroundColor Gray
}

# ----------------------------------------------------------------------------------------
# Shared shutdown: disconnects, checks completeness, writes the run summary, closes the
# transcript and generates the manifest. Called exactly once, whichever path is taken.
# ----------------------------------------------------------------------------------------
function finalizeRun {
    # Runs even on failure, so a partial evidence set is still hashed, logged and marked
    # incomplete rather than left unexplained.
    $script:run.currentReport = 'Shutdown'
    $script:run.endTime = Get-Date

    logWrite ""
    logWrite ("=" * 70) -level Detail
    msDisconnect

    if ($script:run.scopes.Count -gt 0) {
        $script:run.scopes | Export-Csv (Join-Path $script:run.logPath 'scopes-granted.csv') -NoTypeInformation -Encoding utf8
    }

    $completeness = exportVerify

    if ($script:run.events.Count -gt 0) {
        $script:run.events | Export-Csv (Join-Path $script:run.logPath 'events.csv') -NoTypeInformation -Encoding utf8
    }

    [PSCustomObject]@{
        runId = $script:run.id
        started = $script:run.startTime.ToString('o'); completed = $script:run.endTime.ToString('o')
        durationMinutes = [math]::Round(($script:run.endTime - $script:run.startTime).TotalMinutes, 2)
        tenantId = $script:run.tenantId; account = $script:run.account; authMode = $script:run.authMode
        reportsRequested = ($script:sessionReports | Sort-Object -Unique) -join ', '
        reportsSucceeded = (($script:run.status | Where-Object succeeded) | ForEach-Object { $_.report }) -join ', '
        completenessStatus = $completeness.status
        filesExpected = $completeness.filesExpected; filesOk = $completeness.filesOk
        filesEmpty = $completeness.filesEmpty; filesMissing = $completeness.filesMissing
        errors = $completeness.errorCount; warnings = $completeness.warningCount
    } | ConvertTo-Json -Depth 4 | Out-File (Join-Path $script:run.logPath 'run-summary.json') -Encoding utf8

    logWrite ""
    logWrite "RUN SUMMARY"
    logWrite "Run ID   : $($script:run.id)" -level Detail -indent 1
    logWrite "Reports  : $((($script:run.status | Where-Object succeeded) | ForEach-Object { $_.report }) -join ', ')" -level Detail -indent 1
    logWrite "Files    : $($completeness.filesOk) OK, $($completeness.filesEmpty) empty, $($completeness.filesMissing) missing" -level Detail -indent 1
    logWrite "Status   : $($completeness.status)" -level $(if ($completeness.status -eq 'Complete') { 'Success' } else { 'Warning' }) -indent 1

    # Transcript closed BEFORE the manifest - see LOGIC in the header.
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    manifestWrite | Out-Null

    Write-Host ""
    Write-Host "Evidence written to: $($script:run.rootPath)" -ForegroundColor Cyan
    Write-Host "Verify with        : .\Tools\Verify-Export.ps1 -runFolder '$($script:run.rootPath)'" -ForegroundColor Gray
}

if (-not $showMenu) {
    # ------------------------------------------------------------------------------------
    # Scripted / unattended path - unchanged behaviour from before the menu existed.
    # ------------------------------------------------------------------------------------
    $selected = if ($report -contains 'All') { $allReports } else { $allReports | Where-Object { $_ -in $report } }
    if (-not $selected) { throw "No valid reports selected." }
    $script:sessionReports = $selected

    $run = runInitialize -output $output -reports $selected -authMode $PSCmdlet.ParameterSetName -moduleRoot $PSScriptRoot

    try {
        invokeReports -names $selected
    } catch {
        logWrite "RUN FAILED: $($_.Exception.Message)" -level Error
        logWrite $_.ScriptStackTrace -level Detail
    } finally {
        finalizeRun
    }

} else {
    # ------------------------------------------------------------------------------------
    # Interactive menu path - one run, reused connections, loops until Exit.
    # ------------------------------------------------------------------------------------
    $run = runInitialize -output $output -reports @('Interactive') -authMode $PSCmdlet.ParameterSetName -moduleRoot $PSScriptRoot

    try {
        $exit = $false
        while (-not $exit) {
            switch (showMainMenu) {
                'Reports' {
                    $picked = showReportsMenu
                    if ($picked) {
                        try {
                            invokeReports -names $picked
                            $script:sessionReports += $picked
                        } catch {
                            logWrite "REPORT RUN FAILED: $($_.Exception.Message)" -level Error
                            logWrite $_.ScriptStackTrace -level Detail
                        }
                    }
                }
                'Tools' { showToolsMenu }
                'Help'  { showHelpScreen }
                'Exit'  { $exit = $true }
                'Retry' { }
            }
        }
    } catch {
        logWrite "RUN FAILED: $($_.Exception.Message)" -level Error
        logWrite $_.ScriptStackTrace -level Detail
    } finally {
        finalizeRun
    }
}
