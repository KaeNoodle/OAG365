<#--------------------------------------------------------------------------------------

    OAG-MainRunFile.ps1

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
    Resolves the report selection and works out which need Graph and which need Exchange.
    Connects to Graph once with the union of scopes the selected reports need, rather
      than prompting once per report.
    Runs each selected report.
    In the finally block: disconnects, checks completeness, writes the run summary, closes
      the transcript and then generates the manifest.

    The order at the end matters. The transcript is closed before the manifest is
    generated. A transcript that is still recording grows while the manifest is written,
    so its hash would change the instant after being taken and it would fail verification
    on every run. Routine false failures train reviewers to ignore the check.

    PARAMETERS
    -report (optional, one of: CAP, IAM, ORG, DFO, TH, All) accepts multiple values
    -output (optional) evidence folder, defaults to this script's folder
    -appClientId (optional) application ID for enterprise application authentication
    -appTenantId (optional) tenant ID for enterprise application authentication
    -appCertThumbprint (optional) certificate thumbprint, must be in the machine store
    -appSecret (optional) application secret, plain text
    -skipOnPremSync (optional) skip the ORG export that needs Global Administrator
    -nonInteractive (optional) suppress confirmation pauses, for unattended runs

    RUNNING CONTEXT
    Called by  : the operator
    Calls      : runInitialize, msGraphConnect, msExchangeOnlineConnect, the five report
                 writers, msDisconnect, exportVerify, manifestWrite, logWrite

    EXAMPLES
    .\OAG-MainRunFile.ps1
    .\OAG-MainRunFile.ps1 -report CAP,IAM
    .\OAG-MainRunFile.ps1 -report ORG -skipOnPremSync -output C:\Audit\2026
    .\OAG-MainRunFile.ps1 -appClientId "d359..." -appTenantId "a1b2..." -appCertThumbprint "A1B2..."

    CMLETS/PERMISSIONS/SCOPES
    See Documentation\Permissions.md for the full role and scope mapping per report.

    VERSION: 4.0.0
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
$selected = if ($report -contains 'All') { $allReports } else { $allReports | Where-Object { $_ -in $report } }
if (-not $selected) { throw "No valid reports selected." }

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

$run = runInitialize -output $output -reports $selected -authMode $PSCmdlet.ParameterSetName -moduleRoot $PSScriptRoot

try {
    $graphReports = $selected | Where-Object { $requirements[$_].service -eq 'Graph' }
    $exoReports   = $selected | Where-Object { $requirements[$_].service -eq 'ExchangeOnline' }

    if ($graphReports) {
        $scopes  = $graphReports | ForEach-Object { $requirements[$_].scopes }  | Sort-Object -Unique
        $modules = $graphReports | ForEach-Object { $requirements[$_].modules } | Sort-Object -Unique

        $connectArgs = @{ scopes = $scopes; modules = $modules; noPause = $nonInteractive }
        switch ($PSCmdlet.ParameterSetName) {
            'AppCertThumbprint' { $connectArgs += @{ appClientId = $appClientId; appTenantId = $appTenantId; appCertThumbprint = $appCertThumbprint } }
            'AppSecret'         { $connectArgs += @{ appClientId = $appClientId; appTenantId = $appTenantId; appSecret = $appSecret } }
        }

        $graph = msGraphConnect @connectArgs
        if (-not $graph.connected) { throw "Could not connect to Microsoft Graph. No Graph reports can run." }
    }

    if ($exoReports) {
        $exo = msExchangeOnlineConnect -noPause:$nonInteractive
        if (-not $exo.connected) { logWrite "Could not connect to Exchange Online. The DFO report will be skipped." -level Error }
    }

    foreach ($name in $selected) {
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
                if ($exo -and $exo.connected) { dfoReportWrite | Out-Null }
                else { logWrite "Skipped - no Exchange Online connection." -level Warning -indent 1 }
            }
        }
    }

} catch {
    logWrite "RUN FAILED: $($_.Exception.Message)" -level Error
    logWrite $_.ScriptStackTrace -level Detail

} finally {
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
        reportsRequested = $selected -join ', '
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
