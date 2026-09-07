<#
.SYNOPSIS
Exports Microsoft 365 and Entra ID configuration as audit evidence.

.DESCRIPTION
Single entry point for all five M365 export reports. Replaces the five standalone scripts
(Export-M365-CapReport.ps1 and siblings), which shared roughly 1,200 lines of duplicated code
and each produced its own separately timestamped output folder.

Reports remain independently runnable. That was deliberate in the original design and is
preserved here, because the reports have genuinely different permission requirements: the
organisation report needs Global Administrator for one export, the Defender for Office report
uses Exchange Online rather than Graph, and threat hunting needs a scope that requires tenant
administrator consent. Merging them into one mandatory run would mean a single missing
permission blocks everything.

    Reports:
      Cap         Conditional access, authentication method and strength policies
      Iam         Users, devices, service principals, auth methods, admin units, PIM roles
      Org         Tenant configuration, domains, MX / SPF / DKIM / DMARC
      Dfo         Exchange Online Protection and Defender for Office 365 policies
      ThreatHunt  Defender advanced hunting: vulnerabilities, agent health, software inventory

    Output:
      <Output>\M365\<RunId>\
          00_RunLog\           transcript, environment, scopes, events, completeness
          ConditionalAccess\   IAM\   Organisation\   ThreatProtection\   Defender\
          manifest.sha256      SHA-256 of every file produced

.PARAMETER Report
Which reports to run. Defaults to All. Accepts multiple values.

.PARAMETER Output
Folder to write evidence to. Defaults to the script's own folder.

.PARAMETER AppClientId
Registered Entra application's client / application ID, for enterprise application
authentication.

.PARAMETER AppTenantId
Registered Entra application's tenant ID.

.PARAMETER AppCertThumbprint
Thumbprint of the certificate used to authenticate as the registered application. The
certificate must be installed in the machine certificate store.

Create one in an elevated session:
  $cert = New-SelfSignedCertificate -Subject "CN=OagExportM365" -CertStoreLocation "Cert:\LocalMachine\My" `
            -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA `
            -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths(3)
  Export-Certificate -Cert $cert -FilePath "./OagExportScript.cer"

.PARAMETER AppSecret
Registered Entra application's secret, in plain text.

.PARAMETER SkipOnPremSync
Skip the on-premises directory synchronisation export in the Org report. That export requires
Global Administrator and returns 403 under read-only credentials, so skipping it keeps the run
log clean where the gap is already documented.

.PARAMETER NonInteractive
Suppress the confirmation pauses. Required for scheduled or unattended runs.

.EXAMPLE
.\Export-M365.ps1

Runs all five reports, prompting for interactive login.

.EXAMPLE
.\Export-M365.ps1 -Report Cap,Iam

Runs only the conditional access and IAM reports.

.EXAMPLE
.\Export-M365.ps1 -Report Org -SkipOnPremSync -Output C:\Audit\2026

Runs the organisation report to a specified folder, skipping the export that needs Global
Administrator.

.EXAMPLE
.\Export-M365.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppCertThumbprint "A1B2C3D4E5F6..."

Authenticates as a registered enterprise application using an installed certificate.

.INPUTS
None.

.OUTPUTS
CSV and HTML evidence files, a SHA-256 manifest, and a run log.

.NOTES
VERSION: 4.0.0

RELEASE APPROVAL: <UNAPPROVED>

2026-08: Restructured from five standalone scripts into a module. See the release notes in
         OagExportM365.psd1 for the full change list.

.LINK
Docs\README.md
Docs\Permissions.md
Docs\Verification.md
#>
[CmdletBinding(DefaultParameterSetName = 'Prompt')]
param (
    [ValidateSet('Cap', 'Iam', 'Org', 'Dfo', 'ThreatHunt', 'All')]
    [string[]]$Report = 'All',

    [string]$Output = $PSScriptRoot,

    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppClientId,

    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppTenantId,

    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [string]$AppCertThumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppSecret,

    [switch]$SkipOnPremSync,

    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------------------------------
# Import the module. RequiredModules in the manifest enforces the PowerShell and dependency
# versions, so a version problem stops here rather than surfacing as a confusing failure part
# way through a run.
# ------------------------------------------------------------------------------------------
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'OagExportM365.psd1'
if (-not (Test-Path $modulePath)) {
    throw "Cannot find OagExportM365.psd1 in $PSScriptRoot. Run this script from the module folder."
}

Import-Module -Name $modulePath -Force -ErrorAction Stop

# Resolve the report selection.
$allReports = @('Cap', 'Iam', 'Org', 'Dfo', 'ThreatHunt')
$selected = if ($Report -contains 'All') { $allReports } else { $allReports | Where-Object { $_ -in $Report } }

if (-not $selected) {
    throw "No valid reports selected."
}

# ------------------------------------------------------------------------------------------
# Per-report Graph scopes and modules. Declared here so the dispatcher can request the union
# of what the selected reports need in a single sign-in, rather than prompting once per report.
# ------------------------------------------------------------------------------------------
$reportRequirements = @{
    'Cap' = @{
        Scopes  = @('Policy.Read.AuthenticationMethod', 'Policy.Read.All', 'AuthenticationContext.Read.All',
                    'Policy.Read.ConditionalAccess', 'RoleManagement.Read.Directory',
                    'Application.Read.All', 'Directory.Read.All')
        Modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DirectoryObjects',
                    'Microsoft.Graph.Identity.SignIns', 'Microsoft.Graph.Identity.DirectoryManagement',
                    'Microsoft.Graph.Applications')
        Service = 'Graph'
    }
    'Iam' = @{
        Scopes  = @('Directory.Read.All', 'DeviceManagementApps.Read.All', 'Device.Read.All',
                    'Application.Read.All', 'AuditLog.Read.All', 'RoleManagement.Read.Directory',
                    'PrivilegedEligibilitySchedule.Read.AzureADGroup',
                    'PrivilegedAssignmentSchedule.Read.AzureADGroup')
        Modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DirectoryObjects',
                    'Microsoft.Graph.Users', 'Microsoft.Graph.Groups',
                    'Microsoft.Graph.Identity.DirectoryManagement', 'Microsoft.Graph.Applications',
                    'Microsoft.Graph.Reports', 'Microsoft.Graph.Identity.Governance')
        Service = 'Graph'
    }
    'Org' = @{
        Scopes  = @('Organization.Read.All', 'Directory.Read.All', 'Domain.Read.All',
                    'OnPremDirectorySynchronization.Read.All')
        Modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DirectoryObjects',
                    'Microsoft.Graph.Identity.DirectoryManagement')
        Service = 'Graph'
    }
    'Dfo' = @{
        Scopes  = @()
        Modules = @('ExchangeOnlineManagement')
        Service = 'ExchangeOnline'
    }
    'ThreatHunt' = @{
        Scopes  = @('ThreatHunting.Read.All')
        Modules = @('Microsoft.Graph.Authentication')
        Service = 'Graph'
    }
}

# ------------------------------------------------------------------------------------------
# Initialise the run: one run ID, one folder, one transcript, code integrity check.
# ------------------------------------------------------------------------------------------
$Context = Initialize-OagM365Run -Output $Output `
                                 -Reports $selected `
                                 -AuthMode $PSCmdlet.ParameterSetName `
                                 -ModuleRoot $PSScriptRoot

try {
    $graphReports = $selected | Where-Object { $reportRequirements[$_].Service -eq 'Graph' }
    $exoReports   = $selected | Where-Object { $reportRequirements[$_].Service -eq 'ExchangeOnline' }

    $graphConnection = $null
    $exoConnection   = $null

    # --- Graph connection, once, with the union of required scopes -------------------------
    if ($graphReports) {
        $scopes  = $graphReports | ForEach-Object { $reportRequirements[$_].Scopes } | Sort-Object -Unique
        $modules = $graphReports | ForEach-Object { $reportRequirements[$_].Modules } | Sort-Object -Unique

        $connectParams = @{
            Scopes          = $scopes
            RequiredModules = $modules
            Context         = $Context
            NoPause         = $NonInteractive
        }
        switch ($PSCmdlet.ParameterSetName) {
            'AppCertThumbprint' {
                $connectParams += @{ AppClientId = $AppClientId; AppTenantId = $AppTenantId; AppCertThumbprint = $AppCertThumbprint }
            }
            'AppSecret' {
                $connectParams += @{ AppClientId = $AppClientId; AppTenantId = $AppTenantId; AppSecret = $AppSecret }
            }
        }

        $graphConnection = Connect-OagM365Graph @connectParams

        if (-not $graphConnection.Connected) {
            throw "Could not connect to Microsoft Graph. No Graph reports can run."
        }

        $Context.TenantId = $graphConnection.TenantId
        $Context.Account  = $graphConnection.Account

        # Record what the run was actually permitted to see. This is the evidence that
        # distinguishes "the tenant has no such policy" from "we could not read it".
        foreach ($scope in $scopes) {
            $Context.ScopeLog.Add([PSCustomObject]@{
                Service = 'Graph'
                Scope   = $scope
                Granted = ($scope -in $graphConnection.ScopesGranted)
            })
        }
    }

    # --- Exchange Online connection --------------------------------------------------------
    if ($exoReports) {
        $exoConnection = Connect-OagM365ExchangeOnline -Context $Context -NoPause:$NonInteractive
        if (-not $exoConnection.Connected) {
            Write-OagM365Log "Could not connect to Exchange Online. The Dfo report will be skipped." -Level Error -Context $Context
        }
    }

    # --- Run the selected reports ----------------------------------------------------------
    foreach ($reportName in $selected) {
        Write-OagM365Log "" -Context $Context
        Write-OagM365Log ("-" * 70) -Level Detail -Context $Context
        Write-OagM365Log "REPORT: $reportName" -Context $Context
        Write-OagM365Log ("-" * 70) -Level Detail -Context $Context

        switch ($reportName) {
            'Cap' {
                Export-OagM365CapReport -Context $Context -Connection $graphConnection | Out-Null
            }
            'Iam' {
                Export-OagM365IamReport -Context $Context -Connection $graphConnection | Out-Null
            }
            'Org' {
                Export-OagM365OrgReport -Context $Context -Connection $graphConnection -SkipOnPremSync:$SkipOnPremSync | Out-Null
            }
            'Dfo' {
                if ($exoConnection -and $exoConnection.Connected) {
                    Export-OagM365DfoReport -Context $Context -Connection $exoConnection | Out-Null
                } else {
                    Write-OagM365Log "Skipped - no Exchange Online connection." -Level Warning -Indent 1 -Context $Context
                }
            }
            'ThreatHunt' {
                Export-OagM365ThreatHuntReport -Context $Context -Connection $graphConnection | Out-Null
            }
        }
    }

} catch {
    Write-OagM365Log "RUN FAILED: $($_.Exception.Message)" -Level Error -Context $Context
    Write-OagM365Log $_.ScriptStackTrace -Level Detail -Context $Context

} finally {
    # --------------------------------------------------------------------------------------
    # Close out the run. This runs even on failure so a partial evidence set is still hashed,
    # logged and marked incomplete rather than left unexplained.
    # --------------------------------------------------------------------------------------
    $Context.CurrentReport = 'Shutdown'
    $Context.EndTime = Get-Date

    Write-OagM365Log "" -Context $Context
    Write-OagM365Log ("=" * 70) -Level Detail -Context $Context

    Disconnect-OagM365Service -Context $Context

    # Granted scopes, written before the manifest so they are covered by it.
    if ($Context.ScopeLog.Count -gt 0) {
        $Context.ScopeLog | Export-Csv -Path (Join-Path $Context.LogPath 'scopes-granted.csv') -NoTypeInformation -Encoding utf8
    }

    $completeness = Test-OagM365Completeness -Context $Context

    # Event log, written before the manifest for the same reason.
    if ($Context.Events.Count -gt 0) {
        $Context.Events | Export-Csv -Path (Join-Path $Context.LogPath 'events.csv') -NoTypeInformation -Encoding utf8
    }

    $summary = [PSCustomObject]@{
        RunId              = $Context.RunId
        Started            = $Context.StartTime.ToString('o')
        Completed          = $Context.EndTime.ToString('o')
        DurationMinutes    = [math]::Round(($Context.EndTime - $Context.StartTime).TotalMinutes, 2)
        TenantId           = $Context.TenantId
        Account            = $Context.Account
        AuthMode           = $Context.AuthMode
        ReportsRequested   = $selected -join ', '
        ReportsSucceeded   = (($Context.ReportStatus | Where-Object Succeeded) | ForEach-Object { $_.Report }) -join ', '
        CompletenessStatus = $completeness.Status
        FilesExpected      = $completeness.FilesExpected
        FilesOk            = $completeness.FilesOk
        FilesEmpty         = $completeness.FilesEmpty
        FilesMissing       = $completeness.FilesMissing
        Errors             = $completeness.ErrorCount
        Warnings           = $completeness.WarningCount
    }
    $summary | ConvertTo-Json -Depth 4 | Out-File -FilePath (Join-Path $Context.LogPath 'run-summary.json') -Encoding utf8

    Write-OagM365Log "" -Context $Context
    Write-OagM365Log "RUN SUMMARY" -Context $Context
    Write-OagM365Log "Run ID       : $($Context.RunId)" -Level Detail -Indent 1 -Context $Context
    Write-OagM365Log "Duration     : $($summary.DurationMinutes) minute(s)" -Level Detail -Indent 1 -Context $Context
    Write-OagM365Log "Reports      : $($summary.ReportsSucceeded)" -Level Detail -Indent 1 -Context $Context
    Write-OagM365Log "Files        : $($completeness.FilesOk) OK, $($completeness.FilesEmpty) empty, $($completeness.FilesMissing) missing" -Level Detail -Indent 1 -Context $Context
    Write-OagM365Log "Status       : $($completeness.Status)" `
        -Level $(if ($completeness.Status -eq 'Complete') { 'Success' } else { 'Warning' }) -Indent 1 -Context $Context
    Write-OagM365Log "Evidence     : $($Context.RootPath)" -Level Detail -Indent 1 -Context $Context

    # --------------------------------------------------------------------------------------
    # ORDER MATTERS HERE.
    #
    # The transcript must be closed BEFORE the manifest is generated. A transcript that is
    # still recording grows as the manifest is written, so its hash changes the instant after
    # it is taken and the file then fails verification on every single run. That would train
    # reviewers to ignore a failed verification, which defeats the point of having one.
    #
    # Closing the transcript first means the manifest covers a complete, stable file. The
    # trade-off is that the manifest generation messages below appear on the console but not
    # in the transcript, which is the right way round: the evidence of what was hashed is the
    # manifest itself.
    # --------------------------------------------------------------------------------------
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

    New-OagM365Manifest -Context $Context | Out-Null

    Write-Host ""
    Write-Host "Evidence written to: $($Context.RootPath)" -ForegroundColor Cyan
    Write-Host "Verify with        : .\Tools\Verify-Export.ps1 -RunFolder '$($Context.RootPath)'" -ForegroundColor Gray
}
