function Initialize-OagM365ReportPath {
    <#
    .SYNOPSIS
    Creates a report's output subfolder and sets the module-scope export target variables.

    .DESCRIPTION
    Each report previously defined its own $script:exportTargetFolder and $script:exportTarget
    inside its begin{} block. Those definitions are consolidated here, keyed by report.

    Why the ported functions still reference $script:exportTarget
    ------------------------------------------------------------
    Inside a module, $script: resolves to module scope, which every function in the module
    shares. Setting $script:exportTarget here therefore makes the paths visible to all the
    ported export functions without editing the several hundred references to them inside those
    functions. That keeps the port close to a like-for-like move, which matters because the
    original code is signed, reviewed and known to work: a smaller diff is a cheaper review and
    a smaller chance of introducing a regression while restructuring.

    Expected outputs are registered on the context at the same time. Test-OagM365Completeness
    later compares this list against what was actually written, which is how a silently
    truncated or missing export is caught.

    .PARAMETER Context
    The run context returned by Initialize-OagM365Run.

    .PARAMETER Report
    Which report's paths to configure.

    .EXAMPLE
    Initialize-OagM365ReportPath -Context $Context -Report Cap

    .NOTES
        NAME: Initialize-OagM365ReportPath
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Cap', 'Iam', 'Org', 'Dfo', 'ThreatHunt')]
        [string]$Report
    )

    $subFolderName = switch ($Report) {
        'Cap'        { 'ConditionalAccess' }
        'Iam'        { 'IAM' }
        'Org'        { 'Organisation' }
        'Dfo'        { 'ThreatProtection' }
        'ThreatHunt' { 'Defender' }
    }

    $folder = Join-Path -Path $Context.RootPath -ChildPath $subFolderName
    New-Item -ItemType Directory -Force -Path $folder -ErrorAction Stop | Out-Null

    # Module-scope folder object. Property names match the originals so ported code resolves.
    $script:exportTargetFolder = [PSCustomObject]@{
        conditionalAccess = $folder
        iam               = $folder
        organisation      = $folder
        threatProtection  = $folder
        defender          = $folder
    }

    # Per-report file map, merged from the five original begin{} blocks.
    $fileMap = switch ($Report) {
        'Cap' {
            [ordered]@{
                authMethodPolicy   = 'authMethodPolicies.csv'
                authStrengthPolicy = 'authStrengthPolicies.csv'
                capSummary         = 'capSummary.csv'
            }
        }
        'Iam' {
            [ordered]@{
                pimMgmtPolicyRules = 'pimMgmtPolicyRules.csv'
                builtInRoleMembers = 'builtInRoleMembers.csv'
                roleAdminUnits     = 'adminUnits.csv'
                userMfa            = 'userMFA.csv'
                users              = 'users.csv'
                servicePrincipals  = 'servicePrincipals.csv'
                entraDevices       = 'entraDevices.csv'
            }
        }
        'Org' {
            [ordered]@{
                organisations          = 'organisations.csv'
                onPremisesSync         = 'onPremisesSync.csv'
                domains                = 'domains.csv'
                domainDkim             = 'domainDkim.csv'
                domainSpf              = 'domainSpf.csv'
                domainDmarc            = 'domainDmarc.csv'
                domainFederationConfig = 'domainFederationConfig.csv'
                authMethodPolicy       = 'authMethodPolicies.csv'
                authStrengthPolicy     = 'authStrengthPolicies.csv'
                directorySettings      = 'directorySettings.csv'
            }
        }
        'Dfo' {
            [ordered]@{
                dtppSafeLinksPolicy      = 'dtppSafeLinksPolicy.csv'
                dtppSafeLinksRule        = 'dtppSafeLinksRule.csv'
                dtppSafeAttachmentPolicy = 'dtppSafeAttachmentPolicy.csv'
                dtppSafeAttachmentRule   = 'dtppSafeAttachmentRule.csv'
                eopMalwarePolicy         = 'eopMalwarePolicy.csv'
                eopMalwareRule           = 'eopMalwareRule.csv'
                eopPhishPolicy           = 'eopPhishPolicy.csv'
                eopPhishRule             = 'eopPhishRule.csv'
                eopSpamPolicy            = 'eopSpamPolicy.csv'
                dfoQuarantinePolicy      = 'dfoQuarantinePolicy.csv'
            }
        }
        'ThreatHunt' {
            [ordered]@{
                defenderHealthReport = 'defenderHealthReport.csv'
                defenderTvmReport    = 'defenderTvmReport.csv'
                softwareInventory    = 'softwareInventory.csv'
            }
        }
    }

    $target = [ordered]@{}
    foreach ($key in $fileMap.Keys) {
        $target[$key] = Join-Path -Path $folder -ChildPath $fileMap[$key]
    }
    $script:exportTarget = [PSCustomObject]$target

    $Context.CurrentReport = $Report
    Write-OagM365Log "Output folder: $folder" -Level Detail -Indent 1 -Context $Context

    return $folder
}
