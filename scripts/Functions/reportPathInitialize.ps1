function reportPathInitialize {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Creates a report's output subfolder and sets the export target paths that the
    export functions write to.

    Each report previously defined its own folder and file paths in its own begin block.
    Those five definitions are consolidated here.

    Why the export functions still read $script:exportTarget: inside a module, $script:
    is module scope, shared by every function in the module. Setting it here makes the
    paths visible to all the ported export functions without editing the several hundred
    references inside them. The original code is signed, reviewed and known to work, so a
    smaller diff means a cheaper review and less chance of breaking something.

    LOGIC
    Maps the report name to a folder name and creates the folder.
    Builds the per-report file map, merged from the five original scripts.
    Sets $script:exportTargetFolder and $script:exportTarget at module scope.
    Records the report name on the run so log entries are attributed to it.

    PARAMETERS
    -report (required, one of: CAP, IAM, ORG, DFO, TH)

    RUNNING CONTEXT
    Called by  : each of the five report writers
    Calls      : logWrite
    Sets       : $script:exportTarget, $script:exportTargetFolder
    Returns    : the folder path

    CMLETS/PERMISSIONS/SCOPES
    None. Local filesystem only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][ValidateSet('CAP', 'IAM', 'ORG', 'DFO', 'TH')][string]$report
    )

    $folderName = switch ($report) {
        'CAP' { 'ConditionalAccess' }
        'IAM' { 'IAM' }
        'ORG' { 'Organisation' }
        'DFO' { 'ThreatProtection' }
        'TH'  { 'Defender' }
    }

    $folder = Join-Path $script:run.rootPath $folderName
    New-Item -ItemType Directory -Force -Path $folder -ErrorAction Stop | Out-Null

    # Property names match the originals so the ported code resolves unchanged.
    $script:exportTargetFolder = [PSCustomObject]@{
        conditionalAccess = $folder
        iam               = $folder
        organisation      = $folder
        threatProtection  = $folder
        defender          = $folder
    }

    $files = switch ($report) {
        'CAP' { [ordered]@{
            authMethodPolicy = 'authMethodPolicies.csv'; authStrengthPolicy = 'authStrengthPolicies.csv'
            capSummary = 'capSummary.csv' } }
        'IAM' { [ordered]@{
            users = 'users.csv'; entraDevices = 'entraDevices.csv'; servicePrincipals = 'servicePrincipals.csv'
            userMfa = 'userMFA.csv'; roleAdminUnits = 'adminUnits.csv'
            builtInRoleMembers = 'builtInRoleMembers.csv'; pimMgmtPolicyRules = 'pimMgmtPolicyRules.csv' } }
        'ORG' { [ordered]@{
            organisations = 'organisations.csv'; onPremisesSync = 'onPremisesSync.csv'; domains = 'domains.csv'
            domainDkim = 'domainDkim.csv'; domainSpf = 'domainSpf.csv'; domainDmarc = 'domainDmarc.csv'
            domainFederationConfig = 'domainFederationConfig.csv'; authMethodPolicy = 'authMethodPolicies.csv'
            authStrengthPolicy = 'authStrengthPolicies.csv'; directorySettings = 'directorySettings.csv' } }
        'DFO' { [ordered]@{
            eopMalwarePolicy = 'eopMalwarePolicy.csv'; eopMalwareRule = 'eopMalwareRule.csv'
            eopPhishPolicy = 'eopPhishPolicy.csv'; eopPhishRule = 'eopPhishRule.csv'
            eopSpamPolicy = 'eopSpamPolicy.csv'
            dtppSafeLinksPolicy = 'dtppSafeLinksPolicy.csv'; dtppSafeLinksRule = 'dtppSafeLinksRule.csv'
            dtppSafeAttachmentPolicy = 'dtppSafeAttachmentPolicy.csv'
            dtppSafeAttachmentRule = 'dtppSafeAttachmentRule.csv'
            dfoQuarantinePolicy = 'dfoQuarantinePolicy.csv' } }
        'TH'  { [ordered]@{
            defenderHealthReport = 'defenderHealthReport.csv'; defenderTvmReport = 'defenderTvmReport.csv'
            softwareInventory = 'softwareInventory.csv' } }
    }

    $target = [ordered]@{}
    foreach ($key in $files.Keys) { $target[$key] = Join-Path $folder $files[$key] }
    $script:exportTarget = [PSCustomObject]$target

    $script:run.currentReport = $report
    logWrite "Output folder: $folder" -level Detail -indent 1

    return $folder
}
