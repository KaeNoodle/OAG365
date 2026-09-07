@{
    RootModule        = 'OagExportM365.psm1'
    ModuleVersion     = '4.0.0'
    GUID              = 'e7c4a1b9-3d5f-4a82-9c16-7b0e2f8d4a63'
    Author            = 'Office of the Auditor General (Western Australia)'
    CompanyName       = 'Office of the Auditor General (Western Australia)'
    Copyright         = '(c) Office of the Auditor General (Western Australia)'
    Description       = 'Exports Microsoft 365 and Entra ID configuration as audit evidence, with integrity verification of both the module code and the exported evidence.'

    # --------------------------------------------------------------------------------------
    # Requirements
    #
    # These are enforced by PowerShell at import time. The original scripts printed a warning
    # when a module was missing or a version was wrong and then carried on regardless, which
    # meant a run could complete against the wrong module version and nobody would notice until
    # the evidence was reviewed. Declaring them here means the import fails instead.
    # --------------------------------------------------------------------------------------

    PowerShellVersion = '7.2'

    RequiredModules   = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' }

        # ExchangeOnlineManagement is pinned to the 3.9.x line. Version 3.10.1 shipped a
        # regression affecting the policy cmdlets this module depends on; 3.9.0 is the last
        # build confirmed working for this export.
        @{ ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.9.0'; MaximumVersion = '3.9.9' }
    )

    # Loaded on demand by the reports that need them rather than at import, so a run that only
    # needs one report does not pay to load all of Microsoft.Graph.
    RequiredAssemblies = @()

    FunctionsToExport = @(
        'Export-OagM365CapReport'
        'Export-OagM365IamReport'
        'Export-OagM365OrgReport'
        'Export-OagM365DfoReport'
        'Export-OagM365ThreatHuntReport'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # --------------------------------------------------------------------------------------
    # FileList feeds New-FileCatalog. Every file listed here is hashed into OagExportM365.cat,
    # which is then signed. See Tools\New-OagCatalog.ps1 and Docs\Verification.md.
    # --------------------------------------------------------------------------------------
    FileList = @(
        'OagExportM365.psd1'
        'OagExportM365.psm1'
        'Export-M365.ps1'
        'Public\Export-OagM365CapReport.ps1'
        'Public\Export-OagM365IamReport.ps1'
        'Public\Export-OagM365OrgReport.ps1'
        'Public\Export-OagM365DfoReport.ps1'
        'Public\Export-OagM365ThreatHuntReport.ps1'
        'Private\Common'
        'Private\Cap'
        'Private\Iam'
        'Private\Org'
        'Private\Dfo'
        'Private\ThreatHunt'
    )

    PrivateData = @{
        PSData = @{
            Tags         = @('M365', 'Entra', 'Audit', 'Evidence', 'Essential8')
            ReleaseNotes = @'
4.0.0
  - Restructured five standalone scripts into a single module with one entry point
    (Export-M365.ps1 -Report). Reports remain independently runnable.
  - Removed ~1,200 lines of duplicated code. Connect, disconnect and exception handling
    were previously copy-pasted across files; they are now defined once.
  - One run ID and one output folder per execution, replacing five independently
    timestamped folders that could not be correlated.
  - Added SHA-256 evidence manifest and code integrity self-check.
  - Added completeness checking to detect silently empty or missing exports.
  - Fixed Exchange Online temporary cmdlet module scope, which was causing the DFO
    report to produce no output.
  - Fixed the Connect-ExchangeOnline return value test in the DFO report.
  - Fixed silent error suppression in the threat hunting queries.
  - Graph scopes are now declared per report as parameters rather than in a script
    variable that had drifted from the documented list.
'@
        }
    }
}
