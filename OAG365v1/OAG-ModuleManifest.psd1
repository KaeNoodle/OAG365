@{
    RootModule        = 'OAG-FileInfo.psm1'
    ModuleVersion     = '4.0.0'
    GUID              = 'e7c4a1b9-3d5f-4a82-9c16-7b0e2f8d4a63'
    Author            = 'Office of the Auditor General (Western Australia)'
    CompanyName       = 'Office of the Auditor General (Western Australia)'
    Copyright         = '(c) Office of the Auditor General (Western Australia)'
    Description       = 'Exports Microsoft 365 and Entra ID configuration as audit evidence, with integrity verification of both the module code and the exported evidence.'

    # ----------------------------------------------------------------------------------
    # REQUIREMENTS
    #
    # Enforced by PowerShell at import. The original scripts printed a warning when a
    # module was missing or the wrong version and then carried on regardless, so a run
    # could complete against the wrong module and nobody would notice until the evidence
    # was reviewed. Declaring them here means the import fails instead.
    # ----------------------------------------------------------------------------------

    PowerShellVersion = '7.2'

    RequiredModules   = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' }
        @{ ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.9.0'}
    )

    FunctionsToExport = @(
        'capReportWrite'
        'iamReportWrite'
        'orgReportWrite'
        'dfoReportWrite'
        'threatHuntReportWrite'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # ----------------------------------------------------------------------------------
    # FILELIST feeds New-FileCatalog, which produces OAG-FileCatalog.cat.
    # See Tools\New-OagCatalog.ps1 and Documentation\Verification.md.
    # ----------------------------------------------------------------------------------
    FileList = @(
        'OAG-MainRunFile.ps1'
        'OAG-ModuleManifest.psd1'
        'OAG-FileInfo.psm1'
        'ReportWriter'
        'Functions'
    )

    PrivateData = @{
        PSData = @{
            Tags         = @('M365', 'Entra', 'Audit', 'Evidence', 'Essential8')
            ReleaseNotes = 'See Documentation\CHANGELOG.md'
        }
    }
}
