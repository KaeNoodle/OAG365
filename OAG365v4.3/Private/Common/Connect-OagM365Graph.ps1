function Connect-OagM365Graph {
    <#
    .SYNOPSIS
    Connects PowerShell to Microsoft Graph and verifies the granted scopes.

    .DESCRIPTION
    Imports the required Microsoft.Graph modules and connects using one of three
    authentication modes: interactive prompt, application secret, or application certificate.

    This function was previously duplicated byte-for-byte in CapReport, IamReport, OrgReport
    and ThreatHunting. It is now defined once.

    Two behavioural changes from the original:

      1. Scopes are passed in as a parameter rather than read from a script-scope variable
         ($script:exportM365MsGraphScopes). In the original each script declared its scopes in
         a variable at the top of the file and separately listed them in the comment-based help,
         and the two drifted apart. Each report now declares its own scopes next to the code
         that needs them and passes them in, so there is one source of truth.

      2. The granted-versus-requested comparison is returned to the caller as well as printed,
         so it can be written to the run log as evidence of what the export was actually
         permitted to see. A missing scope is the most common cause of a silently incomplete
         export, so this needs to be recorded, not just displayed.

    .PARAMETER Scopes
    The Microsoft Graph scopes this report requires.

    .PARAMETER RequiredModules
    Microsoft.Graph submodules to import for this report.

    .PARAMETER AppClientId
    Registered Entra application's client / application ID.

    .PARAMETER AppTenantId
    M365 tenant ID for the registered Entra application.

    .PARAMETER AppCertThumbprint
    Thumbprint of the certificate used to authenticate to the registered Entra application.
    The certificate must be installed in the machine certificate store.

    .PARAMETER AppSecret
    The registered Entra application's secret (plain text).

    .PARAMETER Context
    Run context object, used for logging.

    .PARAMETER NoPause
    Suppress the interactive pause after connecting. Set by the dispatcher when running
    multiple reports so the operator is not prompted between each one.

    .EXAMPLE
    Connect-OagM365Graph -Scopes @('Policy.Read.All') -RequiredModules @('Microsoft.Graph.Authentication') -Context $Context

    .NOTES
        NAME: Connect-OagM365Graph
        VERSION: 2.0

        GRAPH CMDLETS:
           Connect-MgGraph
           Get-MgContext

        CHANGELOG:
          4.0.0: Scopes and modules are now parameters. Returns a result object rather than
                 a bare boolean so the caller can log granted scopes.
          2025-11-18: Import modules only when required.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Prompt')]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,

        [Parameter(Mandatory = $true)]
        [string[]]$RequiredModules,

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

        [object]$Context,

        [switch]$NoPause
    )

    $result = [PSCustomObject]@{
        Connected      = $false
        TenantId       = $null
        Account        = $null
        AuthType       = $null
        ScopesRequired = $Scopes
        ScopesGranted  = @()
        ScopesMissing  = @()
    }

    try {
        Write-OagM365Log "Microsoft Graph connection required. Please log in if prompted." -Context $Context

        $missingModules = $RequiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
        if ($missingModules) {
            throw "Missing required PowerShell module(s): $($missingModules -join ', '). Install with: Install-Module <name> -Scope CurrentUser"
        }

        $loaded = (Get-Module).Name
        $toImport = $RequiredModules | Where-Object { $_ -notin $loaded }
        if ($toImport) {
            Write-OagM365Log "Importing Graph modules: $($toImport -join ', ')" -Level Detail -Indent 1 -Context $Context
            # Imported at global scope so the cmdlets remain available to every function in
            # the module for the whole run, not just inside this function.
            $toImport | Import-Module -Global -ErrorAction Stop
        }

        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            switch ($PSCmdlet.ParameterSetName) {
                'AppCertThumbprint' {
                    Write-OagM365Log "Connecting using application certificate $AppCertThumbprint" -Indent 1 -Context $Context
                    Connect-MgGraph -ContextScope Process -ClientId $AppClientId -TenantId $AppTenantId `
                                    -CertificateThumbprint $AppCertThumbprint -NoWelcome -ErrorAction Stop
                }
                'AppSecret' {
                    Write-OagM365Log "Connecting using application secret" -Indent 1 -Context $Context
                    $secureString = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
                    $clientSecretCreds = New-Object -TypeName System.Management.Automation.PSCredential `
                                                    -ArgumentList $AppClientId, $secureString
                    Connect-MgGraph -ContextScope Process -TenantId $AppTenantId `
                                    -ClientSecretCredential $clientSecretCreds -NoWelcome -ErrorAction Stop
                }
                'Prompt' {
                    Write-OagM365Log "Connecting interactively, requesting scopes: $($Scopes -join ', ')" -Indent 1 -Context $Context
                    Connect-MgGraph -ContextScope Process -Scopes $Scopes -NoWelcome -ErrorAction Stop
                }
            }
        }

        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $mgContext) {
            throw "Authentication failed. Could not establish a Microsoft Graph context."
        }

        $result.Connected     = $true
        $result.TenantId      = $mgContext.TenantId
        $result.Account       = if ($mgContext.Account) { $mgContext.Account } else { $mgContext.AppName }
        $result.AuthType      = $mgContext.AuthType
        $result.ScopesGranted = @($mgContext.Scopes)
        $result.ScopesMissing = @($Scopes | Where-Object { $_ -notin $mgContext.Scopes })

        switch -Regex ($PSCmdlet.ParameterSetName) {
            'AppCertThumbprint|AppSecret' {
                Write-OagM365Log "Connected as Enterprise Application" -Level Success -Indent 1 -Context $Context
                Write-OagM365Log "Tenant ID : $($mgContext.TenantId)" -Level Detail -Indent 2 -Context $Context
                Write-OagM365Log "Client ID : $($mgContext.ClientId)" -Level Detail -Indent 2 -Context $Context
                Write-OagM365Log "App Name  : $($mgContext.AppName)" -Level Detail -Indent 2 -Context $Context
                Write-OagM365Log "Auth Type : $($mgContext.AuthType)" -Level Detail -Indent 2 -Context $Context
            }
            'Prompt' {
                Write-OagM365Log "Connected as $($mgContext.Account) (tenant $($mgContext.TenantId))" -Level Success -Indent 1 -Context $Context
            }
        }

        # Report on scopes. A missing scope means an incomplete export, so it is logged as a
        # warning rather than printed and forgotten.
        if ($result.ScopesMissing.Count -gt 0) {
            Write-OagM365Log "MISSING $($result.ScopesMissing.Count) required scope(s) - this export will be incomplete:" -Level Warning -Indent 1 -Context $Context
            $result.ScopesMissing | ForEach-Object {
                Write-OagM365Log $_ -Level Warning -Indent 2 -Context $Context
            }
        } else {
            Write-OagM365Log "All $($Scopes.Count) required scope(s) granted" -Level Success -Indent 1 -Context $Context
        }

        if (-not $NoPause) { Pause }

        return $result

    } catch {
        $exception = Format-OagM365Exception -message "Failed connecting to Microsoft Graph PowerShell" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
        return $result
    }
}
