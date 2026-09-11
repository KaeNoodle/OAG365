function msGraphConnect {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Connects to Microsoft Graph and records which of the requested scopes were actually
    granted.

    This function was duplicated byte for byte in four of the five original scripts.
    It is now defined once.

    Scopes are passed in rather than read from a script variable. In the original, each
    script declared its scopes in one place and documented them in another, and the two
    had drifted apart. Each report now declares its own scopes next to the code that
    needs them.

    LOGIC
    Checks the required Graph modules are installed, and stops if any are missing.
    Imports them at global scope so the cmdlets stay available for the whole run.
    Connects using whichever of the three authentication modes was supplied.
    For interactive sign-in, falls back to device code if the Web Account Manager prompt
      does not complete. That prompt is on a short timer and opens its own window, which
      on an embedded terminal frequently appears behind everything else, so the common
      failure is the operator never seeing it rather than the credentials being wrong.
    Compares requested scopes against granted scopes and warns on any shortfall, since
      a missing scope produces a silently incomplete export rather than an error.
    Returns a result object holding tenant, account and the scope comparison.

    PARAMETERS
    -scopes (required) the Graph scopes this report needs
    -modules (required) Microsoft.Graph submodules to import
    -appClientId (optional) application ID, for enterprise application authentication
    -appTenantId (optional) tenant ID, for enterprise application authentication
    -appCertThumbprint (optional) certificate thumbprint, must be in the machine store
    -appSecret (optional) application secret, plain text
    -deviceCode (optional) sign in with a device code instead of the interactive window
    -noPause (optional) skip the confirmation pause, for unattended runs

    RUNNING CONTEXT
    Called by  : runMe.ps1
    Calls      : logWrite, exceptionFormat
    Returns    : object with connected, tenantId, account, scopesGranted, scopesMissing

    CMLETS/PERMISSIONS/SCOPES
    Connect-MgGraph
    Get-MgContext

    --------------------------------------------------------------------------------#>

    [CmdletBinding(DefaultParameterSetName = 'Prompt')]
    Param(
        [Parameter(Mandatory = $true)][string[]]$scopes,
        [Parameter(Mandatory = $true)][string[]]$modules,
        [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appClientId,
        [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appTenantId,
        [Parameter(Mandatory = $true, ParameterSetName = 'AppCertThumbprint')][string]$appCertThumbprint,
        [Parameter(Mandatory = $true, ParameterSetName = 'AppSecret')][string]$appSecret,
        [switch]$deviceCode,
        [switch]$noPause
    )

    $result = [PSCustomObject]@{
        connected      = $false
        tenantId       = $null
        account        = $null
        authType       = $null
        scopesRequired = $scopes
        scopesGranted  = @()
        scopesMissing  = @()
    }

    try {
        logWrite "Microsoft Graph connection required. Please log in if prompted."

        $missing = $modules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
        if ($missing) {
            throw "Missing PowerShell module(s): $($missing -join ', '). Install with: Install-Module <name> -Scope CurrentUser"
        }

        $loaded   = (Get-Module).Name
        $toImport = $modules | Where-Object { $_ -notin $loaded }
        if ($toImport) {
            logWrite "Importing Graph modules: $($toImport -join ', ')" -level Detail -indent 1
            $toImport | Import-Module -Global -ErrorAction Stop
        }

        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            switch ($PSCmdlet.ParameterSetName) {
                'AppCertThumbprint' {
                    logWrite "Connecting with application certificate $appCertThumbprint" -indent 1
                    Connect-MgGraph -ContextScope Process -ClientId $appClientId -TenantId $appTenantId `
                                    -CertificateThumbprint $appCertThumbprint -NoWelcome -ErrorAction Stop
                }
                'AppSecret' {
                    logWrite "Connecting with application secret" -indent 1
                    $secure = ConvertTo-SecureString -String $appSecret -AsPlainText -Force
                    $creds  = New-Object System.Management.Automation.PSCredential ($appClientId, $secure)
                    Connect-MgGraph -ContextScope Process -TenantId $appTenantId `
                                    -ClientSecretCredential $creds -NoWelcome -ErrorAction Stop
                }
                'Prompt' {
                    if ($deviceCode) {
                        logWrite "Connecting by device code, requesting $($scopes.Count) scope(s)" -indent 1
                        Connect-MgGraph -ContextScope Process -Scopes $scopes -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
                    } else {
                        logWrite "Connecting interactively, requesting $($scopes.Count) scope(s)" -indent 1
                        logWrite "The sign-in window can open behind this one - check the taskbar if nothing appears." -level Detail -indent 1

                        try {
                            Connect-MgGraph -ContextScope Process -Scopes $scopes -NoWelcome -ErrorAction Stop
                        } catch {
                            # Web Account Manager sign-in runs on a short timer and its window
                            # often opens behind the terminal, so a failure here is usually the
                            # operator not reaching it in time rather than a real auth problem.
                            # Device code has no window to lose and a far longer timer, so retry
                            # that way once before giving up.
                            logWrite "Interactive sign-in did not complete: $($_.Exception.Message)" -level Warning -indent 1
                            logWrite "Retrying with device code, which allows more time to sign in." -level Warning -indent 1
                            Connect-MgGraph -ContextScope Process -Scopes $scopes -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
                        }
                    }
                }
            }
        }

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $context) { throw "Authentication failed. No Microsoft Graph context was established." }

        $result.connected     = $true
        $result.tenantId      = $context.TenantId
        $result.account       = if ($context.Account) { $context.Account } else { $context.AppName }
        $result.authType      = $context.AuthType
        $result.scopesGranted = @($context.Scopes)
        $result.scopesMissing = @($scopes | Where-Object { $_ -notin $context.Scopes })

        logWrite "Connected as $($result.account) (tenant $($result.tenantId))" -level Success -indent 1

        # A missing scope means an incomplete export, so this is a warning rather than
        # something printed once and forgotten.
        if ($result.scopesMissing.Count -gt 0) {
            logWrite "MISSING $($result.scopesMissing.Count) scope(s). Affected exports will be incomplete:" -level Warning -indent 1
            $result.scopesMissing | ForEach-Object { logWrite $_ -level Warning -indent 2 }
        } else {
            logWrite "All $($scopes.Count) requested scope(s) granted" -level Success -indent 1
        }

        if ($null -ne $script:run) {
            $script:run.tenantId = $result.tenantId
            $script:run.account  = $result.account
            foreach ($scope in $scopes) {
                $script:run.scopes += [PSCustomObject]@{
                    service = 'Graph'
                    scope   = $scope
                    granted = ($scope -in $result.scopesGranted)
                }
            }
        }

        if (-not $noPause) { Pause }
        return $result

    } catch {
        logWrite (exceptionFormat -message "Failed connecting to Microsoft Graph" -exception $_) -level Error
        return $result
    }
}
