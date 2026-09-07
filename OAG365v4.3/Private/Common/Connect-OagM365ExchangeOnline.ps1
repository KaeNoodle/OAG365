function Connect-OagM365ExchangeOnline {
    <#
    .SYNOPSIS
    Connects PowerShell to Exchange Online Management.

    .DESCRIPTION
    Imports the ExchangeOnlineManagement module and connects to Exchange Online.

    FIX: temporary cmdlet module scope
    ----------------------------------
    Connect-ExchangeOnline does not ship static cmdlets. It generates a temporary module named
    tmpEXO_<random> at connection time and imports it into the scope of whatever called it.

    In the original DfoReport script the connect call sat inside a function, so the generated
    module was imported into that function's local scope and was discarded the moment the
    function returned. Every subsequent Get-MalwareFilterPolicy, Get-AntiPhishPolicy and so on
    then failed with "term is not recognised", which is why the DFO report was producing no
    output.

    The fix is the re-import below: after connecting, the generated tmpEXO_* module is
    explicitly re-imported with -Global so its cmdlets survive the return and remain callable
    from the export functions. The same treatment is applied to ExchangeOnlineManagement itself.

    FIX: connection test
    --------------------
    The original called this as `if (Connect-Oag365Dfo-ToExchangeOnline -eq $true)`. PowerShell
    parses that as passing two positional arguments to the function, not as comparing its return
    value, so the branch evaluated on the truthiness of the returned object rather than on
    success. The caller now assigns the result and tests it properly.

    .PARAMETER RequiredModules
    Exchange Online modules to import. Defaults to ExchangeOnlineManagement.

    .PARAMETER UserPrincipalName
    Optional UPN to pass to Connect-ExchangeOnline to pre-fill the login prompt.

    .PARAMETER Context
    Run context object, used for logging.

    .PARAMETER NoPause
    Suppress the interactive pause after connecting.

    .EXAMPLE
    $exo = Connect-OagM365ExchangeOnline -Context $Context
    if ($exo.Connected) { ... }

    .NOTES
        NAME: Connect-OagM365ExchangeOnline
        VERSION: 2.0

        EXO CMDLETS:
           Connect-ExchangeOnline
           Get-ConnectionInformation

        CHANGELOG:
          4.0.0: Global re-import of the tmpEXO_* module. Returns a result object.
          2025-11-18: Import modules only when required.
    #>
    [CmdletBinding()]
    param (
        [string[]]$RequiredModules = @('ExchangeOnlineManagement'),
        [string]$UserPrincipalName,
        [object]$Context,
        [switch]$NoPause
    )

    $result = [PSCustomObject]@{
        Connected         = $false
        UserPrincipalName = $null
        Organization      = $null
        ModuleVersion     = $null
    }

    try {
        Write-OagM365Log "Exchange Online connection required. Please log in if prompted." -Context $Context

        $missingModules = $RequiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
        if ($missingModules) {
            throw "Missing required PowerShell module(s): $($missingModules -join ', '). Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
        }

        # Import at global scope, not function scope.
        $RequiredModules | Import-Module -Global -ErrorAction Stop

        $exoModule = Get-Module -Name ExchangeOnlineManagement | Select-Object -First 1
        if ($exoModule) {
            $result.ModuleVersion = $exoModule.Version.ToString()
            Write-OagM365Log "ExchangeOnlineManagement version $($exoModule.Version)" -Level Detail -Indent 1 -Context $Context

            # Version 3.10.1 shipped a regression affecting these cmdlets. 3.9.x is the last
            # known-good build for this export, and the manifest pins it, but a manually
            # imported module can still bypass that, so warn here as well.
            if ($exoModule.Version -ge [version]'3.10.0') {
                Write-OagM365Log "ExchangeOnlineManagement $($exoModule.Version) is newer than the tested 3.9.x. If cmdlets fail, pin to 3.9.0." -Level Warning -Indent 1 -Context $Context
            }
        }

        Write-OagM365Log "Connecting to Exchange Online..." -Indent 1 -Context $Context

        $connectParams = @{
            ShowProgress = $true
            ShowBanner   = $false
            ErrorAction  = 'Stop'
        }
        if ($UserPrincipalName) { $connectParams['UserPrincipalName'] = $UserPrincipalName }

        Connect-ExchangeOnline @connectParams

        # --- The scope fix ---
        # Re-import the generated temporary cmdlet module at global scope so its cmdlets
        # remain available after this function returns.
        $tmpModules = Get-Module -Name 'tmpEXO_*'
        if ($tmpModules) {
            $tmpModules | Import-Module -Global -Force -DisableNameChecking -ErrorAction Stop
            Write-OagM365Log "Re-imported EXO cmdlet module(s) at global scope: $($tmpModules.Name -join ', ')" -Level Detail -Indent 1 -Context $Context
        } else {
            Write-OagM365Log "No tmpEXO_* cmdlet module found after connecting. EXO cmdlets may be unavailable." -Level Warning -Indent 1 -Context $Context
        }

        $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($connection -and $connection.State -eq 'Connected') {
            $result.Connected         = $true
            $result.UserPrincipalName = $connection.UserPrincipalName
            $result.Organization      = $connection.Organization
            Write-OagM365Log "Connected to $($connection.Organization) as $($connection.UserPrincipalName)" -Level Success -Indent 1 -Context $Context
        } else {
            throw "Exchange Online connection state is '$($connection.State)'."
        }

        # Prove the cmdlets actually survived the function boundary before the export functions
        # rely on them. This is the specific failure the scope fix addresses, so it is verified
        # rather than assumed.
        if (-not (Get-Command -Name 'Get-MalwareFilterPolicy' -ErrorAction SilentlyContinue)) {
            throw "Connected to Exchange Online but the EXO cmdlets are not available in this scope. The temporary cmdlet module did not import correctly."
        }
        Write-OagM365Log "EXO cmdlet availability verified" -Level Success -Indent 1 -Context $Context

        if (-not $NoPause) { Pause }

        return $result

    } catch {
        $exception = Format-OagM365Exception -message "Failed connecting to Exchange Online PowerShell" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
        return $result
    }
}
