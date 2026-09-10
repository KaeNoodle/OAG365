function msExchangeOnlineConnect {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Connects to Exchange Online for the Defender for Office 365 report, and fixes the
    temporary cmdlet module scope problem that stopped that report producing output.

    Connect-ExchangeOnline ships no static cmdlets. It generates a temporary module
    named tmpEXO_<random> at connection time and imports it into whatever scope called
    it. In the original script that call sat inside a function, so the generated module
    was imported into that function's local scope and was discarded the moment the
    function returned. Every later Get-MalwareFilterPolicy and similar then failed with
    "term is not recognised", which is why the DFO report produced nothing.

    LOGIC
    Checks ExchangeOnlineManagement is installed and imports it at global scope.
    Warns if the version is outside the tested 3.9.x line.
    Connects to Exchange Online.
    Re-imports the generated tmpEXO_* module with -Global so its cmdlets survive the
      return. This is the fix.
    Proves a known EXO cmdlet is actually callable before the export functions rely on
      it, rather than assuming the fix worked.

    PARAMETERS
    -modules (optional) defaults to ExchangeOnlineManagement
    -userPrincipalName (optional) pre-fills the login prompt
    -noPause (optional) skip the confirmation pause, for unattended runs

    RUNNING CONTEXT
    Called by  : runMe.ps1, dfoReportWrite
    Calls      : logWrite, exceptionFormat
    Returns    : object with connected, userPrincipalName, organization, moduleVersion

    CMLETS/PERMISSIONS/SCOPES
    Connect-ExchangeOnline
    Get-ConnectionInformation
    Role: Exchange Online View-Only Organization Management, or Global Reader

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [string[]]$modules = @('ExchangeOnlineManagement'),
        [string]$userPrincipalName,
        [switch]$noPause
    )

    $result = [PSCustomObject]@{
        connected         = $false
        userPrincipalName = $null
        organization      = $null
        moduleVersion     = $null
    }

    try {
        logWrite "Exchange Online connection required. Please log in if prompted."

        $missing = $modules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
        if ($missing) {
            throw "Missing PowerShell module(s): $($missing -join ', '). Install with: Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Scope CurrentUser"
        }

        $modules | Import-Module -Global -ErrorAction Stop

        $exo = Get-Module -Name ExchangeOnlineManagement | Select-Object -First 1
        if ($exo) {
            $result.moduleVersion = $exo.Version.ToString()
            logWrite "ExchangeOnlineManagement $($exo.Version)" -level Detail -indent 1
            if ($exo.Version -ge [version]'3.10.0') {
                logWrite "Version $($exo.Version) is newer than the tested 3.9.x. If cmdlets fail, pin to 3.9.0." -level Warning -indent 1
            }
        }

        logWrite "Connecting to Exchange Online..." -indent 1
        $params = @{ ShowProgress = $true; ShowBanner = $false; ErrorAction = 'Stop' }
        if ($userPrincipalName) { $params['UserPrincipalName'] = $userPrincipalName }
        Connect-ExchangeOnline @params

        # The scope fix. Without this the generated cmdlets vanish when this function
        # returns and every DFO export fails.
        $tmp = Get-Module -Name 'tmpEXO_*'
        if ($tmp) {
            $tmp | Import-Module -Global -Force -DisableNameChecking -ErrorAction Stop
            logWrite "Re-imported EXO cmdlet module at global scope: $($tmp.Name -join ', ')" -level Detail -indent 1
        } else {
            logWrite "No tmpEXO_* module found after connecting. EXO cmdlets may be unavailable." -level Warning -indent 1
        }

        $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $connection -or $connection.State -ne 'Connected') {
            throw "Exchange Online connection state is '$($connection.State)'."
        }

        $result.connected         = $true
        $result.userPrincipalName = $connection.UserPrincipalName
        $result.organization      = $connection.Organization
        logWrite "Connected to $($connection.Organization) as $($connection.UserPrincipalName)" -level Success -indent 1

        # Verify rather than assume. This is the exact failure the fix above addresses.
        if (-not (Get-Command -Name 'Get-MalwareFilterPolicy' -ErrorAction SilentlyContinue)) {
            throw "Connected, but the EXO cmdlets are not available in this scope. The temporary cmdlet module did not import correctly."
        }
        logWrite "EXO cmdlet availability verified" -level Success -indent 1

        if (-not $noPause) { Pause }
        return $result

    } catch {
        logWrite (exceptionFormat -message "Failed connecting to Exchange Online" -exception $_) -level Error
        return $result
    }
}
