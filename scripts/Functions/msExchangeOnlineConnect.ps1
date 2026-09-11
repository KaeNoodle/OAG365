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
    Connects to Exchange Online, falling back to a broker-free sign-in and then to device
      code if the first attempt throws. Module 3.10.x can fail with a NullReferenceException
      raised inside Get-ConnectionContext when the Web Account Manager broker returns
      nothing, and that failure otherwise costs the entire DFO report.
    Re-imports the generated tmpEXO_* module with -Global so its cmdlets survive the
      return. This is the fix.
    Proves a known EXO cmdlet is actually callable, and then actually calls it, before the
      export functions rely on it, rather than assuming the fix worked.
    Records the account and tenant if Get-ConnectionInformation will supply them, but does
      not require it to, since it fails on module versions where the session itself is
      fine.

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
                logWrite "Version $($exo.Version) is newer than the tested 3.9.x line." -level Warning -indent 1
                logWrite "3.10.x requires PowerShell 7.6 or later. If the connection below cannot be recovered, pin the module to 3.9.2 and run under PowerShell 7.4 or 7.5." -level Warning -indent 1
            }
        }

        logWrite "Connecting to Exchange Online..." -indent 1
        $params = @{ ShowProgress = $true; ShowBanner = $false; ErrorAction = 'Stop' }
        if ($userPrincipalName) { $params['UserPrincipalName'] = $userPrincipalName }

        # 3.10.x can fail with a NullReferenceException ("Object reference not set to an
        # instance of an object") thrown out of Get-ConnectionContext. The credentials are
        # not the problem - the Web Account Manager broker hands back nothing and the module
        # dereferences it. Both fallbacks below sign in without that broker, so try them
        # before reporting a failure and skipping the whole DFO report. Parameter support
        # differs by module version, so each one is checked before it is used.
        $supported = (Get-Command Connect-ExchangeOnline).Parameters.Keys
        $attempts  = @(
            @{ label = 'default sign-in';                   extra = @{} }
            @{ label = 'sign-in with the broker disabled';  extra = @{ DisableWAM = $true }; needs = 'DisableWAM' }
            @{ label = 'device code sign-in';               extra = @{ Device = $true };     needs = 'Device' }
        )

        $connectError = $null
        foreach ($attempt in $attempts) {
            if ($attempt.needs -and $attempt.needs -notin $supported) { continue }
            if ($connectError) { logWrite "Retrying with $($attempt.label)." -level Warning -indent 1 }

            $extra = $attempt.extra
            try {
                Connect-ExchangeOnline @params @extra
                $connectError = $null
                break
            } catch {
                $connectError = $_
                logWrite "Exchange Online $($attempt.label) failed: $($_.Exception.Message)" -level Warning -indent 1
            }
        }
        if ($connectError) { throw $connectError }

        # The scope fix. Without this the generated cmdlets vanish when this function
        # returns and every DFO export fails.
        $tmp = Get-Module -Name 'tmpEXO_*'
        if ($tmp) {
            $tmp | Import-Module -Global -Force -DisableNameChecking -ErrorAction Stop
            logWrite "Re-imported EXO cmdlet module at global scope: $($tmp.Name -join ', ')" -level Detail -indent 1
        } else {
            logWrite "No tmpEXO_* module found after connecting. EXO cmdlets may be unavailable." -level Warning -indent 1
        }

        # Verify rather than assume. This is the exact failure the module scope fix above
        # addresses.
        if (-not (Get-Command -Name 'Get-MalwareFilterPolicy' -ErrorAction SilentlyContinue)) {
            throw "Connected, but the EXO cmdlets are not available in this scope. The temporary cmdlet module did not import correctly."
        }

        # A real call, not a status property. What the export needs is a session that
        # answers cmdlets, and this proves exactly that with one the DFO report uses anyway.
        try {
            Get-MalwareFilterPolicy -ErrorAction Stop | Select-Object -First 1 | Out-Null
            logWrite "EXO session verified with a live cmdlet call" -level Success -indent 1
        } catch {
            throw "Connected, but the first Exchange Online cmdlet call failed: $($_.Exception.Message)"
        }

        $result.connected = $true

        # Account and tenant are recorded if available, and the run continues if they are
        # not. Get-ConnectionInformation is built on Get-ConnectionContext, which throws a
        # null reference in 3.10.x on some builds even when the session behind it is
        # healthy. Gating the connection on it threw away a working Exchange Online
        # session and skipped the whole DFO report.
        $connection = $null
        try { $connection = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1 } catch { }

        if ($connection) {
            $result.userPrincipalName = $connection.UserPrincipalName
            $result.organization      = $connection.Organization
            logWrite "Connected to $($connection.Organization) as $($connection.UserPrincipalName)" -level Success -indent 1
            if ($connection.State -and $connection.State -ne 'Connected') {
                logWrite "Reported connection state is '$($connection.State)', but cmdlets are answering." -level Warning -indent 1
            }
        } else {
            logWrite "Connected, but Get-ConnectionInformation returned nothing, so the account and tenant are not recorded against this run. Known in module 3.10.x and it does not affect the export." -level Warning -indent 1
        }

        if (-not $noPause) { Pause }
        return $result

    } catch {
        logWrite (exceptionFormat -message "Failed connecting to Exchange Online" -exception $_) -level Error

        # Two specific causes look like generic connection failures and send operators
        # looking in the wrong place, so name them.
        $detail = "$($_.ErrorDetails) $($_.Exception.Message) $($_.Exception.InnerException)"

        if ($detail -match 'Object reference not set|NullReference' -and (Get-Module -Name 'Microsoft.Graph.Authentication')) {
            logWrite "CAUSE: the Microsoft.Graph modules are loaded in this process and have already authenticated." -level Error -indent 1
            logWrite "ExchangeOnlineManagement and Microsoft.Graph each carry their own copy of MSAL and cannot both authenticate in one process. Whichever goes second fails here. No sign-in method avoids it." -level Error -indent 1
            logWrite "FIX: run the DFO report in its own window with .\runMe.ps1 -report DFO" -level Error -indent 1
        }

        if ($detail -match '530035|blocked by security defaults') {
            logWrite "CAUSE: security defaults are enabled in this tenant and block device code sign-in." -level Error -indent 1
            logWrite "FIX: sign in through the normal prompt rather than device code, or turn security defaults off for the duration of the audit if this is a test tenant." -level Error -indent 1
        }

        return $result
    }
}
