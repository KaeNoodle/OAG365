<#--------------------------------------------------------------------------------------

    OAG-FileInfo.psm1 - module loader

    DESCRIPTION
    Loads every function file and exposes only the five report writers to the caller.

    LOGIC
    Dot-sources Functions\ first, then ReportWriter\, because the report writers call the
      shared functions at load time.
    Exports only the report writer functions. Everything under Functions\ stays internal.

    Because these are dot-sourced into module scope, all functions share the module's
    $script: scope. That is what allows the ported export functions to keep reading
    $script:exportTarget without alteration, and what lets every function reach $script:run
    instead of having a context object passed in as a parameter.

    --------------------------------------------------------------------------------------#>

$functionFiles = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Functions')    -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue)
$reportFiles   = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'ReportWriter') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue)

foreach ($file in @($functionFiles + $reportFiles)) {
    try {
        . $file.FullName
    } catch {
        Write-Error "Failed loading $($file.FullName): $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function @(
    'capReportWrite'
    'iamReportWrite'
    'orgReportWrite'
    'dfoReportWrite'
    'threatHuntReportWrite'
    'runInitialize'
    'msGraphConnect'
    'msExchangeOnlineConnect'
    'msDisconnect'
    'exportVerify'
    'manifestWrite'
    'logWrite'
)
