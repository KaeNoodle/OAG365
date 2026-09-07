<#
    OagExportM365 module loader.

    Private functions are loaded before Public because the public report functions call them.
    Only the functions named in FunctionsToExport in the manifest become visible to the caller;
    everything under Private stays internal to the module.

    Because these are dot-sourced into module scope, all functions share the module's $script:
    scope. That is what allows the ported export functions to keep referring to
    $script:exportTarget without alteration - see Initialize-OagM365ReportPath.
#>

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue)
$public  = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public')  -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue)

foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    } catch {
        Write-Error "Failed to load module file $($file.FullName): $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function $public.BaseName
