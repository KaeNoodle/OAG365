<#
.SYNOPSIS
Builds a distributable module that collapses to three signable files.

.DESCRIPTION
Solves a problem that only appears once application control is enforced.

The problem
-----------
Under WDAC, AppLocker or Airlock, PowerShell evaluates trust for EVERY file it loads, and the
language mode is bound to each scriptblock when that file is parsed. Trust is not inherited from
the caller. A signed and trusted runMe.ps1 that dot-sources dozens of unsigned files does not make
those files trusted; each is evaluated on its own and parsed in ConstrainedLanguage.

It then fails harder than you would expect. PowerShell refuses to dot-source a command defined
under a different language mode, to stop untrusted code being promoted into a trusted session:

    Cannot dot-source this command because it was defined in a different language mode.

So partial trust does not give partial function. It gives a broken module. It is all or nothing
across every file the module loads.

The development layout is dozens of .ps1/.psm1/.psd1 files. Signing all of them individually
would mean one signing operation per file for every change, which is exactly the cost the file
catalog was meant to avoid - and a catalog only helps with execution if it is installed into the
system catalog store, which needs administrator rights on every workstation.

The fix
-------
Keep the split source for development, and build a merged distribution for signing and release.
Every function from Functions\ and ReportWriter\ is concatenated into a single .psm1, leaving
three files to sign:

    OAG-ModuleManifest.psd1           manifest
    OAG-M365-AuditingScript.psm1      every function merged
    runMe.ps1                         entry point

Three signing operations instead of dozens, no cross-file dot-sourcing at runtime, and the source
tree stays readable and reviewable. This is a common pattern for production PowerShell modules and
it happens to solve the application control problem cleanly.

Documentation\ and Tools\ are copied across but not merged. Tools scripts are run directly by an
operator, so they need their own signatures if they are to run under enforcement - Debugger.ps1
is written to work in ConstrainedLanguage precisely because it may have to.

.PARAMETER SourceRoot
Module source root. Defaults to the parent of this script's folder.

.PARAMETER OutputPath
Where to write the built module. Defaults to <SourceRoot>\..\dist.

.PARAMETER Clean
Remove the output folder before building.

.EXAMPLE
.\Build-Module.ps1

.EXAMPLE
.\Build-Module.ps1 -OutputPath C:\Build\OAG-M365-AuditingScript -Clean

.NOTES
VERSION: 1.0

After building, sign the three files:

    Get-ChildItem .\dist\OAG-M365-AuditingScript -Include *.ps1,*.psm1,*.psd1 -Recurse |
        Set-AuthenticodeSignature -Certificate $cert -TimestampServer 'http://timestamp.sectigo.com'
#>
[CmdletBinding()]
param (
    [string]$SourceRoot = (Split-Path -Path $PSScriptRoot -Parent),
    [string]$OutputPath,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $OutputPath = Join-Path -Path (Split-Path -Path $SourceRoot -Parent) -ChildPath 'dist'
}
$buildPath = Join-Path -Path $OutputPath -ChildPath 'OAG-M365-AuditingScript'

Write-Host ""
Write-Host "Building OAG-M365-AuditingScript for distribution" -ForegroundColor Cyan
Write-Host "===================================================================="
Write-Host "  Source : $SourceRoot"
Write-Host "  Output : $buildPath"
Write-Host ""

if ($Clean -and (Test-Path $OutputPath)) {
    Remove-Item -Path $OutputPath -Recurse -Force
    Write-Host "  Cleaned output folder" -ForegroundColor DarkYellow
}

New-Item -ItemType Directory -Force -Path $buildPath | Out-Null

# ------------------------------------------------------------------------------------------
# Collect source files. Functions\ first, then ReportWriter\, matching the loader's order in
# OAG-FileInfo.psm1. Order does not actually matter for function definitions in a single file -
# PowerShell resolves calls at invocation, not at parse - but keeping it consistent makes the
# merged file easier to compare against the source tree during review.
# ------------------------------------------------------------------------------------------
$functionFiles = @(Get-ChildItem -Path (Join-Path $SourceRoot 'Functions')    -Filter '*.ps1' -Recurse |
                   Sort-Object FullName)
$reportFiles   = @(Get-ChildItem -Path (Join-Path $SourceRoot 'ReportWriter') -Filter '*.ps1' -Recurse |
                   Sort-Object FullName)

$sourceFiles = @($functionFiles + $reportFiles)

# One file does not mean one function here - graphObjectResolve.ps1 alone defines seven - so
# count definitions rather than files, otherwise the validation at the end compares the wrong
# two numbers and always fails.
$sourceFunctionCount = @($sourceFiles | ForEach-Object {
    Select-String -Path $_.FullName -Pattern '^\s*function\s+\w+' -AllMatches
}).Count

Write-Host "  Shared functions  : $($functionFiles.Count) file(s)"
Write-Host "  Report writers    : $($reportFiles.Count) file(s)"
Write-Host "  Functions defined : $sourceFunctionCount"
Write-Host ""

# ------------------------------------------------------------------------------------------
# Merge into a single .psm1
# ------------------------------------------------------------------------------------------
$sb = [System.Text.StringBuilder]::new()

$null = $sb.AppendLine(@"
<#
    OAG-M365-AuditingScript - merged module
    ==============================

    GENERATED FILE - DO NOT EDIT DIRECTLY.

    Built from the split source tree by Tools\Build-Module.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').
    Edit the source under Functions\ and ReportWriter\, then rebuild.

    All $sourceFunctionCount functions are merged into this single file so the module can be
    distributed as three signable files rather than $($sourceFiles.Count + 3). Under application control,
    PowerShell evaluates trust per file and refuses to dot-source across language modes, so a
    module that loads many separate files needs every one of them trusted. See
    Documentation\Verification.md.

    Source files merged:
$(($sourceFiles | ForEach-Object { "      " + $_.FullName.Substring($SourceRoot.Length).TrimStart('\','/') }) -join "`n")
#>

Set-StrictMode -Version Latest

"@)

foreach ($file in $sourceFiles) {
    $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart('\', '/')
    $content = Get-Content -Path $file.FullName -Raw

    $null = $sb.AppendLine("#region $relative")
    $null = $sb.AppendLine()
    $null = $sb.AppendLine($content.TrimEnd())
    $null = $sb.AppendLine()
    $null = $sb.AppendLine("#endregion $relative")
    $null = $sb.AppendLine()
}

# Take the export list from the manifest, not from filenames. A filename is not a function name
# in this source tree - ConditionalAccessPolicy.ps1 defines capReportWrite - so deriving exports
# from BaseName would export names that do not exist.
$sourceManifest = Import-PowerShellDataFile -Path (Join-Path $SourceRoot 'OAG-ModuleManifest.psd1')
$exportNames = $sourceManifest.FunctionsToExport | ForEach-Object { "'$_'" }
$null = $sb.AppendLine("Export-ModuleMember -Function @($($exportNames -join ', '))")

$psm1Path = Join-Path -Path $buildPath -ChildPath 'OAG-M365-AuditingScript.psm1'
Set-Content -Path $psm1Path -Value $sb.ToString() -Encoding utf8

$psm1Lines = (Get-Content -Path $psm1Path).Count
Write-Host "  Merged .psm1 : $psm1Lines lines" -ForegroundColor Green

# ------------------------------------------------------------------------------------------
# Copy the manifest and entry point, and the folders that are not merged
# ------------------------------------------------------------------------------------------
Copy-Item -Path (Join-Path $SourceRoot 'OAG-ModuleManifest.psd1') -Destination $buildPath -Force
Copy-Item -Path (Join-Path $SourceRoot 'runMe.ps1')               -Destination $buildPath -Force

foreach ($folder in @('Docs', 'Tools')) {
    $src = Join-Path $SourceRoot $folder
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination $buildPath -Recurse -Force
    }
}

# The build script itself is a development tool, not part of the distribution.
$builtBuildScript = Join-Path $buildPath 'Tools\Build-Module.ps1'
if (Test-Path $builtBuildScript) { Remove-Item -Path $builtBuildScript -Force }

# ------------------------------------------------------------------------------------------
# Update the manifest FileList to match the built layout, not the source layout.
# ------------------------------------------------------------------------------------------
$psd1Path = Join-Path $buildPath 'OAG-ModuleManifest.psd1'
$psd1 = Get-Content -Path $psd1Path -Raw

$newFileList = @"
    FileList = @(
        'OAG-ModuleManifest.psd1'
        'OAG-M365-AuditingScript.psm1'
        'runMe.ps1'
    )
"@

$psd1 = [regex]::Replace($psd1, '(?ms)^\s*FileList = @\(.*?^\s*\)', $newFileList)
Set-Content -Path $psd1Path -Value $psd1 -Encoding utf8

# ------------------------------------------------------------------------------------------
# Validate the build
# ------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "VALIDATION" -ForegroundColor Cyan

$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($psm1Path, [ref]$null, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Host "  Merged .psm1 has $($errors.Count) parse error(s):" -ForegroundColor Red
    $errors | Select-Object -First 5 | ForEach-Object {
        Write-Host "    line $($_.Extent.StartLineNumber): $($_.Message)" -ForegroundColor Red
    }
    throw "Build failed: merged module does not parse."
}
Write-Host "  Merged .psm1 parses cleanly" -ForegroundColor Green

$manifest = Test-ModuleManifest -Path $psd1Path -ErrorAction SilentlyContinue
if ($manifest) {
    Write-Host "  Manifest valid - v$($manifest.Version), $($manifest.ExportedFunctions.Count) exported function(s)" -ForegroundColor Green
} else {
    Write-Host "  Manifest could not be fully validated (dependencies may not be installed here)" -ForegroundColor DarkYellow
}

# Confirm every source function made it into the merge
$mergedAst = [System.Management.Automation.Language.Parser]::ParseFile($psm1Path, [ref]$null, [ref]$null)
$mergedFunctions = $mergedAst.FindAll(
    { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
if ($mergedFunctions.Count -eq $sourceFunctionCount) {
    Write-Host "  All $sourceFunctionCount function(s) present in merged module" -ForegroundColor Green
} else {
    Write-Host "  Expected $sourceFunctionCount function(s), found $($mergedFunctions.Count)" -ForegroundColor Red
    throw "Build failed: function count mismatch."
}

# ------------------------------------------------------------------------------------------
# Report what needs signing
# ------------------------------------------------------------------------------------------
$signable = Get-ChildItem -Path $buildPath -Recurse -File |
            Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') }

Write-Host ""
Write-Host "===================================================================="
Write-Host "BUILD COMPLETE" -ForegroundColor Green
Write-Host ""
Write-Host "  Output: $buildPath"
Write-Host ""
Write-Host "  Files requiring signature ($($signable.Count)):" -ForegroundColor Cyan
foreach ($f in ($signable | Sort-Object FullName)) {
    $rel = $f.FullName.Substring($buildPath.Length).TrimStart('\', '/')
    Write-Host "    $rel"
}
Write-Host ""
Write-Host "  Source tree had $($sourceFiles.Count) files needing trust. Built module has $($signable.Count)." -ForegroundColor Green
Write-Host ""
Write-Host "NEXT STEPS" -ForegroundColor Cyan
Write-Host '  $cert = (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0]' -ForegroundColor Gray
Write-Host "  Get-ChildItem '$buildPath' -Include *.ps1,*.psm1,*.psd1 -Recurse |" -ForegroundColor Gray
Write-Host '      Set-AuthenticodeSignature -Certificate $cert -TimestampServer "http://timestamp.sectigo.com"' -ForegroundColor Gray
Write-Host ""
