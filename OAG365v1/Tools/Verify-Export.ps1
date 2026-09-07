<#
.SYNOPSIS
Verifies an exported evidence set against its manifest.

.DESCRIPTION
Run by a reviewer, not by the operator who produced the export. Recomputes the SHA-256 of every
file in a run folder and compares it to manifest.sha256, then reports the run's own record of
what it collected.

Four outcomes are distinguished, because they mean different things:

    Matched    file present and hash agrees with the manifest
    ALTERED    file present but hash differs - the file changed after the run
    MISSING    listed in the manifest but not on disk
    UNLISTED   present on disk but not in the manifest - added after the run

What a clean result does and does not tell you
----------------------------------------------
A clean result shows the files are byte-for-byte what the export produced. It does not show the
export was complete or accurate. An export made with insufficient permissions produces an
incomplete evidence set whose hashes verify perfectly.

So this script also surfaces the run's own completeness record: the scopes that were requested
versus granted, exports that returned no records, and any errors. Read that alongside the hash
result before relying on the evidence.

The remaining limitation, which should be stated in the working paper rather than glossed over:
whoever can alter the exports can also regenerate the manifest. The manifest defends against
later accidental or casual alteration, not against the operator. Closing that requires the
manifest to be signed, or committed promptly to a records system the operator cannot amend.

.PARAMETER RunFolder
Path to the run folder, e.g. C:\Audit\M365\20260823_142530

.EXAMPLE
.\Verify-Export.ps1 -runFolder C:\Audit\M365\20260823_142530

.NOTES
VERSION: 1.0
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$runFolder
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -Path $runFolder)) { throw "Run folder not found: $runFolder" }
$runFolder = (Resolve-Path -Path $runFolder).Path

$manifestPath = Join-Path -Path $runFolder -ChildPath 'manifest.sha256'
if (-not (Test-Path -Path $manifestPath)) {
    throw "No manifest.sha256 found in $runFolder. This evidence set cannot be verified."
}

Write-Host ""
Write-Host "OAG M365 evidence verification" -ForegroundColor Cyan
Write-Host ("=" * 70)
Write-Host "  Run folder : $runFolder"
Write-Host "  Manifest   : $manifestPath"
Write-Host ""

# --- Parse the manifest --------------------------------------------------------------------
$manifest = @{}
foreach ($line in (Get-Content -Path $manifestPath)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    # Format: <64 hex chars><two spaces><relative path>
    if ($line -match '^([0-9A-Fa-f]{64})\s\s(.+)$') {
        $manifest[$Matches[2]] = $Matches[1].ToUpper()
    }
}
Write-Host "Manifest lists $($manifest.Count) file(s)" -ForegroundColor Gray

# --- Recompute and compare -----------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($relative in ($manifest.Keys | Sort-Object)) {
    $full = Join-Path -Path $runFolder -ChildPath $relative
    if (Test-Path -Path $full) {
        $actual = (Get-FileHash -Path $full -Algorithm SHA256).Hash.ToUpper()
        $status = if ($actual -eq $manifest[$relative]) { 'Matched' } else { 'ALTERED' }
    } else {
        $actual = $null
        $status = 'MISSING'
    }
    $results.Add([PSCustomObject]@{
        Status   = $status
        File     = $relative
        Expected = $manifest[$relative]
        Actual   = $actual
    })
}

# Files on disk that the manifest does not list.
# The two manifest artifacts are excluded because neither can contain its own hash: they are
# written at the same moment the hashes are computed. Everything else on disk should be listed,
# and anything that is not was added after the run finished.
$manifestArtifacts = @('manifest.sha256', 'manifest.csv')
$onDisk = Get-ChildItem -Path $runFolder -Recurse -File |
          Where-Object { $_.Name -notin $manifestArtifacts } |
          ForEach-Object { $_.FullName.Substring($runFolder.Length).TrimStart('\', '/') }

foreach ($file in $onDisk) {
    if (-not $manifest.ContainsKey($file)) {
        $results.Add([PSCustomObject]@{
            Status = 'UNLISTED'; File = $file; Expected = $null
            Actual = (Get-FileHash -Path (Join-Path $runFolder $file) -Algorithm SHA256).Hash.ToUpper()
        })
    }
}

$matched  = @($results | Where-Object Status -eq 'Matched')
$altered  = @($results | Where-Object Status -eq 'ALTERED')
$missing  = @($results | Where-Object Status -eq 'MISSING')
$unlisted = @($results | Where-Object Status -eq 'UNLISTED')

Write-Host ""
Write-Host "INTEGRITY" -ForegroundColor Cyan
Write-Host ("  Matched  : {0}" -f $matched.Count)  -ForegroundColor Green
Write-Host ("  Altered  : {0}" -f $altered.Count)  -ForegroundColor $(if ($altered.Count)  { 'Red' } else { 'Gray' })
Write-Host ("  Missing  : {0}" -f $missing.Count)  -ForegroundColor $(if ($missing.Count)  { 'Red' } else { 'Gray' })
Write-Host ("  Unlisted : {0}" -f $unlisted.Count) -ForegroundColor $(if ($unlisted.Count) { 'DarkYellow' } else { 'Gray' })

foreach ($group in @(@{ n = 'ALTERED'; c = $altered; col = 'Red' },
                     @{ n = 'MISSING'; c = $missing; col = 'Red' },
                     @{ n = 'UNLISTED'; c = $unlisted; col = 'DarkYellow' })) {
    if ($group.c.Count -gt 0) {
        Write-Host ""
        Write-Host "  $($group.n):" -ForegroundColor $group.col
        $group.c | ForEach-Object { Write-Host "    $($_.File)" -ForegroundColor $group.col }
    }
}

# --- The run's own record of what it collected ---------------------------------------------
$logPath = Join-Path -Path $runFolder -ChildPath '00_RunLog'

Write-Host ""
Write-Host "COMPLETENESS (as recorded by the run itself)" -ForegroundColor Cyan

$summaryPath = Join-Path $logPath 'run-summary.json'
if (Test-Path $summaryPath) {
    $summary = Get-Content -Path $summaryPath -Raw | ConvertFrom-Json
    Write-Host "  Run ID     : $($summary.RunId)"
    Write-Host "  Tenant     : $($summary.TenantId)"
    Write-Host "  Account    : $($summary.Account)"
    Write-Host "  Auth mode  : $($summary.AuthMode)"
    Write-Host "  Reports    : $($summary.ReportsSucceeded)"
    Write-Host ("  Files      : {0} OK, {1} empty, {2} missing" -f $summary.FilesOk, $summary.FilesEmpty, $summary.FilesMissing)
    Write-Host ("  Status     : {0}" -f $summary.CompletenessStatus) `
        -ForegroundColor $(if ($summary.CompletenessStatus -eq 'Complete') { 'Green' } else { 'DarkYellow' })
} else {
    Write-Host "  No run summary found." -ForegroundColor DarkYellow
}

$scopePath = Join-Path $logPath 'scopes-granted.csv'
if (Test-Path $scopePath) {
    $denied = @(Import-Csv -Path $scopePath | Where-Object { $_.Granted -eq 'False' })
    if ($denied.Count -gt 0) {
        Write-Host ""
        Write-Host "  $($denied.Count) requested scope(s) were NOT granted. Affected exports are incomplete:" -ForegroundColor DarkYellow
        $denied | ForEach-Object { Write-Host "    $($_.Scope)" -ForegroundColor DarkYellow }
    } else {
        Write-Host "  All requested scopes were granted." -ForegroundColor Green
    }
}

$integrityPath = Join-Path $logPath 'code-integrity.json'
if (Test-Path $integrityPath) {
    $integrity = Get-Content -Path $integrityPath -Raw | ConvertFrom-Json
    Write-Host ""
    Write-Host "CODE INTEGRITY (at time of run)" -ForegroundColor Cyan
    Write-Host "  Verdict    : $($integrity.Verdict)" `
        -ForegroundColor $(if ($integrity.Verdict -eq 'Verified') { 'Green' } else { 'DarkYellow' })
    Write-Host "  Catalog    : $($integrity.CatalogStatus)"
    Write-Host "  Signature  : $($integrity.SignatureStatus)"
    if ($integrity.SignerSubject) { Write-Host "  Signer     : $($integrity.SignerSubject)" }
}

# --- Verdict -------------------------------------------------------------------------------
Write-Host ""
Write-Host ("=" * 70)
if ($altered.Count -eq 0 -and $missing.Count -eq 0) {
    Write-Host "RESULT: Evidence unchanged since export." -ForegroundColor Green
    if ($unlisted.Count -gt 0) {
        Write-Host "        Note: $($unlisted.Count) file(s) were added after the run and are not covered by the manifest." -ForegroundColor DarkYellow
    }
    Write-Host "        This confirms the files have not been altered. It does not confirm" -ForegroundColor Gray
    Write-Host "        the export was complete - read the completeness section above." -ForegroundColor Gray
    exit 0
} else {
    Write-Host "RESULT: VERIFICATION FAILED. The evidence does not match the manifest." -ForegroundColor Red
    Write-Host "        Do not rely on this evidence set. Escalate and re-run the export." -ForegroundColor Red
    exit 1
}
