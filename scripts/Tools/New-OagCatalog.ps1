<#
.SYNOPSIS
Generates the OAG-FileCatalog.cat file catalog ready for signing.

.DESCRIPTION
Run this whenever any module file changes, then have the resulting .cat signed by OAG.

Why a catalog rather than signing each file
-------------------------------------------
Authenticode signs one file at a time. This module has around 60 .ps1 files, so per-file
signing would mean 60 signing operations for every change, making even a small fix expensive
and creating a strong incentive to skip re-signing.

New-FileCatalog instead records a SHA-256 hash of every file into a single .cat, and only that
one file is signed. Test-FileCatalog then verifies the whole module in one call, which is what
Test-OagM365Integrity does at the start of every run.

The trade-off worth understanding: with per-file Authenticode, PowerShell itself refuses to run
an altered script under AllSigned execution policy. With a catalog, PowerShell will still run an
altered file - it is this module's own startup check that detects the change and records it. If
enforcement at execution time is a requirement rather than detection and logging, that needs
per-file signing and should be agreed before release.

.PARAMETER ModuleRoot
Path to the module root. Defaults to the parent of this script's folder.

.PARAMETER CatalogVersion
Catalog format version. 2 uses SHA-256 and should be used; 1 uses SHA-1 and is deprecated.

.EXAMPLE
.\New-OagCatalog.ps1

.EXAMPLE
.\New-OagCatalog.ps1 -ModuleRoot C:\Tools\OAG-M365-AuditingScript

.NOTES
VERSION: 1.0

After generating, sign the catalog:
    Set-AuthenticodeSignature -FilePath .\OAG-FileCatalog.cat `
        -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0] `
        -TimestampServer 'http://timestamp.sectigo.com'

Always timestamp. Without a timestamp the signature stops validating the moment the signing
certificate expires; with one it remains valid for code signed during the certificate's
lifetime.
#>
[CmdletBinding()]
param (
    [string]$ModuleRoot = (Split-Path -Path $PSScriptRoot -Parent),
    [ValidateSet(1, 2)]
    [int]$CatalogVersion = 2
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command -Name 'New-FileCatalog' -ErrorAction SilentlyContinue)) {
    throw "New-FileCatalog is not available. Catalog generation requires Windows PowerShell 5.1 or PowerShell 7 on Windows."
}

$catalogPath = Join-Path -Path $ModuleRoot -ChildPath 'OAG-FileCatalog.cat'

Write-Host "Generating file catalog" -ForegroundColor Cyan
Write-Host "  Module root : $ModuleRoot"
Write-Host "  Catalog     : $catalogPath"
Write-Host "  Version     : $CatalogVersion (SHA-$(if ($CatalogVersion -eq 2) { '256' } else { '1' }))"

# Remove any existing catalog first. Leaving it in place would mean the old catalog is itself
# hashed into the new one.
if (Test-Path $catalogPath) {
    Remove-Item -Path $catalogPath -Force
    Write-Host "  Removed previous catalog" -ForegroundColor DarkYellow
}

$catalog = New-FileCatalog -Path $ModuleRoot -CatalogFilePath $catalogPath -CatalogVersion $CatalogVersion

$fileCount = (Get-ChildItem -Path $ModuleRoot -Recurse -File |
              Where-Object { $_.FullName -ne $catalogPath }).Count

Write-Host ""
Write-Host "Catalog created covering $fileCount file(s)" -ForegroundColor Green
Write-Host "  $($catalog.FullName)"
Write-Host ""
Write-Host "NEXT STEP - the catalog is not signed yet." -ForegroundColor DarkYellow
Write-Host "  Set-AuthenticodeSignature -FilePath '$catalogPath' \`" -ForegroundColor Gray
Write-Host "      -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0] \`" -ForegroundColor Gray
Write-Host "      -TimestampServer 'http://timestamp.sectigo.com'" -ForegroundColor Gray
Write-Host ""
Write-Host "Then verify:" -ForegroundColor Cyan
Write-Host "  Test-FileCatalog -CatalogFilePath '$catalogPath' -Path '$ModuleRoot' -Detailed" -ForegroundColor Gray
