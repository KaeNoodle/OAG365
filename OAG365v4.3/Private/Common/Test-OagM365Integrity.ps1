function Test-OagM365Integrity {
    <#
    .SYNOPSIS
    Verifies that the module code has not been altered since it was signed.

    .DESCRIPTION
    Answers the question "can we show the script was not tampered with to produce false
    results?" for the code side. The evidence side is handled separately by New-OagM365Manifest.

    Two independent checks are run:

      1. File catalog (OagExportM365.cat). New-FileCatalog records a SHA-256 hash of every file
         listed in the manifest's FileList; Test-FileCatalog verifies them all in one call. Only
         the .cat needs an Authenticode signature, which is why the module is distributed with
         one catalog rather than a signature on each of ~60 files.

      2. Authenticode on the catalog and on the entry point script, giving the signer identity
         and certificate validity.

    Results are written to 00_RunLog\code-integrity.json so the reviewer can see which code
    version produced the evidence.

    Important limitation to state plainly in any working paper: a valid result here shows the
    code matches what was signed. It does not show the signed code is correct, and it does not
    on its own show the exported data reflects the tenant. Scope coverage and reproducibility
    of a second run carry that part of the argument.

    .PARAMETER ModuleRoot
    Path to the module root folder.

    .PARAMETER Context
    Run context object, used for logging.

    .EXAMPLE
    Test-OagM365Integrity -ModuleRoot $PSScriptRoot -Context $Context

    .NOTES
        NAME: Test-OagM365Integrity
        VERSION: 1.0

        CMDLETS:
           Test-FileCatalog
           Get-AuthenticodeSignature
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$ModuleRoot,
        [object]$Context
    )

    $result = [PSCustomObject]@{
        CheckedAt         = (Get-Date).ToString('o')
        ModuleRoot        = $ModuleRoot
        CatalogPresent    = $false
        CatalogStatus     = 'NotChecked'
        CatalogMismatches = @()
        SignatureStatus   = 'NotChecked'
        SignerSubject     = $null
        SignerThumbprint  = $null
        CertNotAfter      = $null
        FilesChecked      = 0
        Verdict           = 'Unverified'
    }

    Write-OagM365Log "Verifying module code integrity" -Context $Context

    $catalogPath = Join-Path -Path $ModuleRoot -ChildPath 'OagExportM365.cat'

    # --- Check 1: file catalog ---------------------------------------------------------------
    if (Test-Path -Path $catalogPath) {
        $result.CatalogPresent = $true
        try {
            if (Get-Command -Name 'Test-FileCatalog' -ErrorAction SilentlyContinue) {
                $catalogResult = Test-FileCatalog -CatalogFilePath $catalogPath -Path $ModuleRoot -Detailed -ErrorAction Stop
                $result.CatalogStatus = $catalogResult.Status.ToString()
                $result.FilesChecked  = @($catalogResult.CatalogItems.Keys).Count

                if ($catalogResult.Status -ne 'Valid') {
                    # Name the specific files whose hashes differ. "Catalog invalid" alone is
                    # not actionable for whoever has to investigate.
                    foreach ($key in $catalogResult.CatalogItems.Keys) {
                        $catHash = $catalogResult.CatalogItems[$key]
                        $pathHash = $catalogResult.PathItems[$key]
                        if ($catHash -ne $pathHash) { $result.CatalogMismatches += $key }
                    }
                    Write-OagM365Log "CATALOG MISMATCH on $($result.CatalogMismatches.Count) file(s). The module does not match what was signed." -Level Error -Indent 1 -Context $Context
                    $result.CatalogMismatches | ForEach-Object {
                        Write-OagM365Log $_ -Level Error -Indent 2 -Context $Context
                    }
                } else {
                    Write-OagM365Log "File catalog valid - $($result.FilesChecked) file(s) match the signed hashes" -Level Success -Indent 1 -Context $Context
                }
            } else {
                # Test-FileCatalog is Windows-only.
                $result.CatalogStatus = 'NotSupportedOnThisPlatform'
                Write-OagM365Log "Test-FileCatalog is unavailable on this platform. Catalog not verified." -Level Warning -Indent 1 -Context $Context
            }
        } catch {
            $result.CatalogStatus = 'Error'
            Write-OagM365Log "Catalog verification failed: $($_.Exception.Message)" -Level Warning -Indent 1 -Context $Context
        }
    } else {
        $result.CatalogStatus = 'NotPresent'
        Write-OagM365Log "No file catalog found. Module integrity is UNVERIFIED - see Docs\Verification.md." -Level Warning -Indent 1 -Context $Context
    }

    # --- Check 2: Authenticode signature -----------------------------------------------------
    try {
        if (Get-Command -Name 'Get-AuthenticodeSignature' -ErrorAction SilentlyContinue) {
            $signTarget = if (Test-Path $catalogPath) {
                $catalogPath
            } else {
                Join-Path -Path $ModuleRoot -ChildPath 'Export-M365.ps1'
            }

            if (Test-Path $signTarget) {
                $sig = Get-AuthenticodeSignature -FilePath $signTarget -ErrorAction Stop
                $result.SignatureStatus = $sig.Status.ToString()

                if ($sig.SignerCertificate) {
                    $result.SignerSubject    = $sig.SignerCertificate.Subject
                    $result.SignerThumbprint = $sig.SignerCertificate.Thumbprint
                    $result.CertNotAfter     = $sig.SignerCertificate.NotAfter.ToString('yyyy-MM-dd')
                }

                if ($sig.Status -eq 'Valid') {
                    Write-OagM365Log "Signature valid - $($sig.SignerCertificate.Subject)" -Level Success -Indent 1 -Context $Context
                    Write-OagM365Log "Thumbprint : $($result.SignerThumbprint)" -Level Detail -Indent 2 -Context $Context
                    Write-OagM365Log "Expires    : $($result.CertNotAfter)" -Level Detail -Indent 2 -Context $Context

                    # Warn before the certificate expires rather than after, since a lapsed
                    # certificate blocks the next audit cycle.
                    $daysLeft = ($sig.SignerCertificate.NotAfter - (Get-Date)).Days
                    if ($daysLeft -lt 60) {
                        Write-OagM365Log "Signing certificate expires in $daysLeft day(s). Arrange re-signing." -Level Warning -Indent 2 -Context $Context
                    }
                } else {
                    Write-OagM365Log "Signature status: $($sig.Status). $($sig.StatusMessage)" -Level Warning -Indent 1 -Context $Context
                }
            }
        } else {
            $result.SignatureStatus = 'NotSupportedOnThisPlatform'
        }
    } catch {
        $result.SignatureStatus = 'Error'
        Write-OagM365Log "Signature check failed: $($_.Exception.Message)" -Level Warning -Indent 1 -Context $Context
    }

    $result.Verdict = if ($result.CatalogStatus -eq 'Valid' -and $result.SignatureStatus -eq 'Valid') {
        'Verified'
    } elseif ($result.CatalogMismatches.Count -gt 0) {
        'FAILED - code altered since signing'
    } else {
        'Unverified'
    }

    Write-OagM365Log "Code integrity verdict: $($result.Verdict)" `
        -Level $(if ($result.Verdict -eq 'Verified') { 'Success' } else { 'Warning' }) `
        -Indent 1 -Context $Context

    return $result
}
