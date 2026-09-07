function codeVerify {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Checks that the module code has not been altered since it was signed, and records
    the result so the evidence shows which signed code produced it.

    Answers the code half of "can we show the script was not tampered with". The evidence
    half is manifestWrite.

    Important limitation to state in any working paper: a valid result shows the code
    matches what was signed. It does not show the signed code is correct, and it does not
    show the exported data reflects the tenant.

    LOGIC
    Runs two independent checks.
      1. File catalog. Test-FileCatalog compares every file against the hashes recorded
         in OAG-FileCatalog.cat. Where it fails, the specific mismatched files are named,
         because "catalog invalid" on its own is not actionable.
      2. Authenticode. Reads the signer identity, thumbprint and expiry from the catalog
         or, failing that, the main run file.
    Warns if the signing certificate expires within 60 days, since a lapsed certificate
      blocks the next audit cycle.
    Both cmdlets are Windows only. Elsewhere the verdict is recorded as Unverified and
      the run continues.

    PARAMETERS
    -moduleRoot (required) path to the module folder

    RUNNING CONTEXT
    Called by  : runInitialize
    Calls      : logWrite
    Writes     : 00_RunLog\code-integrity.json

    CMLETS/PERMISSIONS/SCOPES
    Test-FileCatalog
    Get-AuthenticodeSignature

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$moduleRoot
    )

    $result = [PSCustomObject]@{
        checkedAt = (Get-Date).ToString('o'); moduleRoot = $moduleRoot
        catalogPresent = $false; catalogStatus = 'NotChecked'; catalogMismatches = @()
        signatureStatus = 'NotChecked'; signerSubject = $null; signerThumbprint = $null
        certNotAfter = $null; filesChecked = 0; verdict = 'Unverified'
    }

    logWrite "Verifying module code integrity"
    $catalogPath = Join-Path $moduleRoot 'OAG-FileCatalog.cat'

    if (Test-Path $catalogPath) {
        $result.catalogPresent = $true
        try {
            if (Get-Command Test-FileCatalog -ErrorAction SilentlyContinue) {
                $check = Test-FileCatalog -CatalogFilePath $catalogPath -Path $moduleRoot -Detailed -ErrorAction Stop
                $result.catalogStatus = $check.Status.ToString()
                $result.filesChecked  = @($check.CatalogItems.Keys).Count

                if ($check.Status -ne 'Valid') {
                    foreach ($key in $check.CatalogItems.Keys) {
                        if ($check.CatalogItems[$key] -ne $check.PathItems[$key]) {
                            $result.catalogMismatches += $key
                        }
                    }
                    logWrite "CATALOG MISMATCH on $($result.catalogMismatches.Count) file(s). The module does not match what was signed." -level Error -indent 1
                    $result.catalogMismatches | ForEach-Object { logWrite $_ -level Error -indent 2 }
                } else {
                    logWrite "File catalog valid - $($result.filesChecked) file(s) match" -level Success -indent 1
                }
            } else {
                $result.catalogStatus = 'NotSupportedOnThisPlatform'
                logWrite "Test-FileCatalog unavailable on this platform. Catalog not verified." -level Warning -indent 1
            }
        } catch {
            $result.catalogStatus = 'Error'
            logWrite "Catalog verification failed: $($_.Exception.Message)" -level Warning -indent 1
        }
    } else {
        $result.catalogStatus = 'NotPresent'
        logWrite "No file catalog found. Module integrity is UNVERIFIED." -level Warning -indent 1
    }

    try {
        if (Get-Command Get-AuthenticodeSignature -ErrorAction SilentlyContinue) {
            $target = if (Test-Path $catalogPath) { $catalogPath } else { Join-Path $moduleRoot 'OAG-MainRunFile.ps1' }
            if (Test-Path $target) {
                $sig = Get-AuthenticodeSignature -FilePath $target -ErrorAction Stop
                $result.signatureStatus = $sig.Status.ToString()

                if ($sig.SignerCertificate) {
                    $result.signerSubject    = $sig.SignerCertificate.Subject
                    $result.signerThumbprint = $sig.SignerCertificate.Thumbprint
                    $result.certNotAfter     = $sig.SignerCertificate.NotAfter.ToString('yyyy-MM-dd')
                }

                if ($sig.Status -eq 'Valid') {
                    logWrite "Signature valid - $($sig.SignerCertificate.Subject)" -level Success -indent 1
                    logWrite "Thumbprint : $($result.signerThumbprint)" -level Detail -indent 2
                    logWrite "Expires    : $($result.certNotAfter)" -level Detail -indent 2

                    $daysLeft = ($sig.SignerCertificate.NotAfter - (Get-Date)).Days
                    if ($daysLeft -lt 60) {
                        logWrite "Signing certificate expires in $daysLeft day(s). Arrange re-signing." -level Warning -indent 2
                    }
                } else {
                    logWrite "Signature status: $($sig.Status). $($sig.StatusMessage)" -level Warning -indent 1
                }
            }
        } else {
            $result.signatureStatus = 'NotSupportedOnThisPlatform'
        }
    } catch {
        $result.signatureStatus = 'Error'
        logWrite "Signature check failed: $($_.Exception.Message)" -level Warning -indent 1
    }

    $result.verdict = if ($result.catalogStatus -eq 'Valid' -and $result.signatureStatus -eq 'Valid') { 'Verified' }
                      elseif ($result.catalogMismatches.Count -gt 0) { 'FAILED - code altered since signing' }
                      else { 'Unverified' }

    logWrite "Code integrity verdict: $($result.verdict)" -level $(if ($result.verdict -eq 'Verified') { 'Success' } else { 'Warning' }) -indent 1
    return $result
}
