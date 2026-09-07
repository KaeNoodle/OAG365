function New-OagM365Manifest {
    <#
    .SYNOPSIS
    Produces a SHA-256 manifest of every file in the run folder.

    .DESCRIPTION
    This is the evidence-integrity half of the tamper-verification requirement. The code half is
    handled by Test-OagM365Integrity.

    Every file in the run folder is hashed and written to manifest.sha256 in the standard
    "<hash>  <relative path>" format, so it can be verified with sha256sum, Get-FileHash, or the
    supplied Verify-Export.ps1. A richer manifest.csv carries size, timestamps and row counts
    alongside the hashes.

    The manifest is generated last, after all reports have finished, so it covers everything
    written during the run including the transcript and the run log.

    What this does and does not establish
    -------------------------------------
    It establishes that the exported files have not been modified since the run completed, and
    it detects accidental corruption or deliberate editing of the evidence afterwards.

    It does not establish that the export is a complete or accurate reflection of the tenant.
    A run made with insufficient permissions produces an incomplete export whose hashes verify
    perfectly. That is why granted scopes, module versions and row counts are recorded in the
    run log, and why the completeness check runs alongside this. Anyone writing this up should
    describe the manifest as detecting post-export alteration, not as proving the evidence is
    correct.

    Note also that the manifest is not itself signed by this function. Anyone able to alter the
    exports could regenerate the manifest. To close that gap the manifest must be committed to
    somewhere the operator cannot alter it, or signed. See Docs\Verification.md.

    .PARAMETER Context
    The run context object.

    .EXAMPLE
    New-OagM365Manifest -Context $Context

    .NOTES
        NAME: New-OagM365Manifest
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context
    )

    Write-OagM365Log "Generating evidence manifest" -Context $Context

    try {
        $manifestTxt = Join-Path -Path $Context.RootPath -ChildPath 'manifest.sha256'
        $manifestCsv = Join-Path -Path $Context.LogPath  -ChildPath 'manifest.csv'

        # Exclude the manifest files themselves; a file cannot contain its own hash.
        $files = Get-ChildItem -Path $Context.RootPath -Recurse -File |
                 Where-Object { $_.FullName -ne $manifestTxt -and $_.FullName -ne $manifestCsv } |
                 Sort-Object FullName

        if (-not $files) {
            Write-OagM365Log "No output files found to hash. The run produced no evidence." -Level Warning -Indent 1 -Context $Context
            return $null
        }

        $rows = foreach ($file in $files) {
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256 -ErrorAction Stop
            $relative = $file.FullName.Substring($Context.RootPath.Length).TrimStart('\', '/')

            # Match against the expected-output register so row counts appear beside hashes.
            $expected = $Context.Expected | Where-Object { $_.Path -eq $file.FullName } | Select-Object -First 1

            [PSCustomObject]@{
                RelativePath  = $relative
                Sha256        = $hash.Hash
                SizeBytes     = $file.Length
                LastWriteTime = $file.LastWriteTimeUtc.ToString('o')
                Report        = if ($expected) { $expected.Report } else { '' }
                Description   = if ($expected) { $expected.Description } else { '' }
                RowCount      = if ($expected) { $expected.RowCount } else { '' }
            }
        }

        # sha256sum-compatible format: hash, two spaces, relative path.
        $lines = $rows | ForEach-Object { "$($_.Sha256)  $($_.RelativePath)" }
        $lines | Out-File -FilePath $manifestTxt -Encoding utf8 -ErrorAction Stop

        $rows | Export-Csv -Path $manifestCsv -NoTypeInformation -Encoding utf8 -ErrorAction Stop

        $totalMb = [math]::Round(($rows | Measure-Object -Property SizeBytes -Sum).Sum / 1MB, 2)
        Write-OagM365Log "Hashed $($rows.Count) file(s), $totalMb MB" -Level Success -Indent 1 -Context $Context
        Write-OagM365Log "Manifest: $manifestTxt" -Level Detail -Indent 1 -Context $Context

        return $rows

    } catch {
        $exception = Format-OagM365Exception -message "Failed generating evidence manifest" -exception $_
        Write-OagM365Log $exception -Level Error -Indent 1 -Context $Context
        return $null
    }
}
