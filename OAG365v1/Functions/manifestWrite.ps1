function manifestWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Hashes every file in the run folder with SHA-256 and writes manifest.sha256.

    This is the evidence half of the tamper question. The code half is codeVerify.

    What it establishes: the exported files have not been modified since the run finished.
    What it does not establish: that the export was complete or accurate. A run made with
    insufficient permissions produces an incomplete evidence set whose hashes verify
    perfectly. That is why granted scopes and row counts are recorded alongside, and why
    anyone writing this up should describe the manifest as detecting later alteration
    rather than as proving the evidence correct.

    Also worth stating plainly: whoever can alter the exports can regenerate the manifest.
    This defends against later or third party alteration, not against the operator.

    LOGIC
    Walks the run folder, excluding the two manifest files since neither can contain its
      own hash.
    Hashes each file and matches it against the registered expected outputs so row counts
      appear beside hashes.
    Writes manifest.sha256 in sha256sum format, and manifest.csv with the detail.

    PARAMETERS
    None. Reads $script:run.

    RUNNING CONTEXT
    Called by  : OAG-MainRunFile.ps1, after the transcript is closed
    Calls      : logWrite, exceptionFormat
    Verified by: Tools\Verify-Export.ps1

    CMLETS/PERMISSIONS/SCOPES
    Get-FileHash

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    logWrite "Generating evidence manifest"

    try {
        $manifestTxt = Join-Path $script:run.rootPath 'manifest.sha256'
        $manifestCsv = Join-Path $script:run.logPath  'manifest.csv'

        $files = Get-ChildItem -Path $script:run.rootPath -Recurse -File |
                 Where-Object { $_.FullName -ne $manifestTxt -and $_.FullName -ne $manifestCsv } |
                 Sort-Object FullName

        if (-not $files) {
            logWrite "No output files found to hash. The run produced no evidence." -level Warning -indent 1
            return $null
        }

        $rows = foreach ($file in $files) {
            $hash     = Get-FileHash -Path $file.FullName -Algorithm SHA256 -ErrorAction Stop
            $relative = $file.FullName.Substring($script:run.rootPath.Length).TrimStart('\', '/')
            $expected = $script:run.expected | Where-Object { $_.path -eq $file.FullName } | Select-Object -First 1

            [PSCustomObject]@{
                relativePath  = $relative
                sha256        = $hash.Hash
                sizeBytes     = $file.Length
                lastWriteTime = $file.LastWriteTimeUtc.ToString('o')
                report        = if ($expected) { $expected.report } else { '' }
                description   = if ($expected) { $expected.description } else { '' }
                rowCount      = if ($expected) { $expected.rowCount } else { '' }
            }
        }

        # sha256sum format: hash, two spaces, relative path.
        $rows | ForEach-Object { "$($_.sha256)  $($_.relativePath)" } |
            Out-File -FilePath $manifestTxt -Encoding utf8 -ErrorAction Stop

        $rows | Export-Csv -Path $manifestCsv -NoTypeInformation -Encoding utf8 -ErrorAction Stop

        $mb = [math]::Round(($rows | Measure-Object -Property sizeBytes -Sum).Sum / 1MB, 2)
        logWrite "Hashed $($rows.Count) file(s), $mb MB" -level Success -indent 1
        logWrite "Manifest: $manifestTxt" -level Detail -indent 1

        return $rows

    } catch {
        logWrite (exceptionFormat -message "Failed generating evidence manifest" -exception $_) -level Error -indent 1
        return $null
    }
}
