function exportVerify {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Compares the exports that were expected against those actually produced, and writes
    completeness.csv to the run log.

    This exists because the highest risk failure for audit evidence is not a crash. A
    crash is visible and gets investigated. The dangerous case is the export that finishes,
    reports success, and quietly writes fewer rows than it should or none at all. The
    evidence looks complete, gets relied on, and the gap surfaces much later.

    LOGIC
    Walks the registered expected outputs and classifies each one:
      Missing        - never written, usually a function that threw and was caught
      Not applicable - never written because the capability is absent from the tenant,
                       for example advanced hunting in a tenant with no Defender XDR.
                       There is no control to evidence, so this is not a gap
      Empty          - written but holds no records, usually a permission or scope problem
      OK             - written with records

    Missing and Not applicable look identical on disk and mean opposite things in the
    working paper, so the caller records which one it is when it registers the export.
    Counts errors and warnings from the run event log.
    Produces an overall status of Complete, Complete with warnings, or Incomplete.

    That status belongs in the working paper. It is the honest statement of whether the
    evidence set is whole.

    PARAMETERS
    None. Reads $script:run.

    RUNNING CONTEXT
    Called by  : runMe.ps1 in its finally block
    Calls      : logWrite
    Returns    : summary object used by the run summary and by Verify-Export.ps1

    CMLETS/PERMISSIONS/SCOPES
    None. Local filesystem only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()

    logWrite "Checking export completeness"

    $rows = foreach ($item in $script:run.expected) {
        $exists = Test-Path $item.path
        $size   = if ($exists) { (Get-Item $item.path).Length } else { 0 }
        $notApplicable = ($null -ne $item.applicable -and -not $item.applicable)
        $state  = if (-not $item.written -or -not $exists) {
                      if ($notApplicable) { 'Not applicable' } else { 'Missing' }
                  }
                  elseif ($item.rowCount -eq 0)            { 'Empty' }
                  else                                     { 'OK' }

        [PSCustomObject]@{
            report = $item.report; fileName = $item.fileName; description = $item.description
            rowCount = $item.rowCount; sizeBytes = $size; status = $state
            reason = $item.reason
        }
    }

    $missing  = @($rows | Where-Object status -eq 'Missing')
    $notAppl  = @($rows | Where-Object status -eq 'Not applicable')
    $empty    = @($rows | Where-Object status -eq 'Empty')
    $ok       = @($rows | Where-Object status -eq 'OK')
    $errors   = @($script:run.events | Where-Object level -eq 'Error')
    $warnings = @($script:run.events | Where-Object level -eq 'Warning')

    $status = if ($missing.Count -gt 0 -or $errors.Count -gt 0) { 'Incomplete' }
              elseif ($empty.Count -gt 0 -or $warnings.Count -gt 0) { 'Complete with warnings' }
              else { 'Complete' }

    if ($rows.Count -gt 0) {
        $rows | Export-Csv (Join-Path $script:run.logPath 'completeness.csv') -NoTypeInformation -Encoding utf8
    }

    logWrite "$($ok.Count) of $($rows.Count) export(s) produced records" -level Detail -indent 1

    if ($missing.Count -gt 0) {
        logWrite "$($missing.Count) expected export(s) MISSING:" -level Error -indent 1
        $missing | ForEach-Object {
            $detail = if ($_.reason) { "$($_.description) - $($_.reason)" } else { $_.description }
            logWrite "$($_.report): $($_.fileName) - $detail" -level Error -indent 2
        }
    }
    if ($notAppl.Count -gt 0) {
        logWrite "$($notAppl.Count) export(s) not applicable to this tenant:" -level Warning -indent 1
        $notAppl | ForEach-Object { logWrite "$($_.report): $($_.fileName) - $($_.reason)" -level Warning -indent 2 }
    }
    if ($empty.Count -gt 0) {
        logWrite "$($empty.Count) export(s) contain no records:" -level Warning -indent 1
        $empty | ForEach-Object { logWrite "$($_.report): $($_.fileName) - $($_.description)" -level Warning -indent 2 }
    }

    return [PSCustomObject]@{
        runId = $script:run.id; status = $status
        filesExpected = $rows.Count; filesOk = $ok.Count
        filesEmpty = $empty.Count; filesMissing = $missing.Count
        filesNotApplicable = $notAppl.Count
        errorCount = $errors.Count; warningCount = $warnings.Count
    }
}
