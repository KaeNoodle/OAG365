function Test-OagM365Completeness {
    <#
    .SYNOPSIS
    Compares the exports that were expected against those actually produced.

    .DESCRIPTION
    Runs at the end of every execution and writes completeness.csv to the run log.

    This exists because the highest-risk failure mode for audit evidence is not a crash. A crash
    is visible and gets investigated. The dangerous case is the export that finishes, reports
    success, and quietly writes a file with fewer rows than it should have or none at all. The
    evidence looks complete, gets relied on, and the gap is only discovered later.

    Three conditions are flagged:

      Missing  - an expected file was never written, usually a function that threw and was caught
      Empty    - the file exists but holds no records, usually a permission or scope problem
      Warnings - errors and warnings raised during the run, carried through from the event log

    The result is a status of Complete, Incomplete or Failed, which is printed in the run summary
    and stored with the evidence. That status belongs in the working paper: it is the honest
    statement of whether the evidence set is whole.

    .PARAMETER Context
    The run context object.

    .EXAMPLE
    $completeness = Test-OagM365Completeness -Context $Context

    .NOTES
        NAME: Test-OagM365Completeness
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context
    )

    Write-OagM365Log "Checking export completeness" -Context $Context

    $rows = foreach ($item in $Context.Expected) {
        $exists = Test-Path -Path $item.Path
        $size   = if ($exists) { (Get-Item -Path $item.Path).Length } else { 0 }

        $status = if (-not $item.Written -or -not $exists) { 'Missing' }
                  elseif ($item.RowCount -eq 0)            { 'Empty' }
                  else                                     { 'OK' }

        [PSCustomObject]@{
            Report      = $item.Report
            FileName    = $item.FileName
            Description = $item.Description
            RowCount    = $item.RowCount
            SizeBytes   = $size
            Status      = $status
        }
    }

    $missing  = @($rows | Where-Object Status -eq 'Missing')
    $empty    = @($rows | Where-Object Status -eq 'Empty')
    $ok       = @($rows | Where-Object Status -eq 'OK')
    $errors   = @($Context.Events | Where-Object Level -eq 'Error')
    $warnings = @($Context.Events | Where-Object Level -eq 'Warning')

    $status = if ($missing.Count -gt 0 -or $errors.Count -gt 0) { 'Incomplete' }
              elseif ($empty.Count -gt 0 -or $warnings.Count -gt 0) { 'Complete with warnings' }
              else { 'Complete' }

    $result = [PSCustomObject]@{
        RunId         = $Context.RunId
        Status        = $status
        FilesExpected = $rows.Count
        FilesOk       = $ok.Count
        FilesEmpty    = $empty.Count
        FilesMissing  = $missing.Count
        ErrorCount    = $errors.Count
        WarningCount  = $warnings.Count
    }

    if ($rows.Count -gt 0) {
        $rows | Export-Csv -Path (Join-Path $Context.LogPath 'completeness.csv') -NoTypeInformation -Encoding utf8
    }

    Write-OagM365Log "$($ok.Count) of $($rows.Count) export(s) produced records" -Level Detail -Indent 1 -Context $Context

    if ($missing.Count -gt 0) {
        Write-OagM365Log "$($missing.Count) expected export(s) MISSING:" -Level Error -Indent 1 -Context $Context
        $missing | ForEach-Object {
            Write-OagM365Log "$($_.Report): $($_.FileName) - $($_.Description)" -Level Error -Indent 2 -Context $Context
        }
    }

    if ($empty.Count -gt 0) {
        Write-OagM365Log "$($empty.Count) export(s) contain no records:" -Level Warning -Indent 1 -Context $Context
        $empty | ForEach-Object {
            Write-OagM365Log "$($_.Report): $($_.FileName) - $($_.Description)" -Level Warning -Indent 2 -Context $Context
        }
    }

    return $result
}
