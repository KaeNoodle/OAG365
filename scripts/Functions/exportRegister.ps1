function exportRegister {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Registers files that were written directly rather than through exportWrite, so they
    still appear in the manifest and the completeness check.

    Some ported functions write their own files. The conditional access report writes one
    HTML file per policy from inside capPolicyToHtml, and the ported CSV exports still call
    Export-Csv directly. Rewriting every one of those call sites was avoided deliberately:
    the original code is signed, reviewed and known to work, so a like-for-like move is a
    cheaper review than several hundred edits. The files are swept up here instead.

    LOGIC
    Skips files already registered, so calling this after exportWrite creates no duplicates.
    Counts CSV rows by line count less the header, so an empty export is still detected.
    Marks non-CSV files with a row count of -1, meaning not measured rather than zero.

    PARAMETERS
    -folder (required) folder to scan
    -filter (required) wildcard, for example CAP_*.html
    -description (optional) applied to the files found

    RUNNING CONTEXT
    Called by  : capReportWrite, iamReportWrite, orgReportWrite, dfoReportWrite
    Calls      : logWrite

    CMLETS/PERMISSIONS/SCOPES
    None. Local filesystem only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$folder,
        [Parameter(Mandatory = $true)][string]$filter,
        [string]$description = 'Export'
    )

    if (-not (Test-Path $folder)) { return }

    $known = $script:run.expected | ForEach-Object { $_.path }
    $found = Get-ChildItem -Path $folder -Filter $filter -File -ErrorAction SilentlyContinue |
             Where-Object { $_.FullName -notin $known }

    foreach ($file in $found) {
        $rows = if ($file.Extension -eq '.csv') {
            [math]::Max(0, @(Get-Content $file.FullName -ErrorAction SilentlyContinue).Count - 1)
        } else { -1 }

        $script:run.expected += [PSCustomObject]@{
            report = $script:run.currentReport; path = $file.FullName
            fileName = $file.Name; description = $description
            rowCount = $rows; written = $true
            timestamp = $file.LastWriteTimeUtc.ToString('o')
        }
    }

    if ($found) { logWrite "Registered $($found.Count) file(s) matching '$filter'" -level Detail -indent 1 }
}
