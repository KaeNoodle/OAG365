function Register-OagM365ExistingOutput {
    <#
    .SYNOPSIS
    Registers files that were written directly rather than through Write-OagM365Export.

    .DESCRIPTION
    Some ported functions write their own files. The conditional access report in particular
    writes one HTML file per policy from inside Convert-OagM365CapPolicyToHtml, and the ported
    CSV exports still call Export-Csv directly.

    Rewriting every one of those call sites was avoided deliberately. The original code is
    signed, reviewed and known to work, so a like-for-like move is a cheaper review and a lower
    risk of regression than several hundred edits. Instead the files those functions produce are
    swept up afterwards by this function, so they still appear in the evidence manifest and the
    completeness check.

    Files already registered are skipped, so calling this after Write-OagM365Export has run does
    not create duplicates.

    .PARAMETER Context
    Run context object.

    .PARAMETER Folder
    Folder to scan.

    .PARAMETER Filter
    Wildcard filter, for example 'CAP_*.html'.

    .PARAMETER Description
    Description applied to the files found.

    .EXAMPLE
    Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter '*.csv' -Description 'CA export'

    .NOTES
        NAME: Register-OagM365ExistingOutput
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Filter,
        [string]$Description = 'Export'
    )

    if (-not (Test-Path -Path $Folder)) { return }

    $known = $Context.Expected | ForEach-Object { $_.Path }
    $found = Get-ChildItem -Path $Folder -Filter $Filter -File -ErrorAction SilentlyContinue |
             Where-Object { $_.FullName -notin $known }

    foreach ($file in $found) {
        # Row count is unknown for files written outside Write-OagM365Export. A CSV is counted
        # by lines less the header so an empty export is still detected; anything else is
        # marked -1 to show the count was not measured rather than implying zero.
        $rowCount = if ($file.Extension -eq '.csv') {
            $lines = @(Get-Content -Path $file.FullName -ErrorAction SilentlyContinue)
            [math]::Max(0, $lines.Count - 1)
        } else {
            -1
        }

        $Context.Expected.Add([PSCustomObject]@{
            Report      = $Context.CurrentReport
            Path        = $file.FullName
            FileName    = $file.Name
            Description = $Description
            RowCount    = $rowCount
            Written     = $true
            Timestamp   = $file.LastWriteTimeUtc.ToString('o')
        })
    }

    if ($found) {
        Write-OagM365Log "Registered $($found.Count) file(s) matching '$Filter'" -Level Detail -Indent 1 -Context $Context
    }
}
