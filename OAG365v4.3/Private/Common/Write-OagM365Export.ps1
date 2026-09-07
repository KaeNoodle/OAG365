function Write-OagM365Export {
    <#
    .SYNOPSIS
    Writes an export file, records the row count, and registers it as an expected output.

    .DESCRIPTION
    A single writer for CSV and HTML output, replacing scattered Export-Csv and Out-File calls.

    The reason this exists is the silent-failure problem. In the original scripts a function
    that returned nothing still produced an empty CSV, and a function that returned a partial
    result produced a truncated one. Neither raised an error, so the run looked clean and the
    evidence gap was invisible until someone opened the file. Routing writes through one place
    means every file's row count is logged, and a zero-row result is reported as a warning at
    the time it happens.

    Registering the file on the context also lets Test-OagM365Completeness confirm at the end of
    the run that everything expected was produced.

    .PARAMETER InputObject
    The objects to export. Accepts pipeline input.

    .PARAMETER Path
    Destination file path.

    .PARAMETER Context
    Run context object.

    .PARAMETER Description
    Human-readable description of the export, used in log messages.

    .PARAMETER Format
    Csv (default) or Html.

    .PARAMETER AllowEmpty
    Suppress the zero-row warning where an empty result is a legitimate finding, for example a
    tenant with no named locations configured. The file is still written and still recorded, so
    an empty result remains visible as evidence rather than being hidden.

    .EXAMPLE
    $users | Write-OagM365Export -Path $script:exportTarget.users -Context $Context -Description "Entra users"

    .NOTES
        NAME: Write-OagM365Export
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        [object[]]$InputObject,

        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Context,

        [string]$Description,
        [ValidateSet('Csv', 'Html')][string]$Format = 'Csv',
        [switch]$AllowEmpty
    )

    begin {
        $collected = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if ($null -ne $InputObject) {
            foreach ($item in $InputObject) {
                if ($null -ne $item) { $collected.Add($item) }
            }
        }
    }

    end {
        $label = if ($Description) { $Description } else { Split-Path -Path $Path -Leaf }
        $count = $collected.Count

        try {
            $parent = Split-Path -Path $Path -Parent
            if ($parent -and -not (Test-Path $parent)) {
                New-Item -ItemType Directory -Force -Path $parent -ErrorAction Stop | Out-Null
            }

            if ($Format -eq 'Csv') {
                # An empty file is still written. A missing file and an empty file mean
                # different things, and the difference matters to a reviewer.
                if ($count -gt 0) {
                    $collected | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8 -ErrorAction Stop
                } else {
                    Set-Content -Path $Path -Value '' -Encoding utf8 -ErrorAction Stop
                }
            } else {
                $collected -join "`n" | Out-File -FilePath $Path -Encoding utf8 -ErrorAction Stop
            }

            $Context.Expected.Add([PSCustomObject]@{
                Report      = $Context.CurrentReport
                Path        = $Path
                FileName    = Split-Path -Path $Path -Leaf
                Description = $label
                RowCount    = $count
                Written     = $true
                Timestamp   = (Get-Date).ToString('o')
            })

            if ($count -eq 0 -and -not $AllowEmpty) {
                Write-OagM365Log "$label - 0 records written. Verify permissions and tenant configuration." -Level Warning -Indent 1 -Context $Context
            } else {
                Write-OagM365Log "$label - $count record(s) -> $(Split-Path -Path $Path -Leaf)" -Level Success -Indent 1 -Context $Context
            }

            return $true

        } catch {
            $Context.Expected.Add([PSCustomObject]@{
                Report      = $Context.CurrentReport
                Path        = $Path
                FileName    = Split-Path -Path $Path -Leaf
                Description = $label
                RowCount    = 0
                Written     = $false
                Timestamp   = (Get-Date).ToString('o')
            })
            $exception = Format-OagM365Exception -message "Failed writing export '$label' to $Path" -exception $_
            Write-OagM365Log $exception -Level Error -Indent 1 -Context $Context
            return $false
        }
    }
}
