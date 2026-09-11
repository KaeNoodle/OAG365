function exportWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Writes an export file, logs how many records it contained, and registers it so the
    completeness check can confirm it was produced.

    This exists because of the silent failure problem. In the original scripts a function
    that returned nothing still wrote an empty CSV, and one that returned a partial result
    wrote a truncated one. Neither raised an error, so the run looked clean and the gap
    was invisible until someone opened the file. Routing writes through one place means
    every row count is logged and a zero-row result is reported when it happens.

    An empty file is still written. A missing file and an empty file mean different things
    to a reviewer, and collapsing that distinction loses information.

    LOGIC
    Collects pipeline input into one list.
    Creates the parent folder if needed.
    Writes CSV or HTML depending on the format requested.
    Registers the file on $script:run.expected with its row count.
    Warns if zero records were written, unless the caller says an empty result is valid.

    PARAMETERS
    -inputObject (pipeline) the objects to export
    -path (required) destination file
    -description (optional) used in the log message
    -format (optional, one of: Csv, Html)
    -allowEmpty (optional) suppress the zero-row warning where empty is a real finding

    RUNNING CONTEXT
    Called by  : threatHuntQueryRun, iamReportWrite, and new code generally
    Calls      : logWrite, exceptionFormat
    Note       : ported export functions still call Export-Csv directly and are picked
                 up afterwards by exportRegister

    CMLETS/PERMISSIONS/SCOPES
    None. Local filesystem only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $true)][AllowNull()]$inputObject,
        [Parameter(Mandatory = $true)][string]$path,
        [string]$description,
        [ValidateSet('Csv', 'Html')][string]$format = 'Csv',
        [switch]$allowEmpty
    )

    begin { $collected = @() }

    process {
        if ($null -ne $inputObject) {
            foreach ($item in $inputObject) { if ($null -ne $item) { $collected += $item } }
        }
    }

    end {
        $label = if ($description) { $description } else { Split-Path $path -Leaf }
        $count = $collected.Count

        try {
            $parent = Split-Path $path -Parent
            if ($parent -and -not (Test-Path $parent)) {
                New-Item -ItemType Directory -Force -Path $parent -ErrorAction Stop | Out-Null
            }

            if ($format -eq 'Csv') {
                if ($count -gt 0) {
                    $collected | Export-Csv -Path $path -NoTypeInformation -Encoding utf8 -ErrorAction Stop
                } else {
                    Set-Content -Path $path -Value '' -Encoding utf8 -ErrorAction Stop
                }
            } else {
                $collected -join "`n" | Out-File -FilePath $path -Encoding utf8 -ErrorAction Stop
            }

            $script:run.expected += [PSCustomObject]@{
                report = $script:run.currentReport; path = $path
                fileName = (Split-Path $path -Leaf); description = $label
                rowCount = $count; written = $true
                timestamp = (Get-Date).ToString('o')
            }

            if ($count -eq 0 -and -not $allowEmpty) {
                logWrite "$label - 0 records written. Check permissions and tenant configuration." -level Warning -indent 1
            } else {
                logWrite "$label - $count record(s) -> $(Split-Path $path -Leaf)" -level Success -indent 1
            }
            return $true

        } catch {
            $script:run.expected += [PSCustomObject]@{
                report = $script:run.currentReport; path = $path
                fileName = (Split-Path $path -Leaf); description = $label
                rowCount = 0; written = $false
                timestamp = (Get-Date).ToString('o')
            }
            logWrite (exceptionFormat -message "Failed writing '$label' to $path" -exception $_) -level Error -indent 1
            return $false
        }
    }
}
