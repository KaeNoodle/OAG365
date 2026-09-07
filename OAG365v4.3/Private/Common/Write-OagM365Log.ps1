function Write-OagM365Log {
    <#
    .SYNOPSIS
    Writes a consistently formatted message to the host and the run event log.

    .DESCRIPTION
    The original scripts used bare Write-Host calls with inconsistent indentation and colour.
    This wraps that so every message is timestamped, levelled, and (when a run context is
    available) appended to the machine-readable event log written to 00_RunLog\events.csv.

    The event log matters for audit evidence: it records warnings and errors that scroll past
    in the console, so an incomplete export can be identified after the fact.

    .PARAMETER Message
    Text to write.

    .PARAMETER Level
    Info, Success, Warning, Error or Detail. Controls colour and how the event is recorded.

    .PARAMETER Indent
    Indentation depth, 0-3. Purely cosmetic.

    .PARAMETER Context
    Optional run context object. When supplied the message is appended to $Context.Events.

    .EXAMPLE
    Write-OagM365Log -Message "Exported 42 users" -Level Success -Indent 1 -Context $Context

    .NOTES
        NAME: Write-OagM365Log
        VERSION: 1.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)][AllowEmptyString()][string]$Message,

        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Detail')]
        [string]$Level = 'Info',

        [ValidateRange(0, 3)]
        [int]$Indent = 0,

        [object]$Context
    )

    $colour = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'DarkYellow' }
        'Error'   { 'Red' }
        'Detail'  { 'Gray' }
        default   { 'Cyan' }
    }

    $prefix = switch ($Indent) {
        1 { '  - ' }
        2 { '      ' }
        3 { '        ' }
        default { '' }
    }

    Write-Host "$prefix$Message" -ForegroundColor $colour

    # $null must be on the LEFT of the comparison. With a collection on the left, PowerShell
    # treats -ne as a filter rather than a boolean test: an empty list filtered against $null
    # returns an empty result, which evaluates as false. Written the other way round, the
    # condition was false whenever the list was empty, so the first event was never added and
    # the log stayed empty for the whole run - silently discarding every warning and error.
    if ($Context -and $null -ne $Context.Events) {
        $Context.Events.Add([PSCustomObject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
            Level     = $Level
            Report    = $Context.CurrentReport
            Message   = $Message
        })
    }
}
