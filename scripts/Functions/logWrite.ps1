function logWrite {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Writes a timestamped, colour-coded message to the console and records it in the
    run event log. Replaces the loose Write-Host calls in the original scripts.

    The event log matters for audit evidence. Warnings scroll past in a long run and
    are then invisible, so they are also written to events.csv where an incomplete
    export can be identified afterwards.

    LOGIC
    Picks a colour and indent prefix from the level, then writes to the console.
    If a run is active, appends the message to $script:run.events.
    If no run is active it still writes to the console, so a function called on its
      own works without any setup.

    PARAMETERS
    -message (required)
    -level (optional, one of: Info, Success, Warning, Error, Detail)
    -indent (optional, 0 to 3)

    RUNNING CONTEXT
    Called by  : effectively every function in the module
    Calls      : no other functions in this module
    Reads      : $script:run, if a run has been initialised

    CMLETS/PERMISSIONS/SCOPES
    None. Console output only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)][AllowEmptyString()][string]$message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Detail')]$level = 'Info',
        [ValidateRange(0, 3)][int]$indent = 0
    )

    $colour = switch ($level) {
        'Success' { 'Green' }
        'Warning' { 'DarkYellow' }
        'Error'   { 'Red' }
        'Detail'  { 'Gray' }
        default   { 'Cyan' }
    }

    $prefix = switch ($indent) {
        1 { '  - ' }
        2 { '      ' }
        3 { '        ' }
        default { '' }
    }

    Write-Host "$prefix$message" -ForegroundColor $colour

    # Only record if a run is active. Running a function on its own still logs to the
    # console, it just has nowhere to file the event.
    if ($null -ne $script:run) {
        $script:run.events += [PSCustomObject]@{
            timestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
            level     = $level
            report    = $script:run.currentReport
            message   = $message
        }
    }
}
