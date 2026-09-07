function Format-OagM365Exception {
    <#
    .SYNOPSIS
    Exception message handling.

    .DESCRIPTION
    Creates and returns a formatted string containing an error message generated from the
    provided exception object.

    Previously duplicated verbatim in all five export scripts. Defined once here.

    .PARAMETER message
    Custom error message provided by the caller to give meaningful context.

    .PARAMETER exception
    Exception object as caught by try / catch.

    .EXAMPLE
    Format-OagM365Exception -message "Failed connecting Graph API to M365 tenant" -exception $_

    .NOTES
        NAME: Format-OagM365Exception
        VERSION: 2.0

        CHANGELOG:
          4.0.0: Moved to module scope. Behaviour unchanged from v1.0.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$message,
        [Parameter(Mandatory = $true)][AllowEmptyString()][object]$exception
    )

    $fmtString = "{0}`n`n{1} : {2}`n{3}`n" +
                 "      + CategoryInfo     : {4}`n" +
                 "      + FullyQualifiedId : {5}`n"

    $exceptionFields = $message,
                       $exception.InvocationInfo.InvocationName,
                       $exception.Exception.Message,
                       $exception.InvocationInfo.PositionMessage,
                       $exception.CategoryInfo.ToString(),
                       $exception.FullyQualifiedErrorId

    return $fmtString -f $exceptionFields
}
