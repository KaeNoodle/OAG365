function exceptionFormat {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Builds a readable error string from a caught exception, with the caller's own
    message on the front so the reader knows what was being attempted.

    LOGIC
    Pulls the invocation name, message, position, category and fully qualified error
      ID out of the exception object.
    Formats them into a single multi-line string and returns it.

    PARAMETERS
    -message (required) plain description of what was being attempted
    -exception (required) the $_ object caught by try/catch

    RUNNING CONTEXT
    Called by  : every function with a try/catch block
    Calls      : no other functions in this module
    Returns    : a string, which the caller passes to logWrite at Error level

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$message,
        [Parameter(Mandatory = $true)][AllowEmptyString()]$exception
    )

    $format = "{0}`n`n{1} : {2}`n{3}`n" +
              "      + CategoryInfo     : {4}`n" +
              "      + FullyQualifiedId : {5}`n"

    return $format -f $message,
                      $exception.InvocationInfo.InvocationName,
                      $exception.Exception.Message,
                      $exception.InvocationInfo.PositionMessage,
                      $exception.CategoryInfo.ToString(),
                      $exception.FullyQualifiedErrorId
}
