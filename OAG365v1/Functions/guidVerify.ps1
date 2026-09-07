function guidVerify {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Tests whether a string is a valid GUID. Used before attempting to resolve an ID
    against Graph, so a malformed value fails here rather than as an opaque API error.

    LOGIC
    Sets up 1 lookup cache to avoid repeat Graph calls for the same ID.

    PARAMETERS
    -InputObject (required)

    RUNNING CONTEXT
    Called by  : adminUnitResolve, capNamedLocationResolve, capPolicyToHtml, directoryObjectListResolve, directoryObjectResolve, servicePrincipalResolve
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>

    [Cmdletbinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )
    process {
        return [guid]::TryParse($InputObject, $([ref][guid]::Empty))
    }
}

$script:cacheMgObject = @{}
