function Resolve-OagM365ServicePrincipalList {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of service principal objects

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of service principal objects

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    Resolve-OagM365ServicePrincipalList -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals

    .NOTES
        NAME: Resolve-OagM365ServicePrincipalList
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-OagM365ServicePrincipal
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((Resolve-OagM365ServicePrincipal -InputObject $PSItem)) | Out-Null
    }

    return $list
}
