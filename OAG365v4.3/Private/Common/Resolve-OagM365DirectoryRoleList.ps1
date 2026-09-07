function Resolve-OagM365DirectoryRoleList {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of directory roles

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of directory roles

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    Resolve-OagM365DirectoryRoleList -srcObject $policy.Conditions.Users.IncludeRoles

    .NOTES
        NAME: Resolve-OagM365DirectoryRoleList
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-OagM365DirectoryRoleTemplate
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((Resolve-OagM365DirectoryRoleTemplate -InputObject $PSItem)) | Out-Null
    }

    return $list
}
