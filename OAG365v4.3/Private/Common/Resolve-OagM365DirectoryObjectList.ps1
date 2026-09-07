function Resolve-OagM365DirectoryObjectList {
    <#
    .SYNOPSIS
    Outputs a list of objects that can be resolved by Resolve-OagM365DirectoryObject (user, group, etc.) based on their GUID from a conditional access policy

    .DESCRIPTION
    Outputs a list of objects that can be resolved by Resolve-OagM365DirectoryObject (user, applications, etc.) based on their GUID from a conditional access policy

    .PARAMETER srcObject
    Iterable source object containing GUIDs. Uses ForEach-Object pipeline

    .PARAMETER getAttrib
    Object attribute to return in list (e.g. displayName or userPrinicpalName)

    .EXAMPLE
    Resolve-OagM365DirectoryObjectList -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"

    .NOTES
        NAME: Resolve-OagM365DirectoryObjectList
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Resolve-OagM365DirectoryObject
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$srcObject,
        [Parameter(Mandatory = $true)][ValidateSet('displayName','userPrincipalName')]$getAttrib
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        if (Test-OagM365Guid -InputObject $PSItem) {
            $object = Resolve-OagM365DirectoryObject -InputObject $PSItem
            switch ($getAttrib) {
                "userPrincipalName" {
                    $list.Add(($object.AdditionalProperties.userPrincipalName)) | Out-Null                    
                }
                "displayName" {
                    $list.Add(($object.AdditionalProperties.displayName)) | Out-Null                    
                }
            }
        } else {
            Write-Verbose " - $($PSItem)"
            $list.Add($PSItem) | Out-Null
        }
    }

    return $list
}
