function Resolve-OagM365DirectoryRoleTemplate {
    <#
    .SYNOPSIS
    Resolve a directory role template 

    .DESCRIPTION
    Resolve a directory role template Id to a display name

    .PARAMETER InputObject
    The Id of the directory role template to resolve

    .EXAMPLE
    Resolve-OagM365DirectoryRoleTemplate -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-OagM365DirectoryRoleTemplate
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryRoleTemplate

        CHANGELOG:
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        try {
            # use hashtable as cache to limit API calls
            if (($script:cacheDirectoryRoleTemplates.ContainsKey($InputObject) -eq $false)) {
                $script:cacheDirectoryRoleTemplates = Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Group-Object -Property Id -AsHashTable
                Write-Host "Found $($script:cacheDirectoryRoleTemplates.Count) directory role templates, adding to cache"
            }

            Write-Verbose "Loading cached directory role template '$InputObject' $(($script:cacheDirectoryRoleTemplates[$InputObject]).DisplayName)"
            return $script:cacheDirectoryRoleTemplates[$InputObject].DisplayName
        } catch {
            Write-Warning "Unable to resolve directory role template with ID '$InputObject'"
        }

        return $InputObject
    }
}
