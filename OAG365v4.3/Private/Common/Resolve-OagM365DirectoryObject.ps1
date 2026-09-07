function Resolve-OagM365DirectoryObject {
    <#
    .SYNOPSIS
    Resolve a Microsoft Graph item GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Graph item GUID to an object containing Get-MgDirectoryObject properties including user principal name, display name and app display name

    .PARAMETER InputObject
    The GUID(s) of the directory object to resolve

    .EXAMPLE
    Resolve-OagM365DirectoryObject -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-OagM365DirectoryObject
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryObject
           Get-MgServicePrincipal

        CHANGELOG:    
           2025-11-17 Added Get-MgServicePrincipal to handle scenario of build in Applications not resolving (e.g. SharePoint / OneDrive)
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-OagM365Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheMgObject.ContainsKey($InputObject)) {
                    Write-Debug "Cached display name for `"$InputObject`""
                    Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    return $script:cacheMgObject[$InputObject]
                } else {
                    try {
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $InputObject -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]
                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    } catch {
                        $exception = Format-OagM365Exception -message "Unable to find object using Get_MgDirectoryObject using provided GUID" -exception $_
                        Write-Host $exception -ForegroundColor DarkYellow

                        $directoryObjectId = (Get-MgServicePrincipal -filter "appId eq '$($InputObject)'").id
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $directoryObjectId -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]

                        Write-Host " - Don't panic. I found the service principal '$($displayName)' using the provided string as an AppID instead."

                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup from ServicePrincipal AppId: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    }
                    return $directoryObject
                }
            }
            catch {
                Write-Warning "Unable to resolve directory object with ID $InputObject, might have been deleted!"
            }
        }
        return $InputObject
    }
}


$script:cacheNamedLocations = @{}
