function Test-OagM365Guid {
    <#
    .SYNOPSIS
    Validates a given input string and checks string is a valid GUID

    .DESCRIPTION
    Validates a given input string and checks string is a valid GUID by using the .NET method Guid.TryParse

    .PARAMETER InputObject
    The GUID(s) to test

    .EXAMPLE
    Test-OagM365Guid -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Test-OagM365Guid
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Uses .NET method [guid]::TryParse()

        CHANGELOG:    
    #>
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
