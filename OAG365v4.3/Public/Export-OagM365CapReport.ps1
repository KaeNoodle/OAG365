function Export-OagM365CapReport {
    <#
    .SYNOPSIS
    Exports Microsoft Entra conditional access, authentication method and authentication
    strength policies.

    .DESCRIPTION
    Produces one HTML report per conditional access policy, a CSV policy summary, and CSV
    exports of the authentication method and authentication strength policies.

    Scopes are declared here, next to the code that needs them, rather than in a script-scope
    variable at the top of the file. In the original the declared list and the documented list
    were separate and had drifted apart.

    .PARAMETER Context
    Run context object from Initialize-OagM365Run.

    .PARAMETER Connection
    Existing connection result from Connect-OagM365Graph. Supplied by the dispatcher so multiple
    reports share one authentication.

    .EXAMPLE
    Export-OagM365CapReport -Context $Context -Connection $Connection

    .NOTES
        NAME: Export-OagM365CapReport
        VERSION: 4.0.0

        REQUIRED ENTRA ROLE:
           Global Reader (or Security Reader for the Defender portal elements)

        GRAPH SCOPES:
           Policy.Read.All, Policy.Read.ConditionalAccess, Policy.Read.AuthenticationMethod,
           AuthenticationContext.Read.All, RoleManagement.Read.Directory,
           Application.Read.All, Directory.Read.All
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$Connection
    )

    $folder = Initialize-OagM365ReportPath -Context $Context -Report 'Cap'

    $status = [PSCustomObject]@{
        Report    = 'Cap'
        Started   = Get-Date
        Completed = $null
        Succeeded = $false
        Folder    = $folder
    }

    try {
        Write-OagM365Log "Exporting conditional access policies" -Context $Context
        Export-OagM365CapPolicy -OutputFolder $folder | Out-Null

        Write-OagM365Log "Exporting authentication method policies" -Context $Context
        Export-OagM365CapAuthMethodPolicy | Out-Null

        Write-OagM365Log "Exporting authentication strength policies" -Context $Context
        Export-OagM365CapAuthStrengthPolicy | Out-Null

        # HTML reports are written directly by Convert-OagM365CapPolicyToHtml rather than
        # through Write-OagM365Export, so they are registered here to appear in the manifest
        # and the completeness check.
        Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter 'CAP_*.html' -Description 'Conditional access policy report'
        Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter '*.csv' -Description 'Conditional access CSV export'

        $status.Succeeded = $true

    } catch {
        $exception = Format-OagM365Exception -message "Conditional access report failed" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
    } finally {
        $status.Completed = Get-Date
        $Context.ReportStatus.Add($status)
    }

    return $status
}
