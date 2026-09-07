function Export-OagM365IamReport {
    <#
    .SYNOPSIS
    Exports Entra identity and access management data: users, devices, service principals,
    authentication methods, administrative units and PIM role assignments.

    .DESCRIPTION
    The heaviest of the five reports. Runs in three stages, matching the original script's
    staging, so a partial failure is attributable to a stage rather than the whole report.

    .PARAMETER Context
    Run context object from Initialize-OagM365Run.

    .PARAMETER Connection
    Existing connection result from Connect-OagM365Graph.

    .EXAMPLE
    Export-OagM365IamReport -Context $Context -Connection $Connection

    .NOTES
        NAME: Export-OagM365IamReport
        VERSION: 4.0.0

        REQUIRED ENTRA ROLE:
           Global Reader plus Reports Reader (sign-in activity)

        GRAPH SCOPES:
           Directory.Read.All, DeviceManagementApps.Read.All, Device.Read.All,
           Application.Read.All, AuditLog.Read.All, RoleManagement.Read.Directory,
           PrivilegedEligibilitySchedule.Read.AzureADGroup,
           PrivilegedAssignmentSchedule.Read.AzureADGroup
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$Connection
    )

    $folder = Initialize-OagM365ReportPath -Context $Context -Report 'Iam'

    $status = [PSCustomObject]@{
        Report    = 'Iam'
        Started   = Get-Date
        Completed = $null
        Succeeded = $false
        Folder    = $folder
    }

    try {
        Write-OagM365Log "STAGE 1 of 3: users, devices and service principals" -Context $Context
        Export-OagM365IamUser | Out-Null
        Export-OagM365IamDevice | Out-Null
        Export-OagM365IamServicePrincipal | Out-Null

        Write-OagM365Log "STAGE 2 of 3: user authentication methods" -Context $Context
        Export-OagM365IamUserAuthMethod | Out-Null

        Write-OagM365Log "STAGE 3 of 3: administrative units and PIM role membership" -Context $Context
        Export-OagM365IamAdminUnit | Out-Null

        # Get-OagM365IamPimRole populates $script:pimMgmtPolicyRules as a side effect, so the
        # role export must run before the policy rules are written. This ordering dependency
        # is inherited from the original and is called out here because it is not obvious.
        $pimRoles = Get-OagM365IamPimRole
        $pimRoles | Write-OagM365Export -Path $script:exportTarget.builtInRoleMembers `
                                        -Context $Context `
                                        -Description 'Built-in role assignments'

        $script:pimMgmtPolicyRules | Write-OagM365Export -Path $script:exportTarget.pimMgmtPolicyRules `
                                                        -Context $Context `
                                                        -Description 'PIM role management policy rules'

        Register-OagM365ExistingOutput -Context $Context -Folder $folder -Filter '*.csv' -Description 'IAM CSV export'

        $status.Succeeded = $true

    } catch {
        $exception = Format-OagM365Exception -message "IAM report failed" -exception $_
        Write-OagM365Log $exception -Level Error -Context $Context
    } finally {
        $status.Completed = Get-Date
        $Context.ReportStatus.Add($status)
    }

    return $status
}
