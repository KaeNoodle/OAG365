function Export-OagM365CapPolicy {
    <#
    .SYNOPSIS
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .DESCRIPTION
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .PARAMETER OutputFolder
    The folder to create the individual HTML files in. Defined in the $script:exportTargetFolder variable.

    .EXAMPLE
    Export-OagM365CapPolicy -OutputFolder $OutputFolder

    .NOTES
        NAME: Export-OagM365CapPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityConditionalAccessPolicy
           Convert-OagM365CapPolicyToHtml
           
        CHANGELOG:
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$OutputFolder
    )
    Write-Host "`nEntra ID conditional access policies"

    Write-Progress -PercentComplete -1 -Activity "Fetching conditional access policies and related data from Graph API"

    # Get Conditional Access Policies
    $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -ExpandProperty "*" -All -ErrorAction Stop
    Write-Host "Found $($conditionalAccessPolicies.Count) conditional access policies"

    Write-Host "Processing policy..."

    $capSummary = @()

    # Process all Conditional Access Policies
    foreach ($policy in $conditionalAccessPolicies) {

        # Display some progress (based on policy count)
        $currentIndex = $conditionalAccessPolicies.indexOf($policy) + 1

        $progress = @{
            Activity         = "Generating Conditional Access Documentation..."
            PercentComplete  = [Decimal]::Divide($currentIndex, $conditionalAccessPolicies.Count) * 100
            CurrentOperation = "Processing Policy `"$($policy.DisplayName)`""
        }
        if ($currentIndex -eq $conditionalAccessPolicies.Count) { $progress.Add("Completed", $true) }

        Write-Progress @progress

        Write-Host "  - $($policy.DisplayName)"

        $capSummary += Get-OagM365CapPolicySummary -policy $policy

        try {
            $capExportFileName = ($policy.DisplayName) -replace '[^a-zA-Z0-9-_ ]', ' '
            $outHtml = Convert-OagM365CapPolicyToHtml -policy $policy 
            $outHtml | Out-File (Join-Path -Path $script:exportTargetFolder.conditionalAccess -ChildPath "CAP_$($capExportFileName).html")
            Write-Host "    Conditional access policy written to 'CAP_$($capExportFileName).html'" -ForegroundColor Green
        } catch {
            #Throw $_
            #Write-Error $PSItem
            #$_ | Format-List * -Force | Out-String
            $_ | Format-List * -Force | Out-String ; $_.InvocationInfo | Format-List * -Force | Out-String
        }
    }

    $capSummary | Export-Csv $script:exportTarget.capSummary -NoTypeInformation
    Write-Host "`nConditional access policy summary written to '$($script:exportTarget.capSummary)'" -ForegroundColor Green

}
