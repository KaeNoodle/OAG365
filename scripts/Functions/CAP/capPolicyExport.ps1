function capPolicyExport {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Obtain and iterate through all conditional access policies and build a HTML report
    for each

    LOGIC
    Runs 1 loop over the returned collection.
    Retrieves data using Get-MgIdentityConditionalAccessPolicy.
    Writes the result to CSV at the path held in $script:exportTarget.
    Writes an HTML file to the report output folder.
    Wrapped in 1 try/catch block; failures are formatted by exceptionFormat and
      written to the run log.

    PARAMETERS
    -OutputFolder (required)

    RUNNING CONTEXT
    Called by  : run directly, or by the report writer for this area
    Calls      : capPolicySummaryGet, capPolicyToHtml

    CMLETS/PERMISSIONS/SCOPES
    Get-MgIdentityConditionalAccessPolicy

    --------------------------------------------------------------------------------#>


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

        $capSummary += capPolicySummaryGet -policy $policy

        try {
            $capExportFileName = ($policy.DisplayName) -replace '[^a-zA-Z0-9-_ ]', ' '
            $outHtml = capPolicyToHtml -policy $policy 
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
