function Invoke-OagM365HuntingQuery {
    <#
    .SYNOPSIS
    Runs a Microsoft Defender advanced hunting query and exports the results.

    .DESCRIPTION
    Replaces the three near-identical inline try/catch blocks in the original ThreatHunting
    script with one code path.

    Changes from the original beyond deduplication:

      - Errors are actually reported. The original assigned the formatted exception to a
        variable in each catch and never wrote it, so a failed query was silent.
      - -ErrorAction Stop is set on Invoke-MgGraphRequest. Graph errors surface as
        non-terminating by default, which means a plain try/catch does not catch them and
        execution continues as if the call succeeded.
      - The consent failure on ThreatHunting.Read.All is detected and reported as a permission
        problem rather than a generic error, because that is the usual cause of this report
        returning nothing and it needs to be distinguishable in the evidence.
      - Results go through Write-OagM365Export so row counts are logged and a zero-row result
        is visible.

    .PARAMETER Definition
    A query definition object from Get-OagM365ThreatHuntQueryDefinition.

    .PARAMETER Context
    Run context object.

    .EXAMPLE
    Get-OagM365ThreatHuntQueryDefinition | ForEach-Object { Invoke-OagM365HuntingQuery -Definition $_ -Context $Context }

    .NOTES
        NAME: Invoke-OagM365HuntingQuery
        VERSION: 1.0

        GRAPH CMDLETS:
           Invoke-MgGraphRequest  (POST v1.0/security/runHuntingQuery)

        GRAPH SCOPES: ThreatHunting.Read.All
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object]$Definition,
        [Parameter(Mandatory = $true)][object]$Context
    )

    Write-OagM365Log "$($Definition.Description) (timespan $($Definition.Timespan))" -Context $Context

    try {
        $jsonBody = @{
            Query    = $Definition.Query
            Timespan = $Definition.Timespan
        } | ConvertTo-Json

        $response = Invoke-MgGraphRequest -Method POST `
                                          -Uri 'v1.0/security/runHuntingQuery' `
                                          -Body $jsonBody `
                                          -OutputType PSObject `
                                          -ErrorAction Stop

        $results = if ($Definition.Select -contains '*') {
            $response.results | Select-Object -Property *
        } else {
            $response.results | Select-Object -Property $Definition.Select
        }

        $targetPath = $script:exportTarget.($Definition.TargetKey)

        $results | Write-OagM365Export -Path $targetPath `
                                       -Context $Context `
                                       -Description $Definition.Description

        return $true

    } catch {
        $message = $_.Exception.Message

        # A consent or permission failure is the common cause here and should not be buried in
        # a generic error. ThreatHunting.Read.All requires tenant admin consent, and without it
        # this report returns nothing however correct the rest of the run is.
        if ($message -match 'Forbidden|403|consent|Authorization_RequestDenied|insufficient privileges') {
            Write-OagM365Log "PERMISSION DENIED for $($Definition.Description)." -Level Error -Indent 1 -Context $Context
            Write-OagM365Log "The ThreatHunting.Read.All scope requires tenant administrator consent. Until consent is granted this report cannot be produced and the gap should be recorded in the working paper." -Level Error -Indent 2 -Context $Context
        } else {
            $exception = Format-OagM365Exception -message "Failed running hunting query '$($Definition.Description)'" -exception $_
            Write-OagM365Log $exception -Level Error -Indent 1 -Context $Context
        }

        # Register the failure so the completeness check reports the file as missing rather
        # than the run appearing clean.
        $Context.Expected.Add([PSCustomObject]@{
            Report      = $Context.CurrentReport
            Path        = $script:exportTarget.($Definition.TargetKey)
            FileName    = Split-Path -Path $script:exportTarget.($Definition.TargetKey) -Leaf
            Description = $Definition.Description
            RowCount    = 0
            Written     = $false
            Timestamp   = (Get-Date).ToString('o')
        })

        return $false
    }
}
