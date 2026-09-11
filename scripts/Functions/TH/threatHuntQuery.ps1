function threatHuntQueryGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns the three advanced hunting query definitions used by the Defender report.

    In the original script there were no functions at all. The three queries ran from
    three near-identical try/catch blocks pasted inline, each building its own JSON body
    and calling Invoke-MgGraphRequest directly. Each catch assigned the formatted error
    to a variable and then never wrote it out, so a failed query produced no console
    output, no transcript entry and no file. The run appeared to succeed while silently
    producing nothing.

    The queries are now data returned by this function and executed by threatHuntQueryRun,
    so all three share one code path with real error handling.

    LOGIC
    Holds the KQL and the Select-Object property lists verbatim from the original, so
      results stay directly comparable with previous exports.
    Returns one object per query carrying the KQL, the property list, the timespan and
      the export target key.

    Queries are held in code rather than in separate .kql files on purpose. An external
    query file sits outside the signed catalog, which would reopen the tamper question
    the signing is meant to close.

    PARAMETERS
    None.

    RUNNING CONTEXT
    Called by  : threatHuntReportWrite
    Calls      : no other functions in this module
    Returns    : three definition objects, passed one at a time to threatHuntQueryRun

    CMLETS/PERMISSIONS/SCOPES
    None directly. The queries it returns need ThreatHunting.Read.All.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param()


    # --- Threat and vulnerability management -------------------------------------------
    $tvmQuery = '
    DeviceTvmSoftwareVulnerabilities
    | join kind=inner (DeviceTvmSoftwareInventory) on DeviceId, SoftwareName
    | project DeviceId, DeviceName, OSPlatform, OSVersion, OSArchitecture, SoftwareVendor, SoftwareName, SoftwareVersion, CveId, VulnerabilitySeverityLevel, CveTags, AadDeviceId, Type, EndOfSupportStatus, EndOfSupportDate, ProductCodeCpe
    | join DeviceTvmSoftwareVulnerabilitiesKB on CveId
    | project DeviceId, AadDeviceId, DeviceName, OSPlatform, OSVersion, OSArchitecture, SoftwareVendor, SoftwareName, SoftwareVersion, CveId, VulnerabilitySeverityLevel, CvssScore, CvssVector, CveSupportability, IsExploitAvailable, CveTags, CveLastModifiedTime = LastModifiedTime, CvePublishedDate = PublishedDate, VulnerabilityDescription, AffectedSoftware, Type, EndOfSupportStatus, EndOfSupportDate, ProductCodeCpe
    '

    $tvmSelect = @(
        "DeviceId", 
        "AadDeviceId", 
        "DeviceName", 
        "OSPlatform", 
        "OSVersion", 
        "OSArchitecture", 
        "SoftwareVendor", 
        "SoftwareName", 
        "SoftwareVersion", 
        "CveId", 
        "VulnerabilitySeverityLevel", 
        "CvssScore", 
        "CvssVector", 
        "CveSupportability", 
        "IsExploitAvailable", 
        @{ Name = 'CveTags'; Expression = { $_.CveTags -join ', ' } },
        "CveLastModifiedTime", 
        "CvePublishedDate", 
        "VulnerabilityDescription", 
        @{ Name = 'AffectedSoftware'; Expression = { $_.AffectedSoftware -join ', ' } },
        "Type", 
        "EndOfSupportStatus", 
        "EndOfSupportDate", 
        "ProductCodeCpe"
    )

    # --- Defender agent health ----------------------------------------------------------
    $healthQuery = '
    DeviceTvmInfoGathering 
    | extend AdditionalFields = parse_json(AdditionalFields)
    | extend AvMode = tostring(AdditionalFields.["AvMode"])
    | extend AvModeDataRefreshTime = tostring(AdditionalFields.["AvModeDataRefreshTime"])
    | extend AvEngineVersion = tostring(AdditionalFields.["AvEngineVersion"])
    | extend AvSignatureVersion = tostring(AdditionalFields.["AvSignatureVersion"])
    | extend AvPlatformVersion = tostring(AdditionalFields.["AvPlatformVersion"])
    | extend AvEngineUpdateTime = tostring(AdditionalFields.["AvEngineUpdateTime"])
    | extend AvSignatureUpdateTime = tostring(AdditionalFields.["AvSignatureUpdateTime"])
    | extend AvPlatformUpdateTime = tostring(AdditionalFields.["AvPlatformUpdateTime"])
    | extend AvIsEngineUptodate = tostring(AdditionalFields.["AvIsEngineUptodate"])
    | extend AvIsSignatureUptoDate = tostring(AdditionalFields.["AvIsSignatureUptoDate"])
    | extend AvIsPlatformUptodate = tostring(AdditionalFields.["AvIsPlatformUptodate"])
    | extend AvEnginePublishTime = tostring(AdditionalFields.["AvEnginePublishTime"])
    | extend AvSignaturePublishTime = tostring(AdditionalFields.["AvSignaturePublishTime"])
    | extend AvPlatformPublishTime = tostring(AdditionalFields.["AvPlatformPublishTime"])
    | extend AvSignatureDataRefreshTime = tostring(AdditionalFields.["AvSignatureDataRefreshTime"])
    | extend AvSignatureRing = tostring(parse_json(AdditionalFields).AvSignatureRing)
    | extend AvPlatformRing = tostring(parse_json(AdditionalFields).AvPlatformRing)
    | extend AvEngineRing = tostring(parse_json(AdditionalFields).AvEngineRing)
    | join DeviceTvmSecureConfigurationAssessment on DeviceId
    | where ConfigurationId == "scid-2011" 
    | project-away DeviceId1, DeviceName1, OSPlatform1, Context, ConfigurationImpact
    | project TimeGenerated, Timestamp, LastSeenTime, DeviceId, DeviceName, OSPlatform, Type, SourceSystem, MachineGroup, AvMode, AvModeDataRefreshTime, AvEngineVersion, AvSignatureVersion, AvPlatformVersion, AvEngineUpdateTime, AvSignatureUpdateTime, AvPlatformUpdateTime, AvIsEngineUptodate, AvIsSignatureUptoDate, AvIsPlatformUptodate, AvEnginePublishTime, AvSignaturePublishTime, AvPlatformPublishTime, AvSignatureDataRefreshTime, AvSignatureRing, AvPlatformRing, AvEngineRing, Type1, ConfigurationId, ConfigurationCategory, ConfigurationSubcategory, Timestamp1, IsCompliant, IsApplicable
    '

    $healthSelect = @(
        "TimeGenerated", 
        "Timestamp", 
        "LastSeenTime", 
        "DeviceId", 
        "DeviceName", 
        "OSPlatform", 
        "Type", 
        "SourceSystem", 
        "MachineGroup", 
        "AvMode", 
        "AvModeDataRefreshTime", 
        "AvEngineVersion", 
        "AvSignatureVersion", 
        "AvPlatformVersion", 
        "AvEngineUpdateTime", 
        "AvSignatureUpdateTime", 
        "AvPlatformUpdateTime", 
        "AvIsEngineUptodate", 
        "AvIsSignatureUptoDate", 
        "AvIsPlatformUptodate", 
        "AvEnginePublishTime", 
        "AvSignaturePublishTime", 
        "AvPlatformPublishTime", 
        "AvSignatureDataRefreshTime", 
        "AvSignatureRing", 
        "AvPlatformRing", 
        "AvEngineRing", 
        "Type1"
        "ConfigurationId", 
        "ConfigurationCategory", 
        "ConfigurationSubcategory", 
        "Timestamp1", 
        "IsCompliant", 
        "IsApplicable"
    )

    # --- Device software inventory ------------------------------------------------------
    $softwareQuery = 'DeviceTvmSoftwareInventory'

    # Timespan values are carried across from the original script: 30 days for the
    # vulnerability and software queries, 45 for agent health.
    return @(
        [PSCustomObject]@{
            key         = 'Tvm'
            description = 'Defender threat and vulnerability management report'
            query       = $tvmQuery
            select      = $tvmSelect
            timespan    = 'P30D'
            targetKey   = 'defenderTvmReport'
        }
        [PSCustomObject]@{
            key         = 'Health'
            description = 'Defender agent health report'
            query       = $healthQuery
            select      = $healthSelect
            timespan    = 'P45D'
            targetKey   = 'defenderHealthReport'
        }
        [PSCustomObject]@{
            key         = 'SoftwareInventory'
            description = 'Defender device software inventory'
            query       = $softwareQuery
            select      = @('*')
            timespan    = 'P30D'
            targetKey   = 'softwareInventory'
        }
    )
}

function threatHuntQueryRun {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Runs one Defender advanced hunting query and exports the results.

    Replaces the three duplicated inline blocks in the original with one code path.

    LOGIC
    Builds the JSON body and posts to v1.0/security/runHuntingQuery.
    Uses -ErrorAction Stop, because Graph errors surface as non-terminating by default,
      which means a plain try/catch does not catch them and execution carries on as if
      the call had succeeded.
    Selects the requested properties and writes through exportWrite so the row count is
      logged and a zero-row result is visible.
    On failure, detects a consent or permission error specifically and reports it as such.
      That is the usual cause of this report returning nothing and it needs to be
      distinguishable in the evidence from a genuine empty result.
    Registers the failure so the completeness check reports the file as missing rather
      than the run appearing clean.

    PARAMETERS
    -definition (required) one object from threatHuntQueryGet

    RUNNING CONTEXT
    Called by  : threatHuntReportWrite, once per query
    Calls      : logWrite, exportWrite, exceptionFormat
    Reads      : $script:exportTarget

    CMLETS/PERMISSIONS/SCOPES
    Invoke-MgGraphRequest (POST v1.0/security/runHuntingQuery)
    Scope: ThreatHunting.Read.All, which requires tenant administrator consent
    Role : Security Reader

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$definition
    )

    logWrite "$($definition.description) (timespan $($definition.timespan))"

    try {
        $body = @{ Query = $definition.query; Timespan = $definition.timespan } | ConvertTo-Json

        $response = Invoke-MgGraphRequest -Method POST `
                                          -Uri 'v1.0/security/runHuntingQuery' `
                                          -Body $body `
                                          -OutputType PSObject `
                                          -ErrorAction Stop

        $results = if ($definition.select -contains '*') {
            $response.results | Select-Object -Property *
        } else {
            $response.results | Select-Object -Property $definition.select
        }

        $results | exportWrite -path $script:exportTarget.($definition.targetKey) `
                               -description $definition.description
        return $true

    } catch {
        $message = $_.Exception.Message
        $status  = $null
        if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
            $status = [int]$_.Exception.Response.StatusCode
        } elseif ($message -match 'Unauthorized') {
            $status = 401
        } elseif ($message -match 'Forbidden') {
            $status = 403
        } elseif ($message -match 'BadRequest|Bad Request') {
            $status = 400
        }

        # The status line says only "Bad Request". Which table the service could not find
        # is in the response body, which the SDK leaves in ErrorDetails.
        $detail = "$($_.ErrorDetails) $message"
        $missingTable = if ($detail -match "named '([^']+)'") { $Matches[1] } else { $null }

        # Two different failures arrive here and they mean opposite things for the audit.
        #
        # 403 or a consent error is a permission gap. The tenant has the capability, this
        # account is not allowed to query it, and re-consenting fixes it.
        #
        # 401 with the scope already granted is not a permission gap. The hunting endpoint
        # answers 401 when the tenant has no Defender XDR workload behind it - not licensed,
        # or licensed but never onboarded, so there is no advanced hunting data to query.
        # Nothing the operator consents to will change that, and recording it as a
        # permission problem would put a false remediation in the working paper.
        $granted = $script:run.scopes |
                   Where-Object { $_.scope -eq 'ThreatHunting.Read.All' -and $_.granted }

        $reason = if ($status -eq 403 -or $message -match 'Forbidden|consent|Authorization_RequestDenied|insufficient privileges') {
            logWrite "PERMISSION DENIED for $($definition.description)." -level Error -indent 1
            logWrite "ThreatHunting.Read.All requires tenant administrator consent. Until that is granted this report cannot be produced, and the gap should be recorded in the working paper." -level Error -indent 2
            'Permission denied - ThreatHunting.Read.All not consented'
        } elseif ($status -eq 401 -and $granted) {
            logWrite "Advanced hunting is not available in this tenant." -level Warning -indent 1
            logWrite "ThreatHunting.Read.All was granted, but the hunting endpoint returned 401. That is the response when no Microsoft Defender XDR workload is onboarded, so there is no hunting data to query." -level Warning -indent 2
            logWrite "Record this as scope not applicable rather than a failed test." -level Warning -indent 2
            'Defender XDR not onboarded in this tenant'
        } elseif ($status -eq 401) {
            logWrite "UNAUTHORISED for $($definition.description). ThreatHunting.Read.All was requested but is not in the granted scope set." -level Error -indent 1
            'Unauthorised - ThreatHunting.Read.All not granted'
        } elseif ($status -eq 400 -and $detail -match 'Failed to resolve table') {
            # The query reached advanced hunting and advanced hunting understood it. The
            # table simply does not exist in this tenant, which is what happens when the
            # workload behind it was never onboarded. The DeviceTvm* tables come from
            # Defender for Endpoint and Defender Vulnerability Management, so a tenant with
            # no onboarded devices has no rows and no table to hold them.
            logWrite "Advanced hunting table $missingTable does not exist in this tenant." -level Warning -indent 1
            logWrite "The DeviceTvm tables come from Defender for Endpoint and Defender Vulnerability Management. With no onboarded devices there is nothing to query." -level Warning -indent 2
            logWrite "Record this as scope not applicable rather than a failed test." -level Warning -indent 2
            "Advanced hunting table $missingTable not available in this tenant"
        } else {
            logWrite (exceptionFormat -message "Failed running hunting query '$($definition.description)'" -exception $_) -level Error -indent 1
            "Query failed - $message"
        }

        # A tenant with no Defender workload behind the endpoint has no control to test, so
        # the absent file is not a gap in the evidence. A permission failure is.
        $notApplicable = ($status -eq 401 -and $granted) -or
                         ($status -eq 400 -and $detail -match 'Failed to resolve table')

        # Both classifications apply to the whole endpoint, not to this one query, so tell
        # the caller to stop rather than repeat the same failure for every remaining query.
        if ($status -eq 401 -or $status -eq 403) {
            $script:threatHuntUnavailable = [PSCustomObject]@{ reason = $reason; notApplicable = $notApplicable }
        }

        threatHuntQuerySkip -definition $definition -reason $reason -notApplicable:$notApplicable
        return $false
    }
}


function threatHuntQuerySkip {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Records a hunting query that produced no file, with the reason, so the completeness
    check can distinguish "not attempted because the capability is absent" from "failed".

    LOGIC
    Appends an expected-but-unwritten entry to the run register.

    PARAMETERS
    -definition (required) one object from threatHuntQueryGet
    -reason (required) why no file was produced
    -notApplicable (optional) the capability does not exist in this tenant, so the absent
      file is not a gap in the evidence and should not make the run read as incomplete

    RUNNING CONTEXT
    Called by  : threatHuntQueryRun, threatHuntReportWrite
    Reads      : $script:exportTarget, $script:run

    CMLETS/PERMISSIONS/SCOPES
    None. Local state only.

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$definition,
        [Parameter(Mandatory = $true)][string]$reason,
        [switch]$notApplicable
    )

    $script:run.expected += [PSCustomObject]@{
        report = $script:run.currentReport
        path = $script:exportTarget.($definition.targetKey)
        fileName = (Split-Path $script:exportTarget.($definition.targetKey) -Leaf)
        description = $definition.description
        rowCount = 0; written = $false
        reason = $reason; applicable = (-not $notApplicable)
        timestamp = (Get-Date).ToString('o')
    }
}
