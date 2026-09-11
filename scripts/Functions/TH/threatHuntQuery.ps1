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

        if ($message -match 'Forbidden|403|consent|Authorization_RequestDenied|insufficient privileges') {
            logWrite "PERMISSION DENIED for $($definition.description)." -level Error -indent 1
            logWrite "ThreatHunting.Read.All requires tenant administrator consent. Until that is granted this report cannot be produced, and the gap should be recorded in the working paper." -level Error -indent 2
        } else {
            logWrite (exceptionFormat -message "Failed running hunting query '$($definition.description)'" -exception $_) -level Error -indent 1
        }

        $script:run.expected += [PSCustomObject]@{
            report = $script:run.currentReport
            path = $script:exportTarget.($definition.targetKey)
            fileName = (Split-Path $script:exportTarget.($definition.targetKey) -Leaf)
            description = $definition.description
            rowCount = 0; written = $false
            timestamp = (Get-Date).ToString('o')
        }
        return $false
    }
}
