function Get-OagM365ThreatHuntQueryDefinition {
    <#
    .SYNOPSIS
    Returns the advanced hunting query definitions used by the Defender report.

    .DESCRIPTION
    In the original ThreatHunting script there were no functions at all. The three hunting
    queries were run by three near-identical try/catch blocks pasted inline in the process{}
    block, each building its own JSON body and calling Invoke-MgGraphRequest directly.

    Worse, each catch block assigned the formatted error to a variable and then never wrote it
    out, so a failed query produced no console output, no transcript entry and no file. The run
    appeared to succeed while silently producing nothing.

    The queries are now data returned by this function and executed by
    Invoke-OagM365HuntingQuery, so all three share one code path with real error handling.

    The KQL and the Select-Object property lists are carried across verbatim from the original
    so results are directly comparable with previous exports.

    Queries are held in code rather than in separate .kql files on purpose: an external query
    file sits outside the signed catalog, which would reopen the tamper question the signing is
    meant to close.

    .EXAMPLE
    Get-OagM365ThreatHuntQueryDefinition | Where-Object Key -eq 'Tvm'

    .NOTES
        NAME: Get-OagM365ThreatHuntQueryDefinition
        VERSION: 1.0

        GRAPH SCOPES: ThreatHunting.Read.All
    #>
    [CmdletBinding()]
    param ()

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
            Key         = 'Tvm'
            Description = 'Defender threat and vulnerability management report'
            Query       = $tvmQuery
            Select      = $tvmSelect
            Timespan    = 'P30D'
            TargetKey   = 'defenderTvmReport'
        }
        [PSCustomObject]@{
            Key         = 'Health'
            Description = 'Defender agent health report'
            Query       = $healthQuery
            Select      = $healthSelect
            Timespan    = 'P45D'
            TargetKey   = 'defenderHealthReport'
        }
        [PSCustomObject]@{
            Key         = 'SoftwareInventory'
            Description = 'Defender device software inventory'
            Query       = $softwareQuery
            Select      = @('*')
            Timespan    = 'P30D'
            TargetKey   = 'softwareInventory'
        }
    )
}
