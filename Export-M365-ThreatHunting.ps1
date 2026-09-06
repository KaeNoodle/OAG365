<#
.SYNOPSIS
OAG export of Microsoft 365 threat hunting reports

.DESCRIPTION
Uses the Microsoft.Graph PowerShell modules to run threat hunting KQL queries which generate
  1. Defender threats and vulnerability management export
  2. Defender agent heatlh status export
  3. Defender device software inventory

 Required Scope: MS Graph Application permissions for Threat Hunting - ThreatHunting.Read.All

 Authentication: The script requires the Graph API scope $script:exportM365MsGraphScopes 
                 If authenticating as a custom enterprise application these should be granted to the application. 
                 Otherwise, the user will be prompted to authenticate. These scopes will then be requested for the built-in MS Graph enterprise application and must be approved before the script can run successfully.

    PS1 Modules: Requires PowerShell module Microsoft.Graph.Authentication

.COMPONENT 
PowerShell module Microsoft.Graph.Authentication

.INPUTS
Authentication credentials for either a custom registered enterprise application or interactive login (to run as MS Graph application)

.OUTPUTS
CSV exports containing
  - Defender TVM report
  - Defender health report
  - Device software inventory

.PARAMETER AppClientId
The registered Entra application's client / application ID. Used when authenticating as an enterprise application by using a certificate thumbprint or a secret.

.PARAMETER AppTenantId
The registered Entra application's tenant ID. Used when authenticating as an enterprise application by using a certificate thumbprint or a secret.

.PARAMETER AppCertThumbprint
The thumbprint of the registered Entra application's authentication certificate. The certificate must be installed into the certificate store of the machine running the script. Required when using certificate based authentication for a registered enterprise application.

Certificates can be created in an elevated PowerShell sesssion:
  # Create the certificate 
  $cert = New-SelfSignedCertificate -Subject "CN=OagExportM365" -CertStoreLocation "Cert:\LocalMachine\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths(3)

  # Export as .cer file and import into the registered Entra Enterprise Application
  Export-Certificate -Cert $cert -FilePath "./OagExportScript.cer"

.PARAMETER AppSecret
The registered Entra application's authentication secret (plain text). Required when using a secret to authenticate as an enterprise application.

.PARAMETER Output
The name of the folder to output exported information to. Defaults to the running scripts folder ($PSScriptRoot\M365\YYYYMMDD_HHIISS\*).

.EXAMPLE
Export-M365-DefenderReport.ps1

By default the script will prompt for credentials. It will request the required scopes, running as the 'MS Graph' enterprise application. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-DefenderReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppSecret "a1bC2d~E3fGh4iJ5kL6mN7oP8qR9sT0uV1wX2yZ3"

The script will use the provided client ID, tenant ID and secret to authenticate as an enterprise application. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-DefenderReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppCertThumbprint "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"

The script will use the provided client ID, tenant ID and authenticate as an enterprise application using a certificate installed on the local machine that has the specified thumbprint. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-DefenderReport.ps1 -Output C:\temp

Optionally you can change the location the exported files are saved to. E.g. C:\temp\M365\YYYYMMDD_HHIISS\*

.NOTES
VERSION: 1.0

RELEASE APPROVAL: <UNAPPROVED>

2026-05-29
 - Initial version

Uses Microsoft.Graph PowerShell module:
    Connect-MgGraph
    Disconnect-MgGraph
    Invoke-MgGraphRequest

.LINK

#>
[Cmdletbinding(DefaultParameterSetName = 'Prompt')]
Param (
	[String]$output = $PSScriptRoot,

    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppClientId,
    
    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppTenantId,

    [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
    [string]$AppCertThumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
    [string]$AppSecret
)

begin {

# Define required graph API scopes
$script:exportM365MsGraphScopes = @(
    "ThreatHunting.Read.All"
)

function Format-Oag365-Exception {
    <#
    .SYNOPSIS
    Exception message handling. 

    .DESCRIPTION
    Create and return a formatted string containing and error message generated from the provided exceptions object.  

    .PARAMETER message
    Custom error message provided by the application to provide meaningful context 

    .PARAMETER exception
    Exception object as caught by try / catch

    .EXAMPLE
    Format-Oag365-Exception -message "Failed connecting Graph API to M365 tenant" -exception $e

    .NOTES
        NAME: Format-Oag365-Exception
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:

        CHANGELOG:
    #>
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][string]$message,
        [Parameter(Mandatory = $true)][AllowEmptyString()][object]$exception
    )

    $fmtString = "{0}`n`n{1} : {2}`n{3}`n" +
                 "      + CategoryInfo     : {4}`n" +
                 "      + FullyQualifiedId : {5}`n"
    
    $exceptionFields = $message, $exception.InvocationInfo.InvocationName, $exception.Exception.Message, $exception.InvocationInfo.PositionMessage, $exception.CategoryInfo.ToString(), $exception.FullyQualifiedErrorId

    return $fmtString -f $exceptionFields
}

function Connect-Oag365-ToGraph {
    <#
    .SYNOPSIS
    Connects PowerShell to M365 Microsoft Graph 

    .DESCRIPTION
    Imports required PowerShell modules, if required, and then connects the Graph PowerShell module.  

    .PARAMETER AppClientId
    The registered Entra application's app ID. Required when using the 'AppCertThumbprint' or 'AppSecret' parameter sets.

    .PARAMETER AppTenantId
    The M365 tenant Id for the registered Entra application. Required when using the 'AppCertThumbprint' or 'AppSecret' parameter sets.

    .PARAMETER AppCertThumbprint
    The thumbprint of the certificate used to authenticate to the registered Entra application. Certificate must be installed into the certificate store of the machine running the script. Required when using the 'AppCertThumbprint' parameter set to authenticate using an app_id and certificate based authentication.

    .PARAMETER AppSecret
    The registered Entra application's secret pass phrase (plain text). Required when using the 'AppSecret' parameter set to authenticate using an app_id and secret (password).

    .EXAMPLE
    Connect-Oag365-ToGraph

    .NOTES
        NAME: Connect-Oag365-ToGraph
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Connect-MgGraph
           Get-MgContext

        CHANGELOG:
          2025-11-18: Updated in include modules only when required. Part of allowing stages to be run independently.
    #>
    [Cmdletbinding(DefaultParameterSetName = 'Prompt')]
    param (       
        [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
        [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
        [string]$AppClientId,
        
        
        [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
        [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
        [string]$AppTenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "AppCertThumbprint")]
        [string]$AppCertThumbprint,

        [Parameter(Mandatory = $true, ParameterSetName = "AppSecret")]
        [string]$AppSecret
    )

    try {
        Write-Host "`nMicrosoft Graph required. Please login if prompted" -ForegroundColor Cyan
        Start-Sleep 2

        if (-not (Get-Module -Name $script:psGraphModulesRequired)) {
            Write-Host " - Importing PowerShell Graph modules..." -ForegroundColor Cyan
            Start-Sleep 2
            $script:psGraphModulesRequired | Import-Module -Verbose
        }
        
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            switch ($PSCmdlet.ParameterSetName) {
                'AppCertThumbprint' { 
                    Connect-MgGraph -ContextScope Process -ClientId $AppClientId -TenantId $AppTenantId -CertificateThumbprint $AppCertThumbprint
                    Write-Host " - Connecting to Microsoft Graph using application certificate thumbprint $AppCertThumbprint" -ForegroundColor Cyan
                }
                'AppSecret' { 
                    $secureString = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
                    $clientSecretCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppClientId, $secureString
                    
                    Connect-MgGraph -ContextScope Process -TenantId $AppTenantId -ClientSecretCredential $clientSecretCreds
                    Write-Host " - Connecting to Microsoft Graph using application secret " -ForegroundColor Cyan
                }
                'Prompt' { 
                    Connect-MgGraph -ContextScope Process -Scopes $script:exportM365MsGraphScopes    
                    Write-Host " - Connecting to Microsoft Graph`n   Prompting for scopes: $($script:exportM365MsGraphScopes)" -ForegroundColor Cyan
                 }
            }
        }

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            switch -Regex ($PSCmdlet.ParameterSetName) {
                'AppCertThumbprint|AppSecret' { 
                    Write-Host " - Connected as Enterprise Application" -ForegroundColor Cyan
                    Write-Host "       Tenant ID: $($context.TenantID)"
                    Write-Host "       Client ID: $($context.ClientID)"
                    Write-Host "        App Name: $($context.AppName)"
                    Write-Host "      Token Type: $($context.TokenCredentialType)"
                    Write-Host "       Auth Type: $($context.AuthType)"
                }
                'Prompt' { 
                    Write-Host " - Connected to Microsoft Graph PowerShell Enterprise Application using $($context.Account) account" -ForegroundColor Cyan
                }
            }

            Write-Host " - Reviewing active scopes" -ForegroundColor Cyan
            
            $compare = Compare-Object -ReferenceObject $script:exportM365MsGraphScopes -DifferenceObject $($context.Scopes) -IncludeEqual
            $compare | ForEach-Object { 
                if ($_.SideIndicator -eq '<=') { 
                    Write-Host "     - Missing $($_.InputObject)" -ForegroundColor Red 
                } elseif ($_.SideIndicator -eq '==') { 
                    Write-Host "     - Granted $($_.InputObject)" -ForegroundColor Green 
                } else { 
                    Write-Host "     - Granted $($_.InputObject) (not required now) " -ForegroundColor Cyan 
                } 
            }
            
            Write-Host "`n"
            Pause
        } else {
            throw "Authentication needed. Failed connecting to Microsoft Graph with scopes `$($script:exportM365MsGraphScopes -join ", ")` "
            pause
            return $false
        }

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed connecting to Microsoft Graph PowerShell" -exception $_
        Write-Host $exception -ForegroundColor Red
        pause
        return $false
    }
}


function Disconnect-Oag365-FromGraph {
    <#
    .SYNOPSIS
    Disconnects from M365 services

    .DESCRIPTION
    Disconnects from Connect-MgGraph and Connect-ExcangeOnline.

    .EXAMPLE
    Disconnect-Oag365-FromGraph

    .NOTES
        NAME: Disconnect-Oag365-FromGraph
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Disconnect-MgGraph

        CHANGELOG:    
    #>

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    } catch {
        Write-Host "ERROR: Failed disconnecting $($_.Exception.message)" -ForegroundColor Red
    }
}



    # -----------------------------------
    # Define variables required by script
    # -----------------------------------

    Clear-Host

    $script:output = ($output | Resolve-Path).Path
    $script:execDateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")

    # Create folders to store output
	$script:exportTargetFolder = [PSCustomObject]@{
		"defender"    = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\Defender"
	}

    # Define where exports will be saved
    $script:exportTarget = [pscustomobject]@{
        "defenderHealthReport"  = Join-Path -Path $script:exportTargetFolder.defender -ChildPath "defenderHealthReport.csv"
        "defenderTvmReport"     = Join-Path -Path $script:exportTargetFolder.defender -ChildPath "defenderTvmReport.csv"
        "softwareInventory"     = Join-Path -Path $script:exportTargetFolder.defender -ChildPath "softwareInventory.csv"
    }

    $script:psGraphModulesRequired = @(
        'Microsoft.Graph.Authentication'
    )

    $TvmKqlQuery = '
    DeviceTvmSoftwareVulnerabilities
    | join kind=inner (DeviceTvmSoftwareInventory) on DeviceId, SoftwareName
    | project DeviceId, DeviceName, OSPlatform, OSVersion, OSArchitecture, SoftwareVendor, SoftwareName, SoftwareVersion, CveId, VulnerabilitySeverityLevel, CveTags, AadDeviceId, Type, EndOfSupportStatus, EndOfSupportDate, ProductCodeCpe
    | join DeviceTvmSoftwareVulnerabilitiesKB on CveId
    | project DeviceId, AadDeviceId, DeviceName, OSPlatform, OSVersion, OSArchitecture, SoftwareVendor, SoftwareName, SoftwareVersion, CveId, VulnerabilitySeverityLevel, CvssScore, CvssVector, CveSupportability, IsExploitAvailable, CveTags, CveLastModifiedTime = LastModifiedTime, CvePublishedDate = PublishedDate, VulnerabilityDescription, AffectedSoftware, Type, EndOfSupportStatus, EndOfSupportDate, ProductCodeCpe
    '

    $TvmKqlQuerySelect = @(
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

    $HealthReportKqlQuery = '
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

    $HealthReportKqlQuerySelect = @(
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

    $SoftwareInventoryKqlQuery = 'DeviceTvmSoftwareInventory'



    # --------------------------------------------------------------------------------------------------------------
    # Initialise script. Start transcript, check for required modules and safe mode. Output parameters for audit log
    # --------------------------------------------------------------------------------------------------------------

    Start-Transcript -Path (Join-Path -Path $output -ChildPath "$($script:execDateTime)_ExportM365_transcript.txt")
    Write-Host "Starting $(Get-Date)"

    Write-Host "`n----------------------------------------------"
    Write-Host " This script will output audit evidence to a series of subfolders under the location '$output'"

    $script:psGraphModulesRequired | ForEach-Object {
        $getModuleResult = Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue 
        if ($getModuleResult) {
            $moduleVersionNo = ($getModuleResult | Sort-Object Version -Descending).Version | Join-String -Separator ", "
            Write-Host "   - Found required PowerShell module '$_' [$moduleVersionNo]." -ForegroundColor Green
        } else {
            Write-Host "   - Missing required PowerShell module '$_'. Abort now." -ForegroundColor Red
            Write-Host "     Try: Install-Module $_ -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Cyan
        }
    }

    Write-Host "   - If exceptions occur check that PowerShell modules are up-to-date." -ForegroundColor Cyan

    if ($ExecutionContext.SessionState.LanguageMode -eq "ConstrainedLanguage") {
	    Write-Host "   - ERROR. This script cannot run in Constrained Language Mode. Abort now." -ForegroundColor Red
    } elseif ($ExecutionContext.SessionState.LanguageMode -eq "FullLanguage") {
	    Write-Host "   - Running in Full Language Mode" -ForegroundColor Green
    } else {		
	    Write-Host "   - Caution'$($ExecutionContext.SessionState.LanguageMode)' Language Mode detected" -ForegroundColor DarkYellow
    }

	if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSVersionTable.PSVersion.Minor -ge 1) {
		Write-Host "   - PowerShell version $($PSVersionTable.PSVersion) found, PowerShell 7.1+ is recommended." -ForegroundColor Green
	} else {
		Write-Host "   - Caution, PowerShell version '$($PSVersionTable.PSVersion)' detected. PowerShell 7.1+ is recommended. Some parts of the script may not operate correctly." -ForegroundColor DarkYellow
	}

    Write-Host "----------------------------------------------`n"

    Write-Host "`nScript Parameters:" -ForegroundColor Cyan
    Write-Host "  Output Folder: $output"

    switch ($PSCmdlet.ParameterSetName) {
        'AppCertThumbprint' { 
            Write-Host "  Auth Mode: Custom Entra Enterprise Application using installed certificate"
            Write-Host "     Tenant ID: $AppTenantId"
            Write-Host "        App ID: $AppClientId"
            Write-Host "   Certificate: $AppCertThumbprint"

            $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $AppCertThumbprint }
            if ($cert) {
                Write-Host "                Found $($cert.Subject) expiring $($cert.NotAfter)" -ForegroundColor Green
            } else {
                Write-Host "                Error. Could not find certificate under Cert:\LocalMachine\My. This does not look promising." -ForegroundColor Red
            }
        }
        'AppSecret' { 
            Write-Host "  Auth Mode: Custom Entra Enterprise Application using secret"
            Write-Host "     Tenant ID: $AppTenantId"
            Write-Host "        App ID: $AppClientId"
            if ($AppSecret.Length -ne 40) {
                Write-Host "        Secret: $($AppSecret.Length) characters (expected 40)" -ForegroundColor Orange
            } else {
                Write-Host "        Secret: $($AppSecret.Length) characters" -ForegroundColor Green
            }
        }
        'Prompt' { 
            Write-Host "  Auth Mode: Interactive (Default to MS Graph Enterprise Application)"
        }
    }

    Write-Host "`n`n"
    Pause
}

process {

    Write-Host "`nGetting Defender for Endpoint threat hunting reports ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

    switch ($PSCmdlet.ParameterSetName) {
        'AppCertThumbprint' { 
            $Connection = Connect-Oag365-ToGraph -AppClientId $AppClientId -AppTenantId $AppTenantId -AppCertThumbprint $AppCertThumbprint
        }
        'AppSecret' { 
            $Connection = Connect-Oag365-ToGraph -AppClientId $AppClientId -AppTenantId $AppTenantId -AppSecret $AppSecret
        }
        'Prompt' { 
            $Connection = Connect-Oag365-ToGraph
        }
    }

    if ($Connection -eq $true) {
        try {
            Write-Host "`n Export Defender TVM report"

            $JsonBody = (@{
                Query    = $TvmKqlQuery
                Timespan = "P30D" 
            }) | ConvertTo-Json

            $Response = Invoke-MgGraphRequest -Method POST `
                                            -Uri "v1.0/security/runHuntingQuery" `
                                            -Body $JsonBody `
                                            -OutputType PSObject
            $Response.results | Select-Object -Property $TvmKqlQuerySelect | Export-Csv $script:exportTarget.defenderTvmReport             

            Write-Host " - Exported Defender TVM report to CSV $($script:exportTarget.defenderTvmReport)" -ForegroundColor Green
        } catch {
            $exception = Format-Oag365-Exception -message "Failed exporting Defender TVM report to CSV" -exception $_
        }

        try {
            Write-Host "`n Export Defender agent health report"
            $JsonBody = (@{
                Query    = $HealthReportKqlQuery
                Timespan = "P45D" 
            }) | ConvertTo-Json

            $Response = Invoke-MgGraphRequest -Method POST `
                                            -Uri "v1.0/security/runHuntingQuery" `
                                            -Body $JsonBody `
                                            -OutputType PSObject
            $Response.results | Select-Object -Property $HealthReportKqlQuerySelect | Export-Csv $script:exportTarget.defenderHealthReport

            Write-Host " - Exported Defender agent health report to CSV $($script:exportTarget.defenderHealthReport)" -ForegroundColor Green
        } catch {
            $exception = Format-Oag365-Exception -message "Failed exporting Defender agent health report to CSV" -exception $_
        }

        try {
            Write-Host "`n Export Defender device software inventory"
            $JsonBody = (@{
                Query    = $SoftwareInventoryKqlQuery
                Timespan = "P30D" 
            }) | ConvertTo-Json

            $Response = Invoke-MgGraphRequest -Method POST `
                                            -Uri "v1.0/security/runHuntingQuery" `
                                            -Body $JsonBody `
                                            -OutputType PSObject
            $Response.results | Select-Object -Property * | Export-Csv $script:exportTarget.softwareInventory

            Write-Host " - Exported Defender device software inventory to CSV $($script:exportTarget.softwareInventory)" -ForegroundColor Green
        } catch {
            $exception = Format-Oag365-Exception -message "Failed exporting Defender device software inventory to CSV" -exception $_
        }
    }
}

end {
    Write-Host "`n`nCompleted $(Get-Date). Disconnecting from Microsoft Graph..."
    Disconnect-Oag365-FromGraph

    Stop-Transcript
}



































#                                                                     
#                                                                     
#         OOOOOOOOO                 AAA                  GGGGGGGGGGGGG
#       OO:::::::::OO              A:::A              GGG::::::::::::G
#     OO:::::::::::::OO           A:::::A           GG:::::::::::::::G
#    O:::::::OOO:::::::O         A:::::::A         G:::::GGGGGGGG::::G
#    O::::::O   O::::::O        A:::::::::A       G:::::G       GGGGGG
#    O:::::O     O:::::O       A:::::A:::::A     G:::::G              
#    O:::::O     O:::::O      A:::::A A:::::A    G:::::G              
#    O:::::O     O:::::O     A:::::A   A:::::A   G:::::G    GGGGGGGGGG
#    O:::::O     O:::::O    A:::::A     A:::::A  G:::::G    G::::::::G
#    O:::::O     O:::::O   A:::::AAAAAAAAA:::::A G:::::G    GGGGG::::G
#    O:::::O     O:::::O  A:::::::::::::::::::::AG:::::G        G::::G
#    O::::::O   O::::::O A:::::AAAAAAAAAAAAA:::::AG:::::G       G::::G
#    O:::::::OOO:::::::OA:::::A             A:::::AG:::::GGGGGGGG::::G
#     OO:::::::::::::OOA:::::A               A:::::AGG:::::::::::::::G
#       OO:::::::::OO A:::::A                 A:::::A GGG::::::GGG:::G
#         OOOOOOOOO  AAAAAAA                   AAAAAAA   GGGGGG   GGBG
#                                                                     
#                                                                     

# SIG # Begin signature block
# MIIV5QYJKoZIhvcNAQcCoIIV1jCCFdICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD7wk/OP9HzIp+/
# CClfcUC0rn6LsvvXQyHtu02tPlvhrqCCEiAwggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggYaMIIEAqADAgECAhBiHW0M
# UgGeO5B5FSCJIRwKMA0GCSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBSb290IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5
# NTlaMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0G
# CSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjI
# ztNsfvxYB5UXeWUzCxEeAEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NV
# DgFigOMYzB2OKhdqfWGVoYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/3
# 6F09fy1tsB8je/RV0mIk8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05Zw
# mRmTnAO5/arnY83jeNzhP06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm
# +qxp4VqpB3MV/h53yl41aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUe
# dyz8rNyfQJy/aOs5b4s+ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz4
# 4MPZ1f9+YEQIQty/NQd/2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBM
# dlyh2n5HirY4jKnFH/9gRvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQY
# MBaAFDLrkpr/NZZILyhAQnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritU
# pimqF6TNDDAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNV
# HSUEDDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsG
# A1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nUm9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsG
# AQUFBzAChjpodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2Rl
# U2lnbmluZ1Jvb3RSNDYucDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0
# aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURh
# w1aVcdGRP4Wh60BAscjW4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0Zd
# OaWTsyNyBBsMLHqafvIhrCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajj
# cw5+w/KeFvPYfLF/ldYpmlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNc
# WbWDRF/3sBp6fWXhz7DcML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalO
# hOfCipnx8CaLZeVme5yELg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJs
# zkyeiaerlphwoKx1uHRzNyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z7
# 6mKnzAfZxCl/3dq3dUNw4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5J
# KdGvspbOrTfOXyXvmPL6E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHH
# j95Ejza63zdrEcxWLDX6xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2
# Bev6SivBBOHY+uqiirZtg0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/
# L9Uo2bC5a4CH2RwwggaLMIIE86ADAgECAhEAw++A/5SOEBq5lH+WfxuWEjANBgkq
# hkiG9w0BAQwFADBUMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1p
# dGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2
# MB4XDTI2MDIyNDAwMDAwMFoXDTI3MDIyNDIzNTk1OVoweTELMAkGA1UEBhMCQVUx
# GjAYBgNVBAgMEVdlc3Rlcm4gQXVzdHJhbGlhMSYwJAYDVQQKDB1PZmZpY2Ugb2Yg
# dGhlIEF1ZGl0b3IgR2VuZXJhbDEmMCQGA1UEAwwdT2ZmaWNlIG9mIHRoZSBBdWRp
# dG9yIEdlbmVyYWwwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQD5gix/
# lAQc8KFHyrteIOU2FXQ/Vt3lJ9bH74qvOgcxN/q/rXZJ6DS9nQqf5aMbuOEP+ALs
# cZM+TnNNQEgll2x0lnCRcggfA1Odd+vmDEkNlOqX9yEvuNqJWUguVQZ6xMLqZQKX
# j93kGr16FMAb11xK7iMzhYSpWg15BBcTmvLbjjFuRhZi59l0JJ2XJAitu7sjGLrd
# cUxN39fzBnvZleWGLfvrEeRq2XTS6eL3J6yzcAo8XE9H014EaZYacVChdrfD+Vff
# XhhBjv8GeDVzgsElSw2DGW7YZ66C9YX/F3a8ZnbMXSGETv3nRQxBK+i4ov5hVdwN
# ZUc8gcA6Se03oBZSJFJCHHVDAnA0gDxuiTvUt9MEbkP2xyDmhDy+UTBhBrB5Hi7Y
# 8kn2WQxs+DRfeeEMadPGwFkhUcFFug8Eeun1Rgd3EBkLfhSg1ZFJrgjXr2zWrnAS
# wS/AoNCU9JUVLHPexsiQLb0vVnvtPU+MumjJdF2L+Yt66GipUCUhK/kgTUEpBf1L
# S2+rRPYWopkBkfdb7+RFx+d85cRDlSprYFSkrUtv5vv8Qo0ChWKq81rPGI6m85m/
# IaLoNieAdYEbiD5uCoC5C3rcnVTrn1zcL+k6j8U2pMg8ate+zxIvucrouizQ+AQ7
# HpoHSMMvsTxe5HK2zwJe90DzU1IVQLMpKDQLBwIDAQABo4IBsTCCAa0wHwYDVR0j
# BBgwFoAUDyrLIIcouOxvSK4rVKYpqhekzQwwHQYDVR0OBBYEFL9SF9gFByCoXc7Q
# TiMSt6+guomCMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQM
# MAoGCCsGAQUFBwMDMEoGA1UdIARDMEEwNQYMKwYBBAGyMQECAQMCMCUwIwYIKwYB
# BQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMAgGBmeBDAEEATBJBgNVHR8E
# QjBAMD6gPKA6hjhodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWND
# b2RlU2lnbmluZ0NBUjM2LmNybDB5BggrBgEFBQcBAQRtMGswRAYIKwYBBQUHMAKG
# OGh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5n
# Q0FSMzYuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAm
# BgNVHREEHzAdgRtTZXJ2aWNlRGVza0BhdWRpdC53YS5nb3YuYXUwDQYJKoZIhvcN
# AQEMBQADggGBAE165OjPZVPtmGXn/ds4B71BfZz2gsXKS8/85WrM5vOcp9j4QsiX
# VKPQl8dTdjVUmWk90bfW/y3zILdZXy1ZrV94rAed5ylJNgrV7/qORMK5PTQyw8Cb
# +biSLoUQmN9wCpzgy75wfqF7oAsg8xFt/QiBoSUlYXUj8SC6QLr4EU19nPjse7/s
# 97p2Th7mtzPvC0/YQHEqmlycucjwc3gmnqysMyAGWg3R+GE00ewDInaOgaMUutF5
# oEGu7KeN3r0PqmGlgp7iHUY7dEzZHg7dl1EFMxFV/awjQVwqKQ7ov530owK4pML/
# 00yx8dLms1NBEozZ7K4iVIbyB+9ZAm0pefNG0BwI2APW36MDXL+JK9KtqymCDYPm
# liBtWUwXKS6NGs6fgUVUZQrWt6bbqBMJ+r1akXsy6WqYwum3ojwW8UVN+e9V+i5z
# 6fAsMKbo0E244eB6rPfeGbGS/2sfxMgVZPV1M4Lwu7Xqzs8zR8YTQNNoi3l55EeS
# 61itZNxutwwWtjGCAxswggMXAgEBMGkwVDELMAkGA1UEBhMCR0IxGDAWBgNVBAoT
# D1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGlnbyBQdWJsaWMgQ29kZSBT
# aWduaW5nIENBIFIzNgIRAMPvgP+UjhAauZR/ln8blhIwDQYJYIZIAWUDBAIBBQCg
# gYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0B
# CQQxIgQgOGEiQwMn7iIEA1rJqHwLP8jl5WcJZmb7jC5y8dL5Be4wDQYJKoZIhvcN
# AQEBBQAEggIAIUvBj3pEhoN4F9aHmr9WmVDWnl263aZ6NjCtioe7vbHaBwEkQQCp
# zE3jHAmgpevS+btPSARCX50PiGf5pAuY4fXO+I2SBmwA3h8JgbZTHyakQP4FEw4f
# MROZDcPgB009WKNjcGQ2rh7HwBKiA2Vxjl2d2LXjgFAReI9dCTLM7P3g+a/yg27k
# iYNuLxSCuYdQUEM0AlrM1SsVeWOzM7WIM0InL+dIqS3lz7zZSb2LCwu2rpxOPak8
# BuQ6Mi4NU6H47ZQQU4gt3vfhWW8C8mMEH2iCFhziFMug9nc6oflmKcHqQUWRMhU/
# BcN37lT6UBaSHMpAafwi10xFe1iqzsYG5OiaCNZLzF3Xj9Kq9+HGJ4zzlY5lc+F3
# YH5uIvIuJDIcafYM7vMJAgichyqJl8AIbEs6WC3QdjwBuXO9YlYK8DP+3QwFTDaT
# HQjDNe8EzwZWtkwpiLN9dbJuG1vQHUOhdcey0xgonV45ojZUeYJB1pyZcnE5I0LF
# xJOPj6JEMdt+qSISeEPrJQH4bfU27yreT7+oXflhKOdS7BGGs4xxgHaH9Y4ymyOS
# NQzuZ+bqLcE6Czarp+kjdSCpf4nLGfzO2Hlji4pZaQAEZ9rqD+FCa13LkULz2Z4Y
# 3Hd7xR/nPyZVz9u328L1YFVmuSY1xX6YwIgMHxA9EzDgB7yXLcCwXn4=
# SIG # End signature block
