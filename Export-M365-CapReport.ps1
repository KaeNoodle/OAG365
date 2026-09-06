<#
.SYNOPSIS
OAG export of Microsoft 365 conditional access policies

.DESCRIPTION
Uses the Microsoft.Graph PowerShell modules to export conditional access policy details.

         Export: Microsoft Entra Conditional Access, Authentication Method and Authentication Strength policies

 Required Scope: MS Graph Application permissions for Policy.Read.AuthenticationMethod, Policy.Read.All, AuthenticationContext.Read.All, Policy.Read.ConditionalAccess, RoleManagement.Read.Directory, Application.Read.All, Directory.Read.All

 Authentication: The script requires the Graph API scopes defined in $script:exportM365MsGraphScopes. 
                 If authenticating as a custom enterprise application these should be granted to the application. 
                 Otherwise, the user will be prompted to authenticate. These scopes will then be requested for the built-in MS Graph enterprise application and must be approved before the script can run successfully.

    PS1 Modules: Requires PowerShell module Microsoft.Graph 
                 Specific requirements are documented in $script:psGraphModulesRequired

.COMPONENT 
PowerShell module Microsoft.Graph 

.INPUTS
Authentication credentials for either a custom registered enterprise application or interactive login (to run as MS Graph application)

.OUTPUTS
CSV exports and HTML reports containing
  - Conditional access policies
  - Authentication method policies 
  - Authentication strength policies

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
Export-M365-CapReport.ps1

By default the script will prompt for credentials. It will request the required scopes, running as the 'MS Graph' enterprise application. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-CapReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppSecret "a1bC2d~E3fGh4iJ5kL6mN7oP8qR9sT0uV1wX2yZ3"

The script will use the provided client ID, tenant ID and secret to authenticate as an enterprise application. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-CapReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppCertThumbprint "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"

The script will use the provided client ID, tenant ID and authenticate as an enterprise application using a certificate installed on the local machine that has the specified thumbprint. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-CapReport.ps1 -Output C:\temp

Optionally you can change the location the exported files are saved to. E.g. C:\temp\M365\YYYYMMDD_HHIISS\*

.NOTES
VERSION: 3.1

RELEASE APPROVAL: <UNAPPROVED>

2026-05-20
 - Added Enterprise Application authentication and improved documentation

2026-05-12
  - Code clean-up. Updated to use approved verbs and clearer function names

2026-05-01: Created as standalone script

Uses Microsoft.Graph PowerShell module:
    # Global
    Connect-MgGraph
    Disconnect-MgGraph
    Get-MgContext
    Get-MgDirectoryObject

    # Conditional access report
    Get-MgPolicyAuthenticationMethodPolicy
    Get-MgPolicyAuthenticationStrengthPolicy
    Get-MgIdentityConditionalAccessPolicy
    Get-MgIdentityConditionalAccessNamedLocation
    Get-MgIdentityConditionalAccessAuthenticationContextClassReference
    Get-MgDirectoryRoleTemplate
    Get-MgServicePrincipal
  
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
    "Policy.Read.AuthenticationMethod",
    "Policy.Read.All",
    "AuthenticationContext.Read.All",
    "Policy.Read.ConditionalAccess",
    "RoleManagement.Read.Directory",
    "Application.Read.All",
    "Directory.Read.All"
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


function Test-Oag365Cap-Guid {
    <#
    .SYNOPSIS
    Validates a given input string and checks string is a valid GUID

    .DESCRIPTION
    Validates a given input string and checks string is a valid GUID by using the .NET method Guid.TryParse

    .PARAMETER InputObject
    The GUID(s) to test

    .EXAMPLE
    Test-Oag365Cap-Guid -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Test-Oag365Cap-Guid
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

function Resolve-Oag365Cap-DirectoryObject {
    <#
    .SYNOPSIS
    Resolve a Microsoft Graph item GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Graph item GUID to an object containing Get-MgDirectoryObject properties including user principal name, display name and app display name

    .PARAMETER InputObject
    The GUID(s) of the directory object to resolve

    .EXAMPLE
    Resolve-Oag365Cap-DirectoryObject -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-Oag365Cap-DirectoryObject
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryObject
           Get-MgServicePrincipal

        CHANGELOG:    
           2025-11-17 Added Get-MgServicePrincipal to handle scenario of build in Applications not resolving (e.g. SharePoint / OneDrive)
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-Oag365Cap-Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheMgObject.ContainsKey($InputObject)) {
                    Write-Debug "Cached display name for `"$InputObject`""
                    Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    return $script:cacheMgObject[$InputObject]
                } else {
                    try {
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $InputObject -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]
                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    } catch {
                        $exception = Format-Oag365-Exception -message "Unable to find object using Get_MgDirectoryObject using provided GUID" -exception $_
                        Write-Host $exception -ForegroundColor DarkYellow

                        $directoryObjectId = (Get-MgServicePrincipal -filter "appId eq '$($InputObject)'").id
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $directoryObjectId -ErrorAction Stop
                        $displayName = $directoryObject.AdditionalProperties["displayName"]

                        Write-Host " - Don't panic. I found the service principal '$($displayName)' using the provided string as an AppID instead."

                        $script:cacheMgObject[$InputObject] = $directoryObject
                        Write-Verbose " - GUID Lookup from ServicePrincipal AppId: $(($script:cacheMgObject[$InputObject]).AdditionalProperties.displayName)"
                    }
                    return $directoryObject
                }
            }
            catch {
                Write-Warning "Unable to resolve directory object with ID $InputObject, might have been deleted!"
            }
        }
        return $InputObject
    }
}


$script:cacheNamedLocations = @{}

function Resolve-Oag365Cap-NamedLocation {
    <#
    .SYNOPSIS
    Resolve a Microsoft Conditional Access Policy named location GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Conditional Access Policy named location GUID to an object containing Get-MgIdentityConditionalAccessNamedLocation properties including display name and CIDR IP address details

    .PARAMETER InputObject
    The GUID(s) of the named location to resolve

    .EXAMPLE
    Resolve-Oag365Cap-NamedLocation -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-Oag365Cap-NamedLocation
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityConditionalAccessNamedLocation

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-Oag365Cap-Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheNamedLocations.ContainsKey($InputObject) -eq $false) {
                    $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
                    $namedLocations | ForEach-Object { $script:cacheNamedLocations[$_.Id] = $_ }
                    Write-Host "Found $($script:cacheNamedLocations.Count) named locations, adding to cache"
                }

                Write-Verbose "Loading cached named location '$InputObject' $(($script:cacheNamedLocations[$InputObject]).DisplayName)"
                return $script:cacheNamedLocations[$InputObject]
            } catch {
                Write-Warning "Unable to resolve named location with ID $InputObject!"
            }
        }
        return $InputObject
    }
}

$script:cacheAdminUnits = @{}

function Resolve-Oag365Cap-EntraAdminUnit {
    <#
    .SYNOPSIS
    Resolve an Entra Administrative Unit ID to an object

    .DESCRIPTION
    Resolve an Entra Administrative Unit ID to an object

    .PARAMETER InputObject
    The GUID(s) of the admin unit to resolve

    .EXAMPLE
    Resolve-Oag365Cap-EntraAdminUnit -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-Oag365Cap-EntraAdminUnit
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryAdministrativeUnit

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-Oag365Cap-Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheAdminUnits.ContainsKey($InputObject) -eq $false) {
                    $adminUnits = Get-MgDirectoryAdministrativeUnit -All -ErrorAction Stop
                    $adminUnits | ForEach-Object { $script:cacheAdminUnits[$_.Id] = $_ }
                    Write-Host "Found $($script:cacheAdminUnits.Count) admin units, adding to cache"
                }

                Write-Verbose "Loading cached admin unit '$InputObject' $(($script:cacheAdminUnits[$InputObject]).DisplayName)"
                return $script:cacheAdminUnits[$InputObject]
            } catch {
                Write-Warning "Unable to resolve admin unit with ID $InputObject!"
            }
        }
        return $InputObject
    }
}

$script:cacheServicePrincipals = @{}

function Resolve-Oag365Cap-EntraServicePrincipal {
    <#
    .SYNOPSIS
    Resolve a service principal appId

    .DESCRIPTION
    Resolve a service principal appId to an object

    .PARAMETER InputObject
    The appId of the service principal to resolve

    .EXAMPLE
    Resolve-Oag365Cap-EntraServicePrincipal -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-Oag365Cap-EntraServicePrincipal
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgServicePrincipal

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        if (Test-Oag365Cap-Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($script:cacheServicePrincipals.ContainsKey($InputObject) -eq $false) {
                    $script:cacheServicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop | Group-Object -Property AppId -AsHashTable
                    Write-Host "Found $($script:cacheServicePrincipals.Count) service principals, adding to cache"
                }

                Write-Verbose "Loading cached service principal '$InputObject' $(($script:cacheServicePrincipals[$InputObject]).DisplayName)"
                return $script:cacheServicePrincipals[$InputObject]
            } catch {
                Write-Warning "Unable to resolve service principal with ID $InputObject!"
            }
        }
        return $InputObject
    }
}


$script:cacheDirectoryRoleTemplates = @{}

function Resolve-Oag365Cap-DirectoryRoleTemplate {
    <#
    .SYNOPSIS
    Resolve a directory role template 

    .DESCRIPTION
    Resolve a directory role template Id to a display name

    .PARAMETER InputObject
    The Id of the directory role template to resolve

    .EXAMPLE
    Resolve-Oag365Cap-DirectoryRoleTemplate -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Resolve-Oag365Cap-DirectoryRoleTemplate
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryRoleTemplate

        CHANGELOG:
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )

    process {
        try {
            # use hashtable as cache to limit API calls
            if (($script:cacheDirectoryRoleTemplates.ContainsKey($InputObject) -eq $false)) {
                $script:cacheDirectoryRoleTemplates = Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Group-Object -Property Id -AsHashTable
                Write-Host "Found $($script:cacheDirectoryRoleTemplates.Count) directory role templates, adding to cache"
            }

            Write-Verbose "Loading cached directory role template '$InputObject' $(($script:cacheDirectoryRoleTemplates[$InputObject]).DisplayName)"
            return $script:cacheDirectoryRoleTemplates[$InputObject].DisplayName
        } catch {
            Write-Warning "Unable to resolve directory role template with ID '$InputObject'"
        }

        return $InputObject
    }
}



function Convert-Oag365Cap-LocationToHTML {
    <#
    .SYNOPSIS
    Outputs HTML table formatted rows containing location data from a conditional access policy

    .DESCRIPTION
    Outputs HTML table formatted rows containing location data from a conditional access policy

    .PARAMETER locations
    List of locations to iterate through and format.

    .PARAMETER type
    IncludeLocations or ExcludeLocations. Determins how row containing location data will be described in HTML report.

    .EXAMPLE
    Convert-Oag365Cap-LocationToHTML -locations $object -type 'IncludeLocation'

    .NOTES
        NAME: Convert-Oag365Cap-LocationToHTML
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-Oag365Cap-NamedLocation
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up named locations
    #>

    [CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][AllowNull()]$locations,
		[Parameter(Mandatory = $true)][ValidateSet('IncludeLocation', 'ExcludeLocation')]$type
    )

    if ($null -eq $locations) { 
        return ''; 
    }

    foreach ($location in $locations) {
        $thisLocation =  Resolve-Oag365Cap-NamedLocation -InputObject $location

        $out_displayName = "(no name found)"
        if ($thisLocation.DisplayName) {
            $out_displayName = $thisLocation.DisplayName
        }

        $out_cidrAddress = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('cidrAddress')) {
            $out_cidrAddress = ($thisLocation.AdditionalProperties['cidrAddress'] | Sort-Object | Out-String)
        }

        $out_ipRanges = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('ipRanges')) {
            $out_ipRanges = ($thisLocation.AdditionalProperties['ipRanges'] | ForEach-Object{ $_.cidrAddress } | Sort-Object) -join $separator
        }

        $out_countriesAndRegions = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('countriesAndRegions')) {
            $out_countriesAndRegions = ($thisLocation.AdditionalProperties['countriesAndRegions'] | Sort-Object | Out-String )
        }

        $out_includeUnknownCountriesAndRegions = ""
        if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('includeUnknownCountriesAndRegions')) {
            $out_cidrAddress = ($thisLocation.AdditionalProperties['includeUnknownCountriesAndRegions'] | Sort-Object | Out-String)
        }

        '<tr><td>{0}<br/><nobr>{1}<nobr></td><td>{2}</td><td><p>CIDR Addresses: {3} </p><p>IP Ranges: {4}</p><p>Countries and Regions: {5} </p><p>Include Unknown Countries and Regions: {6}</p></td></tr>' -f $out_displayName, $location, $type, $out_cidrAddress, $out_ipRanges, $out_countriesAndRegions, $out_includeUnknownCountriesAndRegions
    }
}

function Resolve-Oag365Cap-DirectoryObjects {
    <#
    .SYNOPSIS
    Outputs a list of objects that can be resolved by Resolve-Oag365Cap-DirectoryObject (user, group, etc.) based on their GUID from a conditional access policy

    .DESCRIPTION
    Outputs a list of objects that can be resolved by Resolve-Oag365Cap-DirectoryObject (user, applications, etc.) based on their GUID from a conditional access policy

    .PARAMETER srcObject
    Iterable source object containing GUIDs. Uses ForEach-Object pipeline

    .PARAMETER getAttrib
    Object attribute to return in list (e.g. displayName or userPrinicpalName)

    .EXAMPLE
    Resolve-Oag365Cap-DirectoryObjects -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"

    .NOTES
        NAME: Resolve-Oag365Cap-DirectoryObjects
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Resolve-Oag365Cap-DirectoryObject
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$srcObject,
        [Parameter(Mandatory = $true)][ValidateSet('displayName','userPrincipalName')]$getAttrib
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        if (Test-Oag365Cap-Guid -InputObject $PSItem) {
            $object = Resolve-Oag365Cap-DirectoryObject -InputObject $PSItem
            switch ($getAttrib) {
                "userPrincipalName" {
                    $list.Add(($object.AdditionalProperties.userPrincipalName)) | Out-Null                    
                }
                "displayName" {
                    $list.Add(($object.AdditionalProperties.displayName)) | Out-Null                    
                }
            }
        } else {
            Write-Verbose " - $($PSItem)"
            $list.Add($PSItem) | Out-Null
        }
    }

    return $list
}

function Resolve-Oag365Cap-ServicePrincipals {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of service principal objects

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of service principal objects

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    Resolve-Oag365Cap-ServicePrincipals -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals

    .NOTES
        NAME: Resolve-Oag365Cap-ServicePrincipals
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-Oag365Cap-EntraServicePrincipal
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((Resolve-Oag365Cap-EntraServicePrincipal -InputObject $PSItem)) | Out-Null
    }

    return $list
}


function Resolve-Oag365Cap-DirectoryRoles {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of directory roles

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of directory roles

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    Resolve-Oag365Cap-DirectoryRoles -srcObject $policy.Conditions.Users.IncludeRoles

    .NOTES
        NAME: Resolve-Oag365Cap-DirectoryRoles
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-Oag365Cap-DirectoryRoleTemplate
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((Resolve-Oag365Cap-DirectoryRoleTemplate -InputObject $PSItem)) | Out-Null
    }

    return $list
}

function Convert-Oag365Cap-PolicyToHtml {
    <#
    .SYNOPSIS
    Build an HTML report from a conditional access policy

    .DESCRIPTION
    Build an HTML report from a conditional access policy

    .PARAMETER policy
    Conditional access policy object obtained from Microsoft Graph function Get-MgIdentityConditionalAccessPolicy

    .EXAMPLE
    Convert-Oag365Cap-PolicyToHtml -policy $policy

    .NOTES
        NAME: Convert-Oag365Cap-PolicyToHtml
        VERSION: 1.2

        FUNCTIONS & PERMISSIONS:
           Resolve-Oag365Cap-DirectoryObjects
           Resolve-Oag365Cap-ServicePrincipals
           Convert-Oag365Cap-LocationToHTML
           Resolve-Oag365Cap-DirectoryObjects
           Resolve-Oag365Cap-DirectoryRoles
           Test-Oag365Cap-Guid
           Resolve-Oag365Cap-DirectoryObject
           ExportM365-GroupMembers
           
        CHANGELOG:
          2025-11-21: Implemented formatting and resolver functions
    #>
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)]$policy
	)

try {
$separator = "</br>"

$outHtml = @"
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">
    <style type="text/css">
        table.property-value tbody tr td:nth-child(1) `{
            width: 20%;
            text-wrap: nowrap;
        `}

        table.set5050 tbody tr td `{
            width: 50%;
        }
    </style>
</head>
<body>
"@

$outHtml = $outHtml + @'
<header class="w3-container">
    <h1>{0}</h1>
</header>

<div class="w3-container">
<table class="w3-table-all w3-hoverable w3-small property-value">
    <tr>
        <td class="w3-dark-grey">Policy Description:</td>
        <td>{1}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Policy ID:</td>
        <td>{2}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Policy State:</td>
        <td>{3}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Policy Created:</td>
        <td>{4}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Policy Modified:</td>
        <td>{5}</td>
    </tr>
</table>
</div>

<header class="w3-container w3-margin-top">
    <h2 class="w3-blue">Conditions</h2>
</header>
'@ -f $policy.DisplayName, $policy.Description, $policy.Id, $policy.State, $policy.CreatedDateTime, $policy.ModifiedDateTime



# Sign-in Risk Levels Section
$out_signInRiskLevels = ""
if ($policy.Conditions.SignInRiskLevels) {
    $out_signInRiskLevels = $policy.Conditions.SignInRiskLevels -join $separator
}

$out_userRiskLevels = ""
if ($policy.Conditions.UserRiskLevels) {
    $out_userRiskLevels = $policy.Conditions.UserRiskLevels -join $separator
}

$out_servicePrincipalRiskLevels = ""
if ($policy.Conditions.servicePrincipalRiskLevels) {
    $out_servicePrincipalRiskLevels = $policy.Conditions.servicePrincipalRiskLevels -join $separator
}

$outHtml = $outHtml + @'
<div class="w3-container">
    <h3>Sign-in Risk Levels</h3>

    <table class="w3-table-all w3-hoverable w3-small">
    <thead>
        <tr class="w3-dark-grey">
            <td>SignInRiskLevels</td>
            <td>UserRiskLevels</td>
            <td>ServicePrincipalRiskLevels</td>
        </tr>
    </thead>
    <tbody>
        <tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        </tr>
    </tbody>
    </table>
</div>
'@ -f $out_signInRiskLevels, $out_userRiskLevels, $out_servicePrincipalRiskLevels


# Target Resources Section
# Resolve object IDs of included applications
$includeApps = Resolve-Oag365Cap-DirectoryObjects -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"
$excludeApps = Resolve-Oag365Cap-DirectoryObjects -srcObject $policy.Conditions.Applications.ExcludeApplications -getAttrib "displayName"

$out_includeApps = $includeApps -join $separator
$out_excludeApps = $excludeApps -join $separator

$out_includeUserActions = ""
if ($policy.Conditions.Applications.IncludeUserActions) {
    $out_includeUserActions = $policy.Conditions.Applications.IncludeUserActions -join $separator
}

$includeAuthenticationContext = [System.Collections.Generic.List[Object]]::new()
$policy.Conditions.Applications.IncludeAuthenticationContextClassReferences | ForEach-Object {
    $context = Get-MgIdentityConditionalAccessAuthenticationContextClassReference -Filter "Id eq '$PSItem'"
    $includeAuthenticationContext.Add($context.DisplayName) | Out-Null
}

$out_includeAuthenticationContext = $includeAuthenticationContext -join $separator

$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h3>Target Resources</h3>

    <table class="w3-table-all w3-hoverable w3-small">
    <thead>
        <tr class="w3-dark-grey">
            <td>Included Applications</td>
            <td>Excluded Applications</td>
            <td>Included User Actions</td>
            <td>Authentication Context</td>
        </tr>
    </thead>
    <tbody>
        <tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        </tr>
    </tbody>
    </table>
</div>
'@ -f $out_includeApps, $out_excludeApps, $out_includeUserActions, $out_includeAuthenticationContext



# Client Applications Section
$out_clientAppTypes = ""
if ($policy.Conditions.ClientAppTypes) {
    $out_clientAppTypes = $policy.Conditions.ClientAppTypes -join $separator
}

$includeServicePrincipals = Resolve-Oag365Cap-ServicePrincipals -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals
$excludeServicePrincipals = Resolve-Oag365Cap-ServicePrincipals -srcObject $policy.Conditions.ClientApplications.ExcludeServicePrincipals

$out_includeServicePrincipals = $includeServicePrincipals -join $separator
$out_excludeServicePrincipals = $excludeServicePrincipals -join $separator

$out_servicePrincipalFilterMode = ""
if ($policy.Conditions.ClientApplications.ServicePrincipalFilter.mode) {
    $out_servicePrincipalFilterMode = $policy.Conditions.ClientApplications.ServicePrincipalFilter.mode
}

$out_servicePrincipalFilterRule = ""
if ($policy.Conditions.ClientApplications.ServicePrincipalFilter.rule) {
    $out_servicePrincipalFilterRule = $policy.Conditions.ClientApplications.ServicePrincipalFilter.rule
}


$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h3>Client Applications</h3>

    <table class="w3-table-all w3-hoverable w3-small">
    <thead>
        <tr class="w3-dark-grey">
            <td>Client App Types</td>
            <td>Included ServicePrincipals</td>
            <td>Excluded ServicePrincipals</td>
            <td>ServicePrincipal Filter Mode</td>
            <td>ServicePrincipal Filter</td>
        </tr>
    </thead>
    <tbody>
        <tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>{4}</td>
        </tr>
    </tbody>
    </table>
</div>
'@ -f $out_clientAppTypes, $out_includeServicePrincipals, $out_excludeServicePrincipals, $out_servicePrincipalFilterMode, $out_servicePrincipalFilterRule



# Locations Section
$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h3>Locations</h3>

    <table class="w3-table-all w3-hoverable w3-small">
    <thead>
        <tr class="w3-dark-grey">
            <td>Named Location</td>
            <td>Condition Type</td>
            <td>Details</td>
        </tr>
    </thead>
    <tbody>
'@


try {
    $outHtml += Convert-Oag365Cap-LocationToHTML -type "IncludeLocation" -locations $policy.conditions.Locations.IncludeLocations 
    $outHtml += Convert-Oag365Cap-LocationToHTML -type "ExcludeLocation" -locations $policy.conditions.Locations.ExcludeLocations 
} catch {
    $outHtml += '<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>'
}

$outHtml = $outHtml + @'
    </tbody>
    </table>
</div>
'@


# Platforms Section
$out_includePlatforms = ""
if ($policy.Conditions.Platforms.IncludePlatforms) {
    $out_includePlatforms = $policy.Conditions.Platforms.IncludePlatforms -join $separator
}

$out_excludePlatforms = ""
if ($policy.Conditions.Platforms.ExcludePlatforms) {
    $out_excludePlatforms = $policy.Conditions.Platforms.ExcludePlatforms -join $separator
}

$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h3>Platforms</h3>

    <table class="w3-table-all w3-hoverable w3-small property-value">
        <tbody>
            <tr>
                <td class="w3-dark-grey">Included Platforms</td>
                <td>{0}</td>
            </tr>
            <tr>
                <td class="w3-dark-grey">Excluded Platforms</td>
            <td>{1}</td>
            </tr>
        </tbody>
    </table>
</div>
'@ -f $out_includePlatforms, $out_excludePlatforms


# Device Filter Rule Section
$out_deviceFilterRule = ""
if ($policy.Conditions.Devices.DeviceFilter.Mode) {
    $out_deviceFilterRule = $policy.Conditions.Devices.DeviceFilter.Mode
}

$out_deviceFilterMode = ""
if ($policy.Conditions.Devices.DeviceFilter.Rule) {
    $out_deviceFilterMode = $policy.Conditions.Devices.DeviceFilter.Rule
}

$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h3>Device Filter Rule</h3>

    <table class="w3-table-all w3-hoverable w3-small property-value">
        <tbody>
            <tr>
                <td class="w3-dark-grey">Filter Rule</td>
                <td>{0}</td>
            </tr>
            <tr>
                <td class="w3-dark-grey">Filter Mode</td>
                <td>{1}</td>
            </tr>
        </tbody>
    </table>
</div>
'@ -f $out_deviceFilterRule, $out_deviceFilterMode


$includeUsers = Resolve-Oag365Cap-DirectoryObjects -srcObject $policy.Conditions.Users.IncludeUsers -getAttrib "userPrincipalName"
$excludeUsers = Resolve-Oag365Cap-DirectoryObjects -srcObject $policy.Conditions.Users.ExcludeUsers -getAttrib "userPrincipalName"

# Resolve object IDs of included groups
$includeGroups = [System.Collections.ArrayList]::new()
$includeGroupsMembers = @{}
$policy.Conditions.Users.IncludeGroups | ForEach-Object {
    if (Test-Oag365Cap-Guid -InputObject $PSItem) {
        $group = Resolve-Oag365Cap-DirectoryObject -InputObject $PSItem
        $includeGroups.Add($group) | Out-Null

        $groupMembers = ExportM365-GroupMembers -groupId $group.Id
        if ($includeGroupsMembers.ContainsKey($group.AdditionalProperties.displayName) -eq $false) {
            $includeGroupsMembers[$group.AdditionalProperties.displayName] = [System.Collections.ArrayList]::new()
        }

        $groupMembers | ForEach-Object { 
            $includeGroupsMembers[$group.AdditionalProperties.displayName] += [PSCustomObject]@{
                DisplayName = $_.AdditionalProperties.displayName
                UserPrincipalName = $_.AdditionalProperties.userPrincipalName
            }
        } | Out-Null
    } else {
        Write-Verbose " - $($PSItem)"
        $includeGroups.Add($PSItem) | Out-Null
    }
}

# Resolve object IDs of excluded groups
$excludeGroups = [System.Collections.ArrayList]::new()
$excludeGroupsMembers = @{}
$policy.Conditions.Users.ExcludeGroups | ForEach-Object {
    if (Test-Oag365Cap-Guid -InputObject $PSItem) {
        $group = Resolve-Oag365Cap-DirectoryObject -InputObject $PSItem
        $excludeGroups.Add($group) | Out-Null

        $groupMembers = ExportM365-GroupMembers -groupId $group.Id
        if ($excludeGroupsMembers.ContainsKey($group.AdditionalProperties.displayName) -eq $false) {
            $excludeGroupsMembers[$group.AdditionalProperties.displayName] = [System.Collections.ArrayList]::new()
        }

        $groupMembers | ForEach-Object { 
            $excludeGroupsMembers[$group.AdditionalProperties.displayName] += [PSCustomObject]@{
                DisplayName = $_.AdditionalProperties.displayName
                UserPrincipalName = $_.AdditionalProperties.userPrincipalName
            }
        } | Out-Null
    } else {
        Write-Verbose " - $($PSItem)"
        $excludeGroups.Add($PSItem) | Out-Null
    }
}

$htmlIncludeGroupMembers = ""
if ($includeGroupsMembers.count -gt 0) {
    $htmlIncludeGroupMembers = "Active Group Members: <br/><small>For full membership of PIM enabled groups see role export.</small><ul>"
    foreach ($key in $includeGroupsMembers.Keys) {
        $members = $includeGroupsMembers[$key]
        $htmlIncludeGroupMembers += "<li>$key<ul>"
        foreach ($member in $members) { 
            $htmlIncludeGroupMembers += "<li>$($member.displayName) [$($member.userPrincipalName)]</li>" 
        }
        $htmlIncludeGroupMembers += "</ul></li>"
    }
    $htmlIncludeGroupMembers += "</ul>"
}

$htmlExcludeGroupMembers = ""
if ($excludeGroupsMembers.count -gt 0) {
    $htmlExcludeGroupMembers = "Active Group Members: <br/><small>For full membership of PIM enabled groups see role export.</small><ul>"
    foreach ($key in $excludeGroupsMembers.Keys) {
        $members = $excludeGroupsMembers[$key]
        $htmlExcludeGroupMembers += "<li>$key<ul>"
        foreach ($member in $members) { 
            $htmlExcludeGroupMembers += "<li>$($member.displayName) [$($member.userPrincipalName)] $($separator)</li>" 
        }
        $htmlExcludeGroupMembers += "</ul></li>"
    }
    $htmlExcludeGroupMembers += "</ul>"
}

$includeRoles = Resolve-Oag365Cap-DirectoryRoles -srcObject $policy.Conditions.Users.IncludeRoles
$excludeRoles = Resolve-Oag365Cap-DirectoryRoles -srcObject $policy.Conditions.Users.ExcludeRoles

# Users Section
$outHtml = $outHtml + @"
<div class="w3-container w3-margin-top">
    <h3>Users, Groups and Roles</h3>

    <h4>Users</h4>
    <table class="w3-table-all w3-hoverable w3-small set5050">
    <thead>
        <tr class="w3-dark-grey">
            <td>Included Users</td>
            <td>Excluded Users</td>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>$($includeUsers -join $separator)</td>
            <td>$($excludeUsers -join $separator)</td>
        </tr>
    </tbody>
    </table>

    <h4>Groups</h4>
    <table class="w3-table-all w3-hoverable w3-small set5050">
    <thead>
        <tr class="w3-dark-grey">
            <td>Included Groups</td>
            <td>Excluded Groups</td>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>$(($includeGroups | ForEach-Object { $_.AdditionalProperties.displayName + $separator } ))</td>
            <td>$(($excludeGroups | ForEach-Object { $_.AdditionalProperties.displayName + $separator } ))</td>
        </tr>
        <tr>
            <td>$($htmlIncludeGroupMembers)</td>
            <td>$($htmlExcludeGroupMembers)</td>
        </tr>
    </tbody>
    </table>

    <h4>Roles</h4>
    <table class="w3-table-all w3-hoverable w3-small set5050">
    <thead>
        <tr class="w3-dark-grey">
            <td>Included Roles</td>
            <td>Excluded Roles</td>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>$($includeRoles -join $separator)</td>
            <td>$($excludeRoles -join $separator)</td>
        </tr>
    </tbody>
    </table>
</div>
"@

# External Guests and Tenants
$out_includeGuestOrExternalUserTypes = ""
if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $out_includeGuestOrExternalUserTypes = ($policy.Conditions.Users.IncludeGuestsOrExternalUsers.guestOrExternalUserTypes -split ',') -join $separator
}

$out_includeMembershipKind = ""
if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.MembershipKind) {
    $out_includeMembershipKind = $policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.MembershipKind
}

$out_includeMembers = ""
if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"]) {
    $out_includeMembers = $policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"] -join $separator
}

$out_excludeGuestOrExternalUserTypes = ""
if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $out_excludeGuestOrExternalUserTypes = ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.guestOrExternalUserTypes -split ',') -join $separator
}

$out_excludeMembershipKind = ""
if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.MembershipKind) {
    $out_excludeMembershipKind = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.MembershipKind
}

$out_excludeMembers = ""
if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"]) {
    $out_excludeMembers = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"] -join $separator
}

$outHtml = $outHtml + @'
<div class="w3-container w3-margin-top">
    <h4>External Guests and Tenants</h4>

    <table class="w3-table-all w3-hoverable w3-small">
    <thead>
        <tr class="w3-dark-grey">
            <td>Include / Exclude</td>
            <td>User Types</td>
            <td>Membership Kind</td>
            <td>Guests or Tenants</td>
        </tr>
    </thead>
'@

if ($out_includeGuestOrExternalUserTypes -or $out_includeMembershipKind -or $out_includeMembers -or $out_excludeGuestOrExternalUserTypes -or $out_excludeMembershipKind -or $out_excludeMembers) {
$outHtml = $outHtml + @'
    <tbody>
        <tr>
        <td>Included</td>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        </tr>
        <tr>
        <td>Excluded</td>
        <td>{3}</td>
        <td>{4}</td>
        <td>{5}</td>
        </tr>
    </tbody>
'@ -f $out_includeGuestOrExternalUserTypes, $out_includeMembershipKind, $out_includeMembers, $out_excludeGuestOrExternalUserTypes, $out_excludeMembershipKind, $out_excludeMembers
} else {
    $outHtml = $outHtml + '<tr><td></td><td></td><td></td><td></td></tr>'
}

$outHtml = $outHtml + @'
    </table>
</div>
'@


# Grant Controls Section

$out_excludeMembers = ""
if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"]) {
    $out_excludeMembers = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"] -join $separator
}

$out_excludeMembers = ""
if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"]) {
    $out_excludeMembers = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.AdditionalProperties["members"] -join $separator
}

$outHtml = $outHtml + @'
<header class="w3-container w3-margin-top">
    <h2 class="w3-blue">Grant Controls</h2>
</header>

<div class="w3-container">
<table class="w3-table-all w3-hoverable w3-small">
<thead>
    <tr class="w3-dark-grey">
        <td>BuiltInControls</td>
        <td>Operator</td>
        <td>Auth Strength DisplayName</td>
        <td>Auth Strength AllowedCombinations</td>
    </tr>
</thead>
<tbody>
'@

foreach ($grantControl in $policy.GrantControls) {
    $out_builtInControls = ""
    if ($grantControl.BuiltInControls) {
        $out_builtInControls = $grantControl.BuiltInControls -join $separator
    }

    $out_operator = $grantControl.Operator

    $out_displayName = ""
    if ($grantControl.AuthenticationStrength.DisplayName) {
        $out_displayName = $grantControl.AuthenticationStrength.DisplayName
    }

    $out_allowedCombinations = ""
    if ($grantControl.AuthenticationStrength.AllowedCombinations) {
        $out_allowedCombinations = $grantControl.AuthenticationStrength.AllowedCombinations -join $separator
    }
    
    $outHtml = $outHtml + '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f $out_builtInControls,$out_operator, $out_displayName, $out_allowedCombinations
}

$outHtml = $outHtml + @'
</tbody>
</table>
</div>
'@


# Session Controls Section
$out_applicationEnforcedRestrictions = ""
if ($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
    $out_applicationEnforcedRestrictions = $policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
}

$out_cloudSecurityIsEnabled = ""
if ($policy.SessionControls.CloudAppSecurity.IsEnabled) {
    $out_cloudSecurityIsEnabled = $policy.SessionControls.CloudAppSecurity.IsEnabled
}

$out_disableResilienceDefaults = ""
if ($policy.SessionControls.DisableResilienceDefaults) {
    $out_disableResilienceDefaults = $policy.SessionControls.DisableResilienceDefaults
}

$out_persistentBrowserMode = ""
if ($policy.SessionControls.PersistentBrowser.Mode) {
    $out_persistentBrowserMode = $policy.SessionControls.PersistentBrowser.Mode
}

$out_signInFrequency = ""
if ($policy.SessionControls.SignInFrequency.Value -or $policy.SessionControls.SignInFrequency.Type) {
    $out_signInFrequency = ($policy.SessionControls.SignInFrequency.Value | Out-String) + " " + ($policy.SessionControls.SignInFrequency.Type | Out-String)
}


$outHtml = $outHtml + @'
<header class="w3-container w3-margin-top">
    <h2 class="w3-blue">Session Controls</h2>
</header>

<div class="w3-container">
<table class="w3-table-all w3-hoverable w3-small property-value">
    <tr>
        <td class="w3-dark-grey">Application Enforced Restrictions:</td>
        <td>{0}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Cloud App Security:</td>
        <td>{1}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Disable Resilience Defaults:</td>
        <td>{2}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Persistent Browser Mode:</td>
        <td>{3}</td>
    </tr>
    <tr>
        <td class="w3-dark-grey">Sign-in Frequency:</td>
        <td>{4}</td>
    </tr>
</table>
</div>
'@ -f $out_applicationEnforcedRestrictions, $out_cloudSecurityIsEnabled, $out_disableResilienceDefaults, $out_persistentBrowserMode, $out_signInFrequency

$outHtml = $outHtml + @'
<header class="w3-container w3-margin-top">
    <h2 class="w3-blue">PowerShell Output JSON</h2>
</header>

<div class="w3-container">
    <div class="w3-panel w3-card w3-light-grey">
        <p>Conditional Policy Object</p>
        <div class="w3-code w3-small notranslate">
            <pre>
            {0}
            </pre>
        </div>

        <p>Named Locations</p>
        <div class="w3-code w3-small notranslate">
            <pre>
            {1}
            </pre>
        </div>

        <p>Included Group Members</p>
        <div class="w3-code w3-small notranslate">
            <pre>
            {2}
            </pre>
        </div>

        <p>Excluded Group Members</p>
        <div class="w3-code w3-small notranslate">
            <pre>
            {3}
            </pre>
        </div>
    </div>
</div>

</body>
</html>
'@ -f (ConvertTo-JSON -InputObject $policy -Depth 50), (ConvertTo-JSON -InputObject $namedLocations -Depth 50), (ConvertTo-JSON -InputObject $includeGroupsMembers -Depth 50), (ConvertTo-JSON -InputObject $excludeGroupsMembers -Depth 50)

} catch {
    #Throw $_
    $_ | Format-List * -Force | Out-String ; $_.InvocationInfo | Format-List * -Force | Out-String
}

return $outHtml
}

function Get-Oag365Cap-PolicySummary {
    <#
    .SYNOPSIS
    Build object summarising conditional access policy properties

    .DESCRIPTION
    Build object summarising conditional access policy properties. Used when exporting summary of all policies to CSV.

    .PARAMETER policy
    Object containing conditional access policy.

    .EXAMPLE
    Get-Oag365Cap-PolicySummary -policy $policy

    .NOTES
        NAME: Get-Oag365Cap-PolicySummary
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           
        CHANGELOG:
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$policy
    )

    if ($policy.Conditions.SignInRiskLevels) {
        $signInRiskLevels = $policy.Conditions.SignInRiskLevels -join "; "
    } else {
        $signInRiskLevels = $false
    }

    if ($policy.Conditions.UserRiskLevels) {
        $userRiskLevels = $policy.Conditions.UserRiskLevels -join "; "
    } else {
        $userRiskLevels = $false
    }

    if ($policy.Conditions.servicePrincipalRiskLevels) {
        $servicePrincipalRiskLevels = $policy.Conditions.servicePrincipalRiskLevels -join "; "
    } else {
        $servicePrincipalRiskLevels = $false
    }

    if ($policy.Conditions.Applications.IncludeApplications) {
        $appIncluded = $true
    } else {
        $appIncluded = $false
    }

    if ($policy.Conditions.Applications.ExcludeApplications) {
        $appExcluded = $true
    } else {
        $appExcluded = $false
    }

    if ($policy.Conditions.Applications.IncludeUserActions) {
        $appIncludedActions = $true
    } else {
        $appIncludedActions = $false
    }

    if ($policy.Conditions.Applications.IncludeAuthenticationContextClassReferences) {
        $appIncludeAuthContext = $true
    } else {
        $appIncludeAuthContext = $false
    }
    
    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.ClientApplications.IncludeServicePrincipals)) {
        $clientAppIncludeServicePrincipals = $true
    } else {
        $clientAppIncludeServicePrincipals = $false
    }
    
    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.ClientApplications.ExcludeServicePrincipals)) {
        $clientAppExcludeServicePrincipals = $true
    } else {
        $clientAppExcludeServicePrincipals = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.ClientApplications.ServicePrincipalFilter.rule)) {
        $clientAppsServicePrincipalFilter = $true
    } else {
        $clientAppsServicePrincipalFilter = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.conditions.Locations.IncludeLocations)) {
        $locationsIncluded = $true
    } else {
        $locationsIncluded = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.conditions.Locations.ExcludeLocations)) {
        $locationsExcluded = $true
    } else {
        $locationsExcluded = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.Platforms.IncludePlatforms)) {
        $platformsIncluded = $true
    } else {
        $platformsIncluded = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.Platforms.ExcludePlatforms)) {
        $platformsExcluded = $true
    } else {
        $platformsExcluded = $true
    }

    if ($policy.Conditions.Devices.DeviceFilter) {
        $deviceFilter = $true
    } else {
        $deviceFilter = $true
    }

    if ($policy.Conditions.Users.IncludeUsers) {
        $usersIncluded = $true
    } else {
        $usersIncluded = $false
    }

    if ($policy.Conditions.Users.ExcludeUsers) {
        $usersExcluded = $true
    } else {
        $usersExcluded = $false
    }

    if ($policy.Conditions.Users.IncludeGroups) {
        $groupsIncluded = $true
    } else {
        $groupsIncluded = $false
    }

    if ($policy.Conditions.Users.ExcludeGroups) {
        $groupsExcluded = $true
    } else {
        $groupsExcluded = $false
    }

    if ($policy.Conditions.Users.IncludeRoles) {
        $rolesIncluded = $true
    } else {
        $rolesIncluded = $false
    }

    if ($policy.Conditions.Users.ExcludeRoles) {
        $rolesExcluded = $true
    } else {
        $rolesExcluded = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.Users.IncludeGuestsOrExternalUsers.guestOrExternalUserTypes)) {
        $includeGuestsOrExternalUsers = $true
    } else {
        $includeGuestsOrExternalUsers = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.guestOrExternalUserTypes)) {
        $excludeGuestsOrExternalUsers = $true
    } else {
        $excludeGuestsOrExternalUsers = $false
    }

    if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.MembershipKind) {
        $includeExternalTenants = $true
    } else {
        $includeExternalTenants = $false
    }

    if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.MembershipKind) {
        $excludeExternalTenants = $true
    } else {
        $excludeExternalTenants = $false
    }

    if (($policy.GrantControls).BuiltInControls) {
        $builtInControls = $policy.GrantControls.BuiltInControls -join "; "
    } else {
        $builtInControls = $false
    }

    if (($policy.GrantControls).AuthenticationStrength) {
        $authStrength = $policy.GrantControls.AuthenticationStrength.DisplayName
    } else {
        $authStrength = $false
    }

    if ($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
        $appRestrictionsEnforced = $true
    } else {
        $appRestrictionsEnforced = $false
    }

    if ($policy.SessionControls.CloudAppSecurity.IsEnabled) {
        $cloudAppSecEnabled = $true
    } else {
        $cloudAppSecEnabled = $false
    }

    if ($policy.SessionControls.DisableResilienceDefaults) {
        $disableResilienceDefaults = $true
    } else {
        $disableResilienceDefaults = $false
    }

    if ($policy.SessionControls.PersistentBrowser.Mode) {
        $persistentBrowser = $policy.SessionControls.PersistentBrowser.Mode
    } else {
        $persistentBrowser = $false
    }

    if ($policy.SessionControls.SignInFrequency.Value -or $policy.SessionControls.SignInFrequency.Type) {
        $signInFrequency = ($policy.SessionControls.SignInFrequency.Value | Out-String).Trim() + " " + ($policy.SessionControls.SignInFrequency.Type | Out-String).Trim()
    } else {
        $signInFrequency = $false
    }

    [PSCustomObject]@{
        Id = $policy.Id
        DisplayName = $policy.DisplayName
        State = $policy.State
        CreatedDateTime = $policy.CreatedDateTime
        ModifiedDateTime = $policy.ModifiedDateTime
        "Condition: SignInRiskLevels" = $signInRiskLevels
        "Condition: UserRiskLevels" = $userRiskLevels
        "Condition: ServicePrincipalRiskLevels" = $servicePrincipalRiskLevels
        "Condition: ApplicationsIncluded" = $appIncluded
        "Condition: ApplicationsExcluded" = $appExcluded
        "Condition: ApplicationActionsIncluded" = $appIncludedActions
        "Condition: ApplicationAuthContextIncluded" = $appIncludeAuthContext
        "Condition: ClientAppServicePrincipalsIncluded" = $clientAppIncludeServicePrincipals
        "Condition: ClientAppServicePrincipalsExcluded" = $clientAppExcludeServicePrincipals
        "Condition: ClientAppsServicePrincipalFilter" = $clientAppsServicePrincipalFilter
        "Condition: LocationsIncluded" = $locationsIncluded
        "Condition: LocationsExcluded" = $locationsExcluded
        "Condition: PlatformsIncluded" = $platformsIncluded
        "Condition: PlatformsExcluded" = $platformsExcluded
        "Condition: DeviceFilter" = $deviceFilter
        "Condition: UsersIncluded" = $usersIncluded
        "Condition: UsersExcluded" = $usersExcluded
        "Condition: GroupsIncluded" = $groupsIncluded
        "Condition: GroupsExcluded" = $groupsExcluded
        "Condition: RolesIncluded" = $rolesIncluded
        "Condition: RolesExcluded" = $rolesExcluded
        "Condition: GuestsIncluded" = $includeGuestsOrExternalUsers
        "Condition: GuestsExcluded" = $excludeGuestsOrExternalUsers
        "Condition: ExternalTenantsIncluded" = $includeExternalTenants
        "Condition: ExternalTenantsExcluded" = $excludeExternalTenants
        "Grant: BuiltInControls" = $builtInControls
        "Grant: AuthStrength" = $authStrength
        "Session: AppEnforcedRestrictions" = $appRestrictionsEnforced
        "Session: CloudAppSecurity" = $cloudAppSecEnabled
        "Session: DisableResilienceDefaults" = $disableResilienceDefaults
        "Session: PersistentBrowserMode" = $persistentBrowser
        "Session: SignInFrequency" = $signInFrequency
    }
}


function Export-Oag365Cap-Policies {
    <#
    .SYNOPSIS
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .DESCRIPTION
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .PARAMETER OutputFolder
    The folder to create the individual HTML files in. Defined in the $script:exportTargetFolder variable.

    .EXAMPLE
    Export-Oag365Cap-Policies -OutputFolder $OutputFolder

    .NOTES
        NAME: Export-Oag365Cap-Policies
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityConditionalAccessPolicy
           Convert-Oag365Cap-PolicyToHtml
           
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

        $capSummary += Get-Oag365Cap-PolicySummary -policy $policy

        try {
            $capExportFileName = ($policy.DisplayName) -replace '[^a-zA-Z0-9-_ ]', ' '
            $outHtml = Convert-Oag365Cap-PolicyToHtml -policy $policy 
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

function Export-Oag365Cap-AuthMethodPolicy {
    <#
    .SYNOPSIS
    Exports Microsoft Entra authentication method policies

    .DESCRIPTION
    Exports Microsoft Entra authentication method policies

    .EXAMPLE
    ExportM365-AuthPolicy

    .NOTES
        NAME: ExportM365-AuthPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyAuthenticationMethodPolicy

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID tenancy authentication method policies"

        # Get defined authentication method policies
        $getMgPolicyAuthMethodPolicy = Get-MgPolicyAuthenticationMethodPolicy
        $getMgPolicyAuthMethodPolicyCount = $getMgPolicyAuthMethodPolicy.count
        Write-Host " - Found $getMgPolicyAuthMethodPolicyCount M365 authentication policies"
        
        # Export as CSV
        $getMgPolicyAuthMethodPolicy | ForEach-Object { 
            $authMethodPolicy = $_
            ForEach ($item in $_.AuthenticationMethodConfigurations) {
                $item.AdditionalProperties | ForEach-Object {
                    foreach ($key in $_.Keys) {
                        [PSCustomObject]@{
                            "Id"                     = $authMethodPolicy.Id
                            "DisplayName"            = $authMethodPolicy.DisplayName
                            "LastModifiedDateTime"   = $authMethodPolicy.LastModifiedDateTime
                            "PolicyMigrationState"   = $authMethodPolicy.PolicyMigrationState
                            "PolicyVersion"          = $authMethodPolicy.PolicyVersion
                            "AuthMethodId"           = $item.Id
                            "AuthMethodState"        = $item.State
                            "AUthMethodSetting"      = $key
                            "AuthMethodSettingValue" = ($_[$key] | ConvertTo-Json)
                        }
                    }
                }
            }
        } | Export-CSV $script:exportTarget.authMethodPolicy -NoTypeInformation
        Write-Host " - Exported M365 authentication methods policies to CSV $($script:exportTarget.authMethodPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving Entra authentication method policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Cap-AuthStrengthPolicy {
    <#
    .SYNOPSIS
    Exports Microsoft Entra authentication strength policies

    .DESCRIPTION
    Exports Microsoft Entra authentication strength policies

    .EXAMPLE
    Export-Oag365Cap-AuthStrengthPolicy

    .NOTES
        NAME: Export-Oag365Cap-AuthStrengthPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyAuthenticationStrengthPolicy

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID tenancy authentication strength policies"

        # Get defined authentication strength policies
        $getMgPolicyAuthStrengthPolicy = Get-MgPolicyAuthenticationStrengthPolicy -All
        $getMgPolicyAuthStrengthPolicyCount = $getMgPolicyAuthStrengthPolicy.count
        Write-Host " - Found $getMgPolicyAuthStrengthPolicyCount M365 authentication policies"

        $getMgPolicyAuthStrengthPolicy | Select-Object -Property Id, DisplayName, Description, PolicyType, RequirementsSatisfied, @{l='AllowedCombinations'; e={$_.AllowedCombinations | Out-String }}, @{l='CombinationConfigurations'; e={$_.CombinationConfigurations | Out-String }}, CreatedDateTime, ModifiedDateTime | Export-CSV $script:exportTarget.authStrengthPolicy -NoTypeInformation
        Write-Host " - Exported M365 authentication strength policies to CSV $($script:exportTarget.authStrengthPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving Entra authentication strength policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
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
		"conditionalAccess" = New-Item -ItemType Directory -Force -Path "$($script:output)\M365\$($script:execDateTime)\ConditionalAccess"
	}

    # Define where exports will be saved
    $script:exportTarget = [pscustomobject]@{
        "authMethodPolicy"          = Join-Path -Path $script:exportTargetFolder.conditionalAccess -ChildPath "authMethodPolicies.csv"
        "authStrengthPolicy"        = Join-Path -Path $script:exportTargetFolder.conditionalAccess -ChildPath "authStrengthPolicies.csv"
        "capSummary"                = Join-Path -Path $script:exportTargetFolder.conditionalAccess -ChildPath "capSummary.csv"
    }

    $script:psGraphModulesRequired = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DirectoryObjects',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Applications'
    )


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
            Write-Host "   - Found required PowerShell module '$_' [$moduleVersionNo]. " -ForegroundColor Green
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
    Write-Host "`nExport Microsoft Entra conditional access policies ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

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
        # Get conditional access policies, parse and output as HTML reports
        Export-Oag365Cap-Policies -OutputFolder $script:exportTargetFolder.conditionalAccess

        # Export Authentication Method and Strength Policies
        Export-Oag365Cap-AuthMethodPolicy | Out-Null
        Export-Oag365Cap-AuthStrengthPolicy | Out-Null
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCh5xzykw3Ru/aj
# RtufMPhbvK1mw9R7mlU7clHjY6BQuKCCEiAwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# CQQxIgQgTJTad8D0m3n5y/zd50apvOLOopb+2WRyYkqeiK9/MmcwDQYJKoZIhvcN
# AQEBBQAEggIARGHEg4y4GbjMv+TyrosA19tBrbyw+uVoAKMy/y4KTmUVdXABK4FE
# JEXQxCxqB3Kw2pzAo9oWGzO4rkhpYV3hurNQ8Nmpg/VyK32TgxWqaFGf+93hdROU
# UZ+FbmvNOI8RdWvxcpHUBeN0wJclu+rLBDrGD0JND4Xy6R+0sxGhoYosTkRk9McA
# QDqlJw51IWWHUcq9S8il+UrSNt6n506+pYx4Ec7HPcGor3Xq6/DKz5ra7F4fCJun
# O5Vl3Dj/EcRn0Cf5rCg1s6+sJ9wMa8wA2FVtFRGvTAYSDexvljrW7zNZUJjwoG69
# 5jNOj6wWM4uJjJIriGolKcC+A+pXCIundeYO9R/8WAiD0/BZ93PIjgUnPihKMj92
# IBR1IpC/VRkj/8N+Qbc6IHLU6TSUc/zQ5SvkT7yMtZsCSqjqPIAq5xMyzo6K06KX
# 60tmzKafo9dKp3leCgwg/lYqXDq/ceVje5OYIWPk6U6nIMZhyJBmCp03p4Fv/sMI
# ctlgXpMzeoUNTjNZJh6t0j1UZlJ/SlwjWWx6TQCL38xwpPSHXKnyuOFMHR9URQ32
# xVRQHdCiHJ9pd0ktVoMPW5X1pBhFwEHpaPnAvrvW8tZDacoXtVM9OLc2NA6Nl93n
# Z72h8/zt5TsuyBRQlw8iKthwXMLuQ0/0y6ydDWG28li7BTPXprZ8Reg=
# SIG # End signature block
