<#
.SYNOPSIS
OAG export of Microsoft 365 Identity and Access Management data

.DESCRIPTION
Uses the Microsoft.Graph PowerShell modules to export users devices and service principals, user MFA settings and privileged user built-in role memberships.

    Stage 1: Export details of enabled Microsoft Entra users
            Required Scope: MS Graph Application permissions for Directory.Read.All, DeviceManagementApps.Read.All, Device.Read.All, Application.Read.All

             Export details of Microsoft Entra managed devices
            Required Scope: MS Graph Application permissions for Device.Read.All, Directory.Read.All

             Export details of Microsoft Entra Service Principals with application role assignments or oAuth2 grants
            Required Scope: MS Graph Application permissions for Application.Read.All, Directory.Read.All
    
    Stage 2: Export of User MFA status report
            Required Scope: MS Graph Application permissions for AuditLog.Read.All

    Stage 3: Export Microsoft 365 built-in role memberships
            Required Scope: MS Graph Application permissions for Directory.Read.All, RoleManagement.Read.Directory, PrivilegedEligibilitySchedule.Read.AzureADGroup, PrivilegedAssignmentSchedule.Read.AzureADGroup


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
CSV exports containing
  - Entra Users, devices and service principals
  - User MFA settings
  - Built-in role membership
  - PIM request and approval requirements

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
The name of the folder to output exported information to. Defaults to the running scripts folder ($PSScriptRoot).

.EXAMPLE
Export-M365-IamReport.ps1

By default the script will prompt for credentials. It will request the required scopes, running as the 'MS Graph' enterprise application. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-IamReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppSecret "a1bC2d~E3fGh4iJ5kL6mN7oP8qR9sT0uV1wX2yZ3"

The script will use the provided client ID, tenant ID and secret to authenticate as an enterprise application. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-IamReport.ps1 -AppClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c" -AppTenantId "a1b2c3d4-e5f6-7890-abcd-ef0123456789" -AppCertThumbprint "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"

The script will use the provided client ID, tenant ID and authenticate as an enterprise application using a certificate installed on the local machine that has the specified thumbprint. The application must be created and the required scopes granted before running the script. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-IamReport.ps1 -Output C:\temp

Optionally you can change the location the exported files are saved to. E.g. C:\temp\M365\YYYYMMDD_HHIISS\*

.NOTES
VERSION: 3.1

RELEASE APPROVAL: <UNAPPROVED>

2026-05-20
 - Added Enterprise Application authentication and improved documentation

 2026-05-18
 - Added PIM role and group management policy exports that show request / approval settings

2026-05-12
  - Updated to use approved verbs and clearer function names
  - Updated role membership export to add 'OagRightsSummary' field containing summary of rights and their source
  - Updated role membership export to add 'OagRightsFrom' field containing name of group rights are inherited from

2026-05-01: Created as standalone script

Uses Microsoft.Graph PowerShell module:
    # Global
    Connect-MgGraph
    Disconnect-MgGraph
    Get-MgContext
    Get-MgDirectoryObject

    # Stage 1: Users, Devices and Services
    Get-MgUser
    Get-MgDevice
    Get-MgServicePrincipal
    Get-MgServicePrincipalAppRoleAssignment
    Get-MgServicePrincipalOauth2PermissionGrant

    # Stage 2: User MFA
    Get-MgReportAuthenticationMethodUserRegistrationDetail

    # Stage 3: Roles
    Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule
    Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule
    Get-MgRoleManagementDirectoryRoleAssignmentSchedule
    Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance
    Get-MgRoleManagementDirectoryRoleEligibilitySchedule
    Get-MgRoleManagementDirectoryRoleAssignment
    Get-MgGroupTransitiveMember
  
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
    "Directory.Read.All",
    "DeviceManagementApps.Read.All",
    "Device.Read.All", 
    "Application.Read.All",
    "AuditLog.Read.All",
    "RoleManagement.Read.Directory",
    "PrivilegedEligibilitySchedule.Read.AzureADGroup",
    "PrivilegedAssignmentSchedule.Read.AzureADGroup"
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



function Export-Oag365Iam-EntraUsers {
    <#
    .SYNOPSIS
    Exports Microsoft Entra users

    .DESCRIPTION
    Exports Microsoft Entra users

    .EXAMPLE
    Export-Oag365Iam-EntraUsers

    .NOTES
        NAME: Export-Oag365Iam-EntraUsers
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgUser

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID users"

        # Get Entra users
        $users = Get-MgUser -All -Filter 'accountEnabled eq true' -ConsistencyLevel eventual -Property @("DisplayName", "UserPrincipalName", "SignInActivity", "UserType", "Mail", "AssignedLicenses", "AccountEnabled", "EmployeeId", "UsageLocation", "OnPremisesLastSyncDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "LastPasswordChangeDateTime", "PasswordPolicies", "CreatedDateTime", "CreationType", "DeletedDateTime", "id")
        $userCount = $users.count
        Write-Host " - Found $userCount M365 users"

        # Export to CSV
        $users | Select-Object $script:userExportProperties | Export-CSV $script:exportTarget.users -NoTypeInformation
        Write-Host " - Exported M365 users to CSV $($script:exportTarget.users)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving M365 users" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Iam-EntraDevices {
    <#
    .SYNOPSIS
    Exports Microsoft Entra managed devices

    .DESCRIPTION
    Exports Microsoft Entra managed devices

    .EXAMPLE
    Export-Oag365Iam-EntraDevices

    .NOTES
        NAME: Export-Oag365Iam-EntraDevices
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDevice

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID devices"

        # Get Entra devices
        $getMgDevice = Get-MgDevice -All
        $getMgDeviceCount = $getMgDevice.count
        Write-Host " - Found $getMgDeviceCount Entra device records"

        # Export to CSV
        $getMgDevice | Select-Object -Property $script:deviceExportProperties | Export-CSV $script:exportTarget.entraDevices -NoTypeInformation
        Write-Host " - Exported Entra devices to CSV $($script:exportTarget.entraDevices)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving M365 devices" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }

}

function Export-Oag365Iam-EntraServicePrincipals {
    <#
    .SYNOPSIS
    Exports Microsoft Entra service principals and their assigned rights

    .DESCRIPTION
    Exports Microsoft Entra service principals and their assigned rights from any application role assignments or oAuth grants.

    .EXAMPLE
    Export-Oag365Iam-EntraServicePrincipals

    .NOTES
        NAME: Export-Oag365Iam-EntraServicePrincipals
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgServicePrincipal
           Get-MgServicePrincipalAppRoleAssignment
           Get-MgServicePrincipalOauth2PermissionGrant
           Get-MgDirectoryObject
           Find-MgGraphPermission

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID Service Principals"

        $pCache = @{}
        $scopeCache = @{}
        $spCache = @{}
        $returnServicePrincpals = @()
        $servicePrincipals = Get-MgServicePrincipal -All
        $spCount = $servicePrincipals.Count
        $spProcessed = 0
        Write-Host " - Found $($spCount) service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))"

        $servicePrincipals | ForEach-Object {
            $sp = $_
        
            Write-Verbose " - Processing assigned app roles and permission grants for $($sp.DisplayName) ($($sp.Id))"

            $spProcessed++
            $completed = [math]::Ceiling(($spProcessed / $spCount) * 100)
            Write-Progress -Activity "Processing $spProcessed of $spCount '$($sp.DisplayName)'" -PercentComplete $completed 

            # Collect roles assigned to this service principal
            $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -All -ServicePrincipalId $_.Id
            $appRolesCount = $appRoleAssignments.Count
            $appRolesProcessed = 0
            if ($appRolesCount -ge 1) {
                Write-Host "    + Found $($appRolesCount) application role assignments for $($($sp.DisplayName))"
            }

            # Add roles to assigned permissions list using custom object
            foreach ($appRoleAssignment in $appRoleAssignments) {
                $appRolesProcessed++
                $completedAppRoles = [math]::Ceiling(($appRolesProcessed / $appRolesCount) * 100)
                Write-Progress -Activity "Processing $($appRoleAssignment.PrincipalDisplayName) <-> $($appRoleAssignment.ResourceDisplayName): $appRolesProcessed of $appRolesCount" -PercentComplete $completedAppRoles 

                Write-Verbose "    + Documenting app role assignments for $appRoleAssignment "

                if (!$spCache.ContainsKey($appRoleAssignment.ResourceId)) { 
                    $spCache[$appRoleAssignment.ResourceId] = Get-MgServicePrincipal -ServicePrincipalId $appRoleAssignment.ResourceId
                    Write-Verbose "    + Retrieving service principal details"
                } else { 
                    Write-Verbose "    + Using cache for service principal details"
                }

                $role = $spCache[$appRoleAssignment.ResourceId].AppRoles | Where-Object Id -eq $appRoleAssignment.AppRoleId
                $item = [pscustomobject]@{
                        'ServicePrincipalId' = $sp.Id
                        'ServicePrincipalDisplayName' = $sp.DisplayName
                        'TypeOfRights' = 'App Role Assignment'
                        'ServicePrincipalType' = $sp.ServicePrincipalType
                        'ServicePrincipalAccountEnabled' = $sp.AccountEnabled
                        'ServicePrincipalDeletedDateTime' = $sp.DeletedDateTime

                        'AppId' = $sp.AppId
                        'AppDisplayName' = $sp.AppDisplayName

                        'ResourceId' = $appRoleAssignment.ResourceId
                        'ResourceDisplayName' = $appRoleAssignment.ResourceDisplayName

                        'ConsentType' = $appRoleAssignment.ConsentType

                        'PrincipalId' = $appRoleAssignment.PrincipalId
                        'PrincipalDisplayName' = $appRoleAssignment.PrincipalDisplayName
                        'PrincipalType' = $appRoleAssignment.PrincipalType

                        'Permission' = $role.Value
                        'PermissionDesc' = "$($role.DisplayName) - $($role.Description)"
                        'PermissionEnabled' = $role.IsEnabled

                        'PermissionType' = 'AppRole'
                        'PermissionType Identifier' = $appRoleAssignment.AppRoleId
                        'PermissionType DeletedDateTime' = $appRoleAssignment.DeletedDateTime
                        'PermissionType CreatedDateTime' = $appRoleAssignment.CreatedDateTime
                    }

                $returnServicePrincpals += $item
            }

            # Get oAuth2 permission grants for this service principal
            $grants = Get-MgServicePrincipalOauth2PermissionGrant -All -ServicePrincipalId $_.Id
            $grantsCount = $grants.Count
            $grantsProcessed = 0
            if ($grantsCount -ge 1) {
                Write-Host "    + Found $($grantsCount) permission grants for $($($sp.DisplayName))"
            }

            # Add grants to assigned permissions list using custom object
            foreach ($grant in $grants) {
                $grantsProcessed++
                $completedGrants = [math]::Ceiling(($grantsProcessed / $grantsCount) * 100)
                Write-Progress -Activity "Processing $($sp.AppDisplayName) <-> $($resource.DisplayName): $grantsProcessed of $grantsCount" -PercentComplete $completedGrants 

                if (!$spCache.ContainsKey($grant.ResourceId)) { 
                    $spCache[$grant.ResourceId] = Get-MgServicePrincipal -ServicePrincipalId $grant.ResourceId
                }

                if ($grant.PrincipalId) {
                    if (!$pCache.ContainsKey($grant.PrincipalId)) { 
                        $pCache[$grant.PrincipalId] = Get-MgDirectoryObject -DirectoryObjectId $grant.PrincipalId | 
                            Select-Object -Property Id, @{l='displayName'; e={$_.AdditionalProperties["displayName"]}}, @{l='userPrincipalName'; e={$_.AdditionalProperties["userPrincipalName"]}}, @{l='principalType'; e={$_.AdditionalProperties["@odata.type"]}}
                    }

                    $principal = $pCache[$grant.PrincipalId]
                } else {
                    $principal = @{}
                }

                $resource = $spCache[$grant.ResourceId]
                foreach ($scope in $grant.Scope.split(' ')) {
                    if ($scope -ne '') {                
                        $role = $resource.AppRoles | Where-Object Value -eq $scope
                        Write-Verbose "    + Documenting Oauth2 permission grants for $role in $scope "

                        if (!$scopeCache.ContainsKey($scope)) {
                            $scopeCache[$scope] = Find-MgGraphPermission -ExactMatch $scope -PermissionType Delegated -ErrorAction SilentlyContinue
                            Write-Verbose "    + Retrieving permission details for $scope"
                        } else { 
                            Write-Verbose "    + Using cache for permission details for $scope "
                        }
                        $permissions = $scopeCache[$scope]                

                        $item = [pscustomobject]@{
                                'ServicePrincipalId' = $sp.Id
                                'ServicePrincipalDisplayName' = $sp.DisplayName
                                'TypeOfRights' = 'Oauth2 Permission Grant'
                                'ServicePrincipalType' = $sp.ServicePrincipalType
                                'ServicePrincipalAccountEnabled' = $sp.AccountEnabled
                                'ServicePrincipalDeletedDateTime' = $sp.DeletedDateTime
        
                                'AppId' = $sp.AppId
                                'AppDisplayName' = $sp.AppDisplayName
            
                                'ResourceId' = $grant.ResourceId
                                'ResourceDisplayName' = $resource.DisplayName
        
                                'ConsentType' = $grant.ConsentType
        
                                'PrincipalId' = $grant.PrincipalId
                                'PrincipalDisplayName' = $principal.displayName
                                'PrincipalType' = $principal.principalType
        
                                'Permission' = $scope
                                'PermissionDesc' = $permissions.Description
                                'PermissionEnabled' = 'N/A'

                                'PermissionType' = $permissions.PermissionType    
                                'PermissionType Identifier' = $permissions.Id
                                'PermissionType DeletedDateTime' = $grant.DeletedDateTime
                                'PermissionType CreatedDateTime' = 'N/A'
                            }
                        $returnServicePrincpals += $item    
                    }
                }
            }
        }

        Write-Progress -Completed -Activity "Processing complete"

        # Export collected rights to CSV
        $returnServicePrincpals | Export-CSV $script:exportTarget.servicePrincipals -NoTypeInformation
        Write-Host " - Exported M365 service principals to CSV $($script:exportTarget.servicePrincipals)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving M365 service principals" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Iam-EntraUserAuthMethods {
    <#
    .SYNOPSIS
    Returns Microsoft Entra MFA settings for all users

    .DESCRIPTION
    Creates a Microsoft Entra user registration details report for each user including and SSPR, MFA settings and preferences. Returns details for all users.

    .EXAMPLE
    Export-Oag365Iam-EntraUserAuthMethods

    .NOTES
        NAME: Export-Oag365Iam-EntraUserAuthMethods
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgReportAuthenticationMethodUserRegistrationDetail

        CHANGELOG:    
    #>

    try {
        Write-Host "`nEntra ID user MFA registration report"

        # Get user MFA regsistration report
        $UserMFA = Get-MgReportAuthenticationMethodUserRegistrationDetail -All:$true -Property * |
            Select-Object UserDisplayName, UserPrincipalName, UserType, `
                IsAdmin, IsMfaCapable, IsMfaRegistered, IsPasswordlessCapable, IsSsprCapable, IsSsprEnabled, IsSsprRegistered, IsSystemPreferredAuthenticationMethodEnabled, `
                @{n='MethodsRegistered';e={($_.MethodsRegistered -join ", ")}}, `
                @{n='SystemPreferredAuthenticationMethods';e={($_.SystemPreferredAuthenticationMethods -join ", ")}}, @{n='UserPreferredMethodForSecondaryAuthentication';e={($_.UserPreferredMethodForSecondaryAuthentication -join ", ")}}, `
                LastUpdatedDateTime, Id
        Write-Host "Found MFA settings for $(($UserMFA | Measure-Object).Count) users"

        # Export to CSV
        $UserMFA | Export-Csv -Path $script:exportTarget.userMfa -NoTypeInformation -Append
        Write-Host "Exported MFA settings to $($script:exportTarget.userMfa)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving Entra user MFA report" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}



function Export-Oag365Iam-EntraAdminUnits {
    <#
    .SYNOPSIS
    Export any administrative units used to manage role allocations

    .DESCRIPTION
    Export any administrative units used to manage role allocations

    .EXAMPLE
    Export-Oag365Iam-EntraAdminUnits

    .NOTES
        NAME: Export-Oag365Iam-EntraAdminUnits
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           
        CHANGELOG:
    #>

    try {
        Write-Host "Entra Administrative Units"

        # Get all Entra role definitions to resolve IDs
        $adminUnits =  Get-MgDirectoryAdministrativeUnit -All
        Write-Host " - Found $(($adminUnits | Measure-Object).count) admin units"

        # Export to CSV
        $adminUnits | Select-Object -Property Id, DisplayName, Description, IsMemberManagementRestricted, MembershipRule, MembershipRuleProcessingState, MembershipType, Visibility, Extensions, @{l='AdditionalProperties'; e={$_.AdditionalProperties | ConvertTo-Json }} | Export-CSV $script:exportTarget.roleAdminUnits -NoTypeInformation
        Write-Host " - Exported M365 admin units to CSV $($script:exportTarget.roleAdminUnits)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving admin units" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}




function Get-Oag365Iam-PimRoleAssignments{
    <#
    .SYNOPSIS
    Return the requested Entra PIM role assignment data

    .DESCRIPTION
    Return the requested Entra PIM role assignment data

    .PARAMETER type
    The type of object to return. Either activated or schedule

    .EXAMPLE
    Get-Oag365Iam-PimRoleAssignments -type "activated"

    .NOTES
        NAME: ExportM365-Pim-RoleAssignments
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRoleManagementDirectoryRoleAssignmentSchedule
           Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance

        CHANGELOG:    
    #>

	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('schedule', 'activated')]
        [string]$type
    )

	process {
		try {
			switch ($type) {
				'schedule' {
					$assignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -All -ErrorAction Stop
				}
				'activated' {
					$assignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -Filter "AssignmentType eq 'Activated'" -ErrorAction Stop
				}
			}
			
			Write-Host " - Lookup: Found $(($assignments | Measure-Object).Count) PIM role assignment schedules"
			return $assignments	
		} catch {
            $exception = Format-Oag365-Exception -message "Failed retrieving PIM role assignment schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM role assignment schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}

function Get-Oag365Iam-PimGroupAssignments {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group assignment data

    .DESCRIPTION
    Return the requested Entra PIM group assignment data

    .PARAMETER groupId
    The ID of the group to return assignments for

    .EXAMPLE
    Get-Oag365Iam-PimGroupAssignments -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Get-Oag365Iam-PimGroupAssignments
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule

        CHANGELOG:    
    #>
	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [string]$groupId
    )

	process {
		try {
			$assignments = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule -All -Filter "groupId eq '$($groupId)'" -ErrorAction Stop
			Write-Host " - Lookup: Found $(($assignments | Measure-Object).Count) PIM group assignment schedules"
            return $assignments
		} catch {
            $exception = Format-Oag365-Exception -message "Failed retrieving PIM group assignment schedules"  -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM group eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}


function Get-Oag365Iam-PimRoleEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM role eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM role eligibilitiy data

    .EXAMPLE
    Get-Oag365Iam-PimRoleEligibility

    .NOTES
        NAME: Get-Oag365Iam-PimRoleEligibility
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRoleManagementDirectoryRoleEligibilitySchedule

        CHANGELOG:    
    #>    
	[Cmdletbinding()]param()
	
	process {
		try {
			$eligible = (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty principal,roleDefinition -All -ErrorAction Stop | Select-Object id, principalId, directoryScopeId, roleDefinitionId, status, @{Name = "OagRightsSummary"; Expression = { "Role-Direct-EligibleMember" }}, principal, roleDefinition)
			Write-Host " - Lookup: Found $(($eligible | Measure-Object).Count) PIM role eligibility schedules"
			return $eligible
		} catch {
            $exception = Format-Oag365-Exception -message "Failed retrieving PIM role eligibility schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM role eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}

		return $null
	}
}

function Get-Oag365Iam-PimGroupEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM group eligibilitiy data

    .PARAMETER groupId
    The ID of the group to return eligibilitiy for

    .EXAMPLE
    Get-Oag365Iam-PimGroupEligibility -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: Get-Oag365Iam-PimGroupEligibility
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule

        CHANGELOG:    
    #>       
	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [string]$groupId
    )
	
	process {
		try {
			$eligible = Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule -All -Filter "groupId eq '$($groupId)'" -ExpandProperty principal -ErrorAction Stop
			Write-Host " - Lookup: Found $(($eligible | Measure-Object).Count) PIM group eligibility schedules"
			return $eligible
		} catch {            
            $exception = Format-Oag365-Exception -message "Failed retrieving PIM group eligibility schedules" -exception $_
            if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                Write-Warning "Failed retrieving PIM group eligibility schedules. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
            } else {
                Write-Host $exception -ForegroundColor Red
            }
			return $null
		}
		
		return $null
	}
}



function Get-Oag365Iam-PimRoles {
    <#
    .SYNOPSIS
    Obtain permanent, eligible and active built-in role assignments. Return a custom object containing the assignment details. 

    .DESCRIPTION
    Obtain permanent, eligible and active built-in role assignments. Return a custom object containing the assignment details. 
       Permanently assigned to a built-in role
       Permanently assigned via a group within a built-in role
       Eligible through PIM
       Eligible and currently activated
       Eligible through a PIM group

    .EXAMPLE
    Get-Oag365Iam-PimRoles

    .NOTES
        NAME: Get-Oag365Iam-PimRoles
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            Get-MgRoleManagementDirectoryRoleAssignment
            ExportM365-Pim-RoleAssignments
            Get-Oag365Iam-PimRoleEligibility
            Get-MgGroupTransitiveMember
            Get-Oag365Iam-PimGroupAssignments
            Get-Oag365Iam-PimGroupEligibility
            Get-Oag365Iam-PimRoleMembers
           
        CHANGELOG:
          2025-12-04: Rewrite of PIM related functions used in Stage 4 (built-in role memberships)
    #>
	[Cmdletbinding()]param ()

    Write-Host "`nEntra membership of built-in and privileged roles"

	# Data retrieval: Get all directory roles
	$pimRoleMembers = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty principal -All
	$rolesDefinition = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty roleDefinition -All 
	$assignmentSchedules = Get-Oag365Iam-PimRoleAssignments -type 'schedule'

	# Data retrieval: For each directory role, add the role definition and assignment schedule properties
	foreach ($role in $pimRoleMembers) {
		$newRoleDefinition = ($rolesDefinition | Where-Object { $_.id -eq $role.id }).roleDefinition
		$role.RoleDefinition = $newRoleDefinition

        if (-not $script:pimMgmtPolicyRolesCache[$role.roleDefinitionId]) {
            $roleMgmtPolicies = Set-Oag365Iam-PimRoleMgmtPolicies -roleId $role.roleDefinitionId -roleName $newRoleDefinition.DisplayName
        }
		
		if ($null -ne $assignmentSchedules) {
			$assignmentSchedule = $assignmentSchedules | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.principalId -eq $role.PrincipalId) }
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $assignmentSchedule
		} else {
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $null
		}

        Add-Member -InputObject $role -MemberType NoteProperty -Name "OagRightsSummary" -Value "Role-Direct-ActiveMember"        
	}

	# Data retrieval: Add eligible PIM roles to roles array
	$eligibleRoles = Get-Oag365Iam-PimRoleEligibility
	if ($null -ne $eligibleRoles) {
		$pimRoleMembers += $eligibleRoles
	}
    Write-Host " - Found $(($pimRoleMembers | Measure-Object).count) direct members of roles"
    

	# Data retrieval: Identify duplicates in roles array based on active and eligible
	$pimRoleActivations = Get-Oag365Iam-PimRoleAssignments -type 'activated'

	if ($null -ne $pimRoleActivations) {
        if ($null -ne $pimRoleMembers) {
            foreach ($activation in $pimRoleActivations) {
                $pimRoleMembers | Where-Object { ($null -ne $_) -and ($_.id -eq $activation.RoleAssignmentOriginId) -and ($_.Id -eq $activation.Id) } | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "duplicate" -Value $true }
            }
        }
	} else {
        Write-Host " - No PIM role activations found"
    }


    Write-Host " - Identifying inherited permission through the membership of groups assigned to roles"

	# Data retrieval: For each role assigned to a group, build arrays of active and eligible group members
	$pimRoleMembers | Where-Object { $_.Principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } | ForEach-Object {	
        $inheritedActiveRoleMembers = @()
		$inheritedEligibleRoleMembers = @()

        if (-not $script:pimMgmtPolicyGroupsCache[$_.principalId]) {
            $groupMgmtPolicies = Set-Oag365Iam-PimRoleMgmtPolicies -groupId $_.principalId -groupName $_.Principal.AdditionalProperties.displayName
        }

		$activeGroupMembers = Get-MgGroupTransitiveMember -GroupId $_.principalId -All -Property id, displayName, userPrincipalName
		
		$assignmentSchedules = Get-Oag365Iam-PimGroupAssignments -groupId $_.principalId		
		$eligibleRoleMembers = Get-Oag365Iam-PimGroupEligibility -groupId $_.principalId

        Write-Host " + Resolving members of $($_.Principal.AdditionalProperties.displayName) group ($($_.principalId))"

		# Add active role members from group, include active schedule and eligibilty info 
        if ($null -ne $activeGroupMembers) {
            foreach ($member in $activeGroupMembers) {
                Write-Host "      Adding active member $($member.AdditionalProperties.displayName)"
                
                $assignmentSchedule = $null
                $eligibilitySchedule = $null
                $assignmentSchedule = $assignmentSchedules | Where-Object { $_.PrincipalId -eq $member.Id }
                $eligibilitySchedule = $eligibleRoleMembers | Where-Object { $_.PrincipalId -eq $member.Id }
                
                $inheritedActiveRoleMembers += ([PSCustomObject][Ordered]@{
                    "PrincipalType" = $member.AdditionalProperties.'@odata.type'
                    "PrincipalId" = $member.Id
                    "PrincipalDisplayName" = $member.AdditionalProperties.displayName
                    "PrincipalUpn" = $member.AdditionalProperties.userPrincipalName
                    "AssignmentType" = "Active via Group"
                    "AssignmentStartDateTime" = $assignmentSchedule.ScheduleInfo.StartDateTime
                    "AssignmentEndDateTime" = $assignmentSchedule.ScheduleInfo.Expiration.EndDateTime
                    "GroupName" = $_.Principal.AdditionalProperties.displayName
                    "GroupId" = $_.PrincipalId
                    "Status" = $assignmentSchedule.Status
                    "OagRightsSummary" = "Role-Group-ActiveMember"
                    "EligibleStartDateTime" = $eligibilitySchedule.ScheduleInfo.StartDateTime
                    "EligibleEndDateTime" = $eligibilitySchedule.ScheduleInfo.Expiration.EndDateTime
                    "ScheduleExpireType" = $assignmentSchedule.ScheduleInfo.Expiration.Type
                    "ScheduleExpireDuration" = $assignmentSchedule.ScheduleInfo.Expiration.Duration
                    "ScheduleRecurrence" = ($assignmentSchedule.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
                })
            }
        }

		
        # Add eligible role members from PIM role eligibility schedules
        if ($null -ne $eligibleRoleMembers) {
            foreach ($member in $eligibleRoleMembers) {
                Write-Host "      Adding eligible role member $($member.principal.AdditionalProperties.displayName)"

                $inheritedEligibleRoleMembers += ([PSCustomObject][Ordered]@{
                    "PrincipalType" = $member.principal.AdditionalProperties.'@odata.type'
                    "PrincipalId" = $member.principal.Id
                    "PrincipalDisplayName" = $member.principal.AdditionalProperties.displayName
                    "PrincipalUpn" = $member.principal.AdditionalProperties.userPrincipalName
                    "AssignmentType" = "Eligible via Group"
                    "AssignmentStartDateTime" = ""
                    "AssignmentEndDateTime" = ""
                    "GroupName" = $_.Principal.AdditionalProperties.displayName
                    "GroupId" = $_.principalId
                    "Status" = $member.Status
                    "OagRightsSummary" = "Role-Group-EligibleMember"
                    "EligibleStartDateTime" = $member.ScheduleInfo.StartDateTime
                    "EligibleEndDateTime" = $member.ScheduleInfo.Expiration.EndDateTime
                    "ScheduleExpireType" = $member.ScheduleInfo.Expiration.Type
                    "ScheduleExpireDuration" = $member.ScheduleInfo.Expiration.Duration
                    "ScheduleRecurrence" = ($member.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
                })
            }
        }
            
		$_ | Add-Member -MemberType NoteProperty -Name "activeGroupMembers" -Value $inheritedActiveRoleMembers
		$_ | Add-Member -MemberType NoteProperty -Name "eligibleRoleMembers" -Value $inheritedEligibleRoleMembers
	}
	
	$privilegedUsers = Get-Oag365Iam-PimRoleMembers -pimRoleMembers $pimRoleMembers -pimRoleActivations $pimRoleActivations
	return $privilegedUsers
}


function Get-Oag365Iam-PimRoleMembers {
    <#
    .SYNOPSIS
    Build an array of privileged role members from data provided by Get-Oag365Iam-PimRoles

    .DESCRIPTION
    Build an array of privileged role members from data provided by Get-Oag365Iam-PimRoles.

    In addition to the information obtained through the Graph API, add the OAG summary fields:
         OagRightsFrom: Specifies the name of the group the user gets their role membership from 
                        if they are not a direct member of the role.
      OagRightsSummary: Summaries how the user got their rights and if they are active or eligible

    Users granted a role directly (not using PIM) will only be listed once in that role. However, users granted
    a role using PIM will be listed as 'Eligible' and, if activated, a second time as 'Assigned'.
       
    .EXAMPLE
    Get-Oag365Iam-PimRoleMembers

    .NOTES
        NAME: Get-Oag365Iam-PimRoleMembers
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            Get-Oag365Iam-PimRoles (called by, requires data from)
           
        CHANGELOG:
    #>
	[Cmdletbinding()]param (
        [Parameter(Mandatory = $true)]
        [object]$pimRoleMembers,
        [Parameter(Mandatory = $true)][AllowNull()]
        [object]$pimRoleActivations
	)
	
	Write-Host "`nBuilding privileged role members array"
	$privilegedUsers = @()

	foreach ($role in $pimRoleMembers) {	
		if (!$role.status) {
			# Role is permanent. No start/end dates to obtain
			$role | Add-Member -MemberType NoteProperty -Name "scriptedStartDateTime" -Value "Permanent"
			$role | Add-Member -MemberType NoteProperty -Name "scriptedEndDateTime" -Value "Permanent"

			if ($role.assignmentSchedule.AssignmentType) {
				$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value $role.assignmentSchedule.AssignmentType			
			} else {
				$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value "Permanent"			
			}

			if ($role.assignmentSchedule.Status) {
				$role | Add-Member -MemberType NoteProperty -Name "Status" -Value $role.assignmentSchedule.Status
			} else {
				$role | Add-Member -MemberType NoteProperty -Name "Status" -Value "Permanent (non PIM)"
			}
		} else {
			# Role requires activation. Use activation data to determine start/end dates
			if ($role.Principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
				$roleActivation = $pimRoleActivations | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($role.activeGroupMembers.ContainsKey($_.principalId)) -and ($_.MemberType -eq "Group") }
			} else {
				$roleActivation = $pimRoleActivations | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.PrincipalId -eq $role.PrincipalId) }
			}
			
            if ($roleActivation.startDateTime) { 
                $tmpStartDateTime = $roleActivation.startDateTime | Select-Object -Unique | Sort-Object | Select-Object -First 1 
                $tmpAssignmentType = "Eligible (Active)" 
            } else { 
                $tmpStartDateTime = $null 
                $tmpAssignmentType = "Eligible" 
            }

            if ($roleActivation.endDateTime) { 
                $tmpEndDateTime = $roleActivation.endDateTime | Select-Object -Unique | Sort-Object -Descending | Select-Object -First 1 
            } else { 
                $tmpEndDateTime = $null 
            }

			$role | Add-Member -MemberType NoteProperty -Name "scriptedStartDateTime" -Value $tmpStartDateTime
			$role | Add-Member -MemberType NoteProperty -Name "scriptedEndDateTime" -Value $tmpEndDateTime
			$role | Add-Member -MemberType NoteProperty -Name "scriptedAssignmentType" -Value $tmpAssignmentType
		}
		
		switch ($role.principal.AdditionalProperties.'@odata.type') {
			'#microsoft.graph.servicePrincipal' {
				$thisPrincipalId = $role.principal.AdditionalProperties.appId
			}
			
			default {
				$thisPrincipalId = $role.PrincipalId
			}
		}
		
		Write-Host " - $($role.RoleDefinition.displayName): $($role.principal.AdditionalProperties.'@odata.type') $($role.principal.AdditionalProperties.displayName)"

		# Add PIM role assignment data to privileged users array
		$privilegedUsers += [PSCustomObject]@{		
			"PrincipalType" = $role.principal.AdditionalProperties.'@odata.type'
			"PrincipalId" = $thisPrincipalId
			"PrincipalDisplayName" = $role.principal.AdditionalProperties.displayName
			"PrincipalUpn" = $role.principal.AdditionalProperties.userPrincipalName
			"RoleDisplayName" = $role.RoleDefinition.displayName
			"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
			"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
			"OagRightsFrom" = ""
            "OagRightsSummary" = $role.OagRightsSummary
			"AssignmentType" = $role.scriptedAssignmentType
			"AssignmentStart" = $role.scriptedStartDateTime
			"AssignmentEnd" = $role.scriptedEndDateTime
			"EligibleStartDateTime" = ""
			"EligibleEndDateTime" = ""
			"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
			"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
			"RoleTemplateId" = $role.RoleDefinition.templateId
			"AssignmentStatus" = $role.Status
			"AssignmentAppScope" = $role.AppScope.DisplayName
			"AssignmentAppScopeId" = $role.AppScope.Id
			"AssignmentDirectoryScope" = $role.directoryScopeId
			"ScheduleExpireType" = $role.assignmentSchedule.ScheduleInfo.Expiration.Type
			"ScheduleExpireDuration" = $role.assignmentSchedule.ScheduleInfo.Expiration.Duration
			"ScheduleRecurrence" = ($role.assignmentSchedule.ScheduleInfo.Recurrence | ConvertTo-Json -Compress)
		}

		# If active PIM group members have been defined for this role, add to privileged users array.
		$role.activeGroupMembers | ForEach-Object {
			foreach ($activeMember in $_) {
				Write-Host "     + Active PIM role member $($activeMember.PrincipalDisplayName)"
		
				if ($role.AppScope.DisplayName) { $appScope = "$($role.AppScope.DisplayName) *" } else { $appScope = "" }			
				if ($role.AppScope.Id) { $appScopeId = "$($role.AppScope.Id) *" } else { $appScopeId = "" }
				if ($role.directoryScopeId) { $dirScope = "$($role.directoryScopeId) *" } else { $dirScope = "" }			
			
				if ($activeMember.Status) { $status = $activeMember.Status } else { $status = $role.Status }
					
				$privilegedUsers += [PSCustomObject]@{
					"PrincipalType" = $activeMember.PrincipalType
					"PrincipalId" = $activeMember.PrincipalId
					"PrincipalDisplayName" = $activeMember.PrincipalDisplayName
					"PrincipalUpn" = $activeMember.PrincipalUpn
					"RoleDisplayName" = $role.RoleDefinition.displayName
					"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
					"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
					"OagRightsFrom" = $activeMember.GroupName
                    "OagRightsSummary" = $activeMember.OagRightsSummary
					"AssignmentType" = $activeMember.AssignmentType
					"AssignmentStart" = $activeMember.AssignmentStartDateTime
					"AssignmentEnd" = $activeMember.AssignmentEndDateTime
					"EligibleStartDateTime" = ""
					"EligibleEndDateTime" = ""
					"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
					"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
					"RoleTemplateId" = $role.RoleDefinition.templateId
					"AssignmentStatus" = $status
					"AssignmentAppScope" = $appScope
					"AssignmentAppScopeId" = $appScopeId
					"AssignmentDirectoryScope" = $dirScope
					"ScheduleExpireType" = $activeMember.ScheduleExpireType
					"ScheduleExpireDuration" = $activeMember.ScheduleExpireDuration
					"ScheduleRecurrence" = $activeMember.ScheduleRecurrence
				}
			}
		}
			
		# If eligible PIM group members have been defined for this role, add to privileged users array.
		$role.eligibleRoleMembers | ForEach-Object {
			foreach ($eligibleMember in $_) {
				Write-Host "     + Eligible PIM Role Member $($eligibleMember.PrincipalDisplayName)"
				
				if ($role.AppScope.DisplayName) { $appScope = "$($role.AppScope.DisplayName) *" } else { $appScope = "" }			
				if ($role.AppScope.Id) { $appScopeId = "$($role.AppScope.Id) *" } else { $appScopeId = "" }
				if ($role.directoryScopeId) { $dirScope = "$($role.directoryScopeId) *" } else { $dirScope = "" }			
				
				$privilegedUsers += [PSCustomObject]@{
					"PrincipalType" = $eligibleMember.PrincipalType
					"PrincipalId" = $eligibleMember.PrincipalId
					"PrincipalDisplayName" = $eligibleMember.PrincipalDisplayName
					"PrincipalUpn" = $eligibleMember.PrincipalUpn
					"RoleDisplayName" = $role.RoleDefinition.displayName
					"RoleIsBuiltIn" = $role.RoleDefinition.isBuiltIn
					"RoleIsEnabled" = $role.RoleDefinition.IsEnabled
					"OagRightsFrom" = $eligibleMember.GroupName
                    "OagRightsSummary" = $eligibleMember.OagRightsSummary
					"AssignmentType" = $eligibleMember.AssignmentType
					"AssignmentStart" = ""
					"AssignmentEnd" = ""
					"EligibleStartDateTime" = $eligibleMember.EligibleStartDateTime
					"EligibleEndDateTime" = $eligibleMember.EligibleEndDateTime
					"RoleInheritsPermissionsFrom" = $role.RoleDefinition.InheritsPermissionsFrom
					"RoleRscScopes" = ($role.RoleDefinition.ResourceScopes -join "; ")
					"RoleTemplateId" = $role.RoleDefinition.templateId
					"AssignmentStatus" = $eligibleMember.Status
					"AssignmentAppScope" = $appScope
					"AssignmentAppScopeId" = $appScopeId
					"AssignmentDirectoryScope" = $dirScope
					"ScheduleExpireType" = $eligibleMember.ScheduleExpireType
					"ScheduleExpireDuration" = $eligibleMember.ScheduleExpireDuration
					"ScheduleRecurrence" = $eligibleMember.ScheduleRecurrence
				}
			}
		}
	}
	
	return $privilegedUsers
}



function Get-Oag365Iam-PimRoleMgmtPolicyRules {
    <#
    .SYNOPSIS
    Return the requested Entra PIM management policy assignments 

    .DESCRIPTION
    Return the requested Entra PIM management policy assignments 

    .PARAMETER rolePolicyAssignment
    The results of the policy look up in Get-Oag365Iam-PimRoleMgmtPolicyAssignment.

    .EXAMPLE
    Get-Oag365Iam-PimRoleMgmtPolicyRule -rolePolicyAssignment $rolePolicyAssignment

    .NOTES
        NAME: Get-Oag365Iam-PimRoleMgmtPolicyRules
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgPolicyRoleManagementPolicyRule

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true)][object]$rolePolicyAssignment
    )

    $rules = @()
    $rolePolicyAssignment | ForEach-Object {
        $rule = Get-MgPolicyRoleManagementPolicyRule -UnifiedRoleManagementPolicyId $_.PolicyId -All
        $rules += $rule
    }

    return $rules
}


function Set-Oag365Iam-PimRoleMgmtPolicies {
    <#
    .SYNOPSIS
    Set the Entra PIM management policies for either a role or group

    .DESCRIPTION
    Adds the provided Entra PIM management policies for either a role or group to the PIM role management policies cache object.

    .PARAMETER roleId
    The ID of the role to lookup management policies for

    .PARAMETER roleName
    The display name of the role - for ease of log reading.

    .PARAMETER groupId
    The ID of the group to lookup management policies for

    .PARAMETER groupName
    The display name of the group - for ease of log reading.

    .EXAMPLE
    Set-Oag365Iam-PimRoleMgmtPolicies -roleId $roleId -roleName $roleDisplayName

    .EXAMPLE
    Set-Oag365Iam-PimRoleMgmtPolicies -groupId $groupId -groupName $groupDisplayName

    .NOTES
        NAME: Set-Oag365Iam-PimRoleMgmtPolicies
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-Oag365Iam-PimRoleMgmtPolicyRules
           Format-Oag365Iam-PimRoleMgmtPolicies
           Get-MgPolicyRoleManagementPolicyAssignment
           Get-MgPolicyRoleManagementPolicyRule

        CHANGELOG:    
    #>
    [Cmdletbinding(DefaultParameterSetName = 'ByRole')]param (
        [Parameter(Mandatory = $true, ParameterSetName='ByRole')][object]$roleId,
        [Parameter(Mandatory = $true, ParameterSetName='ByRole')][object]$roleName,
        [Parameter(Mandatory = $true, ParameterSetName='ByGroup')][object]$groupId,
        [Parameter(Mandatory = $true, ParameterSetName='ByGroup')][object]$groupName
    )

    switch ($PSCmdlet.ParameterSetName) {
        'ByRole' {
            Write-Host " - PIM role management policy for $roleName ($roleId) "

            if ($script:pimMgmtPolicyRolesCache[$roleId]) {
                Write-Host "     - Returning cached role management policy results "
                return $script:pimMgmtPolicyRolesCache[$roleId]
            }

            try {
                $assignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '$($roleId)'"  -ExpandProperty "policy" -All -ErrorAction Stop
                Write-Host "     - Found $(($assignments | Measure-Object).Count) DirectoryRole policy assignments "
            } catch {
                $exception = Format-Oag365-Exception -message "Failed retrieving PIM role management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM role management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = Get-Oag365Iam-PimRoleMgmtPolicyRules -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) DirectoryRole policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += Format-Oag365Iam-PimRoleMgmtPolicies -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
                }
            }

            if ($policyRules) {
                $script:pimMgmtPolicyRules += $policyRules
                $script:pimMgmtPolicyRolesCache[$roleId] = $policyRules
            } else {
                Write-Host "     - No role management policy assignments or rules found "
            }

            return $script:pimMgmtPolicyRolesCache[$roleId]
        }
        'ByGroup' {
            Write-Host " - PIM group management policy for $groupName ($groupId) "

            if ($script:pimMgmtPolicyGroupsCache[$groupId]) {
                Write-Host "     - Returning cached group management policy "
                return $script:pimMgmtPolicyGroupsCache[$groupId]
            } 

            try {
                $assignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '$($groupId)' and scopeType eq 'Group'" -ExpandProperty "policy" -All -ErrorAction Stop
                Write-Host "     - Found $(($assignments | Measure-Object).Count) group policy assignments "
            } catch {
                $exception = Format-Oag365-Exception -message "Failed retrieving PIM group management policies" -exception $_
                if ($_.ErrorDetails -match 'ErrorCode: AadPremiumLicenseRequired') {
                    Write-Warning "Failed retrieving PIM group management policies. PIM may not be a licensed feature for this tenant. `n - Error: $($_.Exception.Message)"
                } else {
                    Write-Host $exception -ForegroundColor Red
                }
                return $null
            }

            $policyRules = @()
            $assignments | ForEach-Object {
                $rules = Get-Oag365Iam-PimRoleMgmtPolicyRules -rolePolicyAssignment $_
                Write-Host "     - Found $(($rules | Measure-Object).Count) Group policy rules in policy $($_.PolicyId)"

                if ($null -ne $rules) {
                    $policyRules += Format-Oag365Iam-PimRoleMgmtPolicies -mgmtPolicyRules $rules -mgmtPolicyAssignment $_
                }
            }

            if ($policyRules) {
                $script:pimMgmtPolicyRules += $policyRules
                $script:pimMgmtPolicyGroupsCache[$groupId] = $policyRules
            } else {
                Write-Host "     - No group management policy assignments or rules found "
            }

            return $script:pimMgmtPolicyGroupsCache[$groupId]
        }
    } 
}


function Format-Oag365Iam-PimRoleMgmtPolicies {
    <#
    .SYNOPSIS
    Returns a formatted object with the Entra PIM management policies for output

    .DESCRIPTION
    Format Entra PIM management policies obtained in Set-Oag365Iam-PimRoleMgmtPolicies for output

    .PARAMETER mgmtPolicyAssignment
    The management policy assignment obtained from Get-MgPolicyRoleManagementPolicyAssignment

    .PARAMETER mgmtPolicyRules
    The management policy rules obtained from Get-Oag365Iam-PimRoleMgmtPolicyRules for this assignment

    .EXAMPLE
    Format-Oag365Iam-PimRoleMgmtPolicies -mgmtPolicyRules $mgmtPolicyRules -mgmtPolicyAssignments $mgmtPolicyAssignments

    .NOTES
        NAME: Format-Oag365Iam-PimRoleMgmtPolicies
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Format-Oag365Iam-PimRoleMgmtPolicies

        CHANGELOG:    
    #>
    [Cmdletbinding()]param (
        [Parameter(Mandatory = $true)][object]$mgmtPolicyAssignment,
        [Parameter(Mandatory = $true)][object]$mgmtPolicyRules
    )

    $policyRules = @()

    $mgmtPolicyRules | ForEach-Object {
        $rule = $_

        $policyRules += [PsCustomObject]@{
            "Id" = $mgmtPolicyAssignment.Id
            "PolicyId" = $mgmtPolicyAssignment.PolicyId
            "PolicyDescription" = $mgmtPolicyAssignment.Policy.Description
            "PolicyDisplayName" = $mgmtPolicyAssignment.Policy.DisplayName
            "PolicyIsOrganizationDefault" = $mgmtPolicyAssignment.Policy.isOrganizationDefault
            "PolicyLastModifiedByName" = $mgmtPolicyAssignment.Policy.LastModifiedBy.DisplayName
            "PolicyLastModifiedById" = $mgmtPolicyAssignment.Policy.LastModifiedBy.Id
            "PolicyLastModifiedDateTime" = $mgmtPolicyAssignment.Policy.LastModifiedDateTime
            "PolicyAdditionalProperties" = $mgmtPolicyAssignment.AdditionalProperties | ConvertTo-Json -Depth 50   # Complex
            "RoleDefinitionId" = $mgmtPolicyAssignment.RoleDefinitionId
            "ScopeId" = $mgmtPolicyAssignment.ScopeId
            "ScopeType" = $mgmtPolicyAssignment.ScopeType
            "RuleId" = $rule.Id
            "RuleDataType" = $rule.AdditionalProperties.'@odata.type'
            "RuleTargetCaller" = $rule.Target.Caller
            "RuleTargetLevel" = $rule.Target.Level
            "RuleTargetOperations" = $rule.Target.Operations | ConvertTo-Json -Depth 50     #complex
            "RuleTargetObjects" = $rule.Target.TargetObjects
            "RuleTargetEnforcedSettings" = $rule.Target.EnforcedSettings | ConvertTo-Json -Depth 50   #Complex
            "RuleTargetInheritableSettings" = $rule.Target.InheritableSettings | ConvertTo-Json -Depth 50    #Complex
            "RuleApEnabledRules" = $rule.AdditionalProperties.enabledRules | ConvertTo-Json -Depth 50   # Complex 
            "RuleApIsEnabled" = $rule.AdditionalProperties.isEnabled
            "RuleApIsExpirationRequired" = $rule.AdditionalProperties.isExpirationRequired
            "RuleApMaxDuration" = $rule.AdditionalProperties.maximumDuration   
            "RuleApNotificationType" = $rule.AdditionalProperties.notificationType
            "RuleApReceipientType" = $rule.AdditionalProperties.receipientType
            "RuleApNotificationLevel" = $rule.AdditionalProperties.notificationLevel
            "RuleApIsDefaultRecipientsEnabled" = $rule.AdditionalProperties.isDefaultRecipientsEnabled
            "RuleApNotificationRecipients" = $rule.AdditionalProperties.notificationRecipients | ConvertTo-Json -Depth 50  # Complex 
            "RuleApSettingIsApprovalRequired" = $rule.AdditionalProperties.setting.isApprovalRequired
            "RuleApSettingIsApprovalRequiredForExtension" = $rule.AdditionalProperties.setting.isApprovalRequiredForExtension
            "RuleApSettingIsRequestorJustificationRequired" = $rule.AdditionalProperties.setting.isRequestorJustificationRequired
            "RuleApSettingApprovalMode" = $rule.AdditionalProperties.setting.approvalMode
            "RuleApSettingApprovalStages" = $rule.AdditionalProperties.setting.approvalStages | ConvertTo-Json -Depth 50   #complex
        }
    }

    return $policyRules
}







    # -----------------------------------
    # Define variables required by script
    # -----------------------------------

    Clear-Host

    $script:output = ($output | Resolve-Path).Path
    $script:execDateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")

    # Create folders to store output
	$script:exportTargetFolder = [PSCustomObject]@{
		"iam" = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\IAM"
	}

    # Define where exports will be saved
    $script:exportTarget = [pscustomobject]@{
        "pimMgmtPolicyRules"        = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "pimMgmtPolicyRules.csv"
        "builtInRoleMembers"        = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "builtInRoleMembers.csv"
        "roleAdminUnits"            = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "adminUnits.csv"
        "userMfa"                   = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "userMFA.csv"
        "users"                     = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "users.csv"
        "servicePrincipals"         = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "servicePrincipals.csv"
        "entraDevices"              = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "entraDevices.csv"
    }

    $script:psGraphModulesRequired = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DirectoryObjects',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Identity.Governance'
    )

    $script:userExportProperties = @(
        "DisplayName", 
        "UserPrincipalName", 
        @{n='lastSignInDateTime'; e={[datetime]$_.SignInActivity.lastSignInDateTime}}, 
        @{n='lastSignInRequestId'; e={[datetime]$_.SignInActivity.lastSignInRequestId}}, 
        @{n='lastSuccessfulSignInDateTime'; e={[datetime]$_.SignInActivity.lastSuccessfulSignInDateTime}}, 
        @{n='lastSuccessfulSignInRequestId'; e={[datetime]$_.SignInActivity.lastSuccessfulSignInRequestId}}, 
        @{n='lastNonInteractiveSignInDateTime'; e={[datetime]$_.SignInActivity.lastNonInteractiveSignInDateTime}}, 
        @{n='lastNonInteractiveSignInRequestId'; e={[datetime]$_.SignInActivity.lastNonInteractiveSignInRequestId}}, 
        "UserType", 
        "Mail", 
        "AccountEnabled", 
        "EmployeeId", 
        "UsageLocation", 
        "OnPremisesLastSyncDateTime", 
        "OnPremisesDistinguishedName", 
        "LastPasswordChangeDateTime", 
        "PasswordPolicies", 
        "CreatedDateTime", 
        "CreationType", 
        "DeletedDateTime", 
        "id"
    )

    $script:deviceExportProperties = @(
        "Id", 
        "DisplayName", 
        "DeviceId", 
        "DeviceOwnership", 
        "EnrollmentType", 
        "TrustType", 
        "AccountEnabled", 
        "IsCompliant", 
        "IsManaged", 
        "IsManagementRestricted", 
        "IsRooted", 
        "ManagementType", 
        "Manufacturer", 
        "Model", 
        "OperatingSystem", 
        "OperatingSystemVersion", 
        "ApproximateLastSignInDateTime", 
        "ComplianceExpirationDateTime", 
        "RegistrationDateTime", 
        "DeletedDateTime", 
        "OnPremisesLastSyncDateTime", 
        "OnPremisesSecurityIdentifier", 
        "OnPremisesSyncEnabled", 
        "DeviceCategory", 
        "DeviceMetadata", 
        "DeviceVersion", 
        "EnrollmentProfileName", 
        "Extensions", 
        "MdmAppId", 
        "MemberOf", 
         @{l='PhysicalIds'; e={$_.PhysicalIds -join '; '}}, 
        "ProfileType", 
        "RegisteredOwners", 
        "RegisteredUsers", 
        @{l='SystemLabels'; e={$_.SystemLabels -join '; '}}, 
        "TransitiveMemberOf", 
        @{l='AdditionalProperties'; e={$_.AdditionalProperties | ConvertTo-Json }}
    )
 
    $script:pimMgmtPolicyRolesCache = @{}
    $script:pimMgmtPolicyGroupsCache = @{}
    $script:pimMgmtPolicyRules = @()

    

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
        Write-Host "`nSTAGE 1: Getting M365 users and service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan
        Export-Oag365Iam-EntraUsers | Out-Null
        Export-Oag365Iam-EntraDevices | Out-Null
        Export-Oag365Iam-EntraServicePrincipals | Out-Null


        Write-Host "`nSTAGE 2: Retrieving M365 user MFA settings ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan
        Export-Oag365Iam-EntraUserAuthMethods | Out-Null


        Write-Host "`nSTAGE 3: Retrieving M365 user built-in role memberships ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan
        Export-Oag365Iam-EntraAdminUnits | Out-Null

        Get-Oag365Iam-PimRoles | Export-CSV $script:exportTarget.builtInRoleMembers -NoTypeInformation
        Write-Host "Exported built-in role assignments to $($script:exportTarget.builtInRoleMembers)" -ForegroundColor Green

        $script:pimMgmtPolicyRules | Export-CSV $script:exportTarget.pimMgmtPolicyRules -NoTypeInformation
        Write-Host "Exported PIM role management policy rules to $($script:exportTarget.pimMgmtPolicyRules)" -ForegroundColor Green
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBY+cRa65x1tSbw
# ZmtiEOhNfeS+2vzTgjHSQ8kIVTObraCCEiAwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# CQQxIgQgXN5OzCviRQweObQAdwlKBxCH8QEwfAsZTOGI4r77iC0wDQYJKoZIhvcN
# AQEBBQAEggIApzZyPJ8r/a+Ma2SA5OMj+HyY9pjLl1o/R22kSXqbWU9s4YqF2DH7
# Sy9HwRTOvVRk6gR3TTKGDUXRpS7d3b7Qm52CoB5MzBgBU8U6r+kZvXnrwXAhTYRC
# NeJkXzLWo6hjgYtsGEmS/mvaj2BuR67nNBXACoV87c9itWkbajB681gJOO8nCb//
# HL3akTDzwDI9oFqk8Yed1eVM2uR4311TJoVHQYkKpvyQX9gO8UTX2ZYFoLrPtcJF
# 5EsP32Kv8YXQ4649RFOoSMnNYdlYBcnSJf839+uPObq36/UXxEfkxKtBJb4qhUDz
# 50yZjslS0FhFr2WqrvBAhryYFC53cOeODwuGTT3f+Jiw+7IAyCUgVCV+iX30iy9E
# 6qIiwzP/HhRQ7v5uUgFrByYOQ/DN5VvfhYQV98tILaLxAMNJEMd0gEPMoJXGGHRw
# ukMeoJ75N1XKOB/adh4v4SWnl/k2hyu1M36Oek/3zfFBBjlGGY6P42xSJmZvV/Xg
# 7nGZKcIAOOciWNPPz6AQxzfUY68461Hi/uyDSHQflBtviPGSVWjZlbb1JWzEJ2IO
# j8yF08X1jLBKYxbH10FqydROzqyGjRMzgCer63a8Sysu5QmtKVOec2INJgY8QT0t
# gikl8sSno132FcyZ733R5ca6FSRctp22Y0Q7+Bwxoit0GK99v3nuKro=
# SIG # End signature block
