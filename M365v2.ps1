<#
.SYNOPSIS
OAG export of Microsoft 365 tenancy settings, users devices and service accounts, privileged rights, authentication settings, audit log history and threat protection policies for testing

.DESCRIPTION
Uses the Microsoft.Graph and ExchangeOnline PowerShell modules to export key tenancy settings, users devices and service principals, user MFA settings, privileged user built-in role memberships, and authentication policies, conditional access policies, audit logs and sign-in history, Exchange Online Protection and Defender Threat Policy settings.

    Stage 1: Export of organisation wide tenacy settings and SPF / DKIM / DMARC records
               Graph Scope: User.Read, Organization.Read.All, Directory.Read.All, Domain.Read.All, OnPremDirectorySynchronization.Read.All	
                    Author: OAG

    Stage 2: Export details of enabled Microsoft Entra users
               Graph Scope: Directory.Read.All, DeviceManagementApps.Read.All, Device.Read.All, Application.Read.All
                    Author: OAG

             Export details of Microsoft Entra managed devices
               Graph Scope: Device.Read.All, Directory.Read.All
                    Author: OAG

             Export details of Microsoft Entra Service Principals with application role assignments or oAuth2 grants
               Graph Scope: Application.Read.All, Directory.Read.All
                    Author: OAG
    
    Stage 3: Export of User MFA status report
              Graph Scope: AuditLog.Read.All
                   Author: OAG
                      URL: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.reports/get-mgreportauthenticationmethoduserregistrationdetail

    Stage 4: Export Microsoft 365 built-in role memberships
               Graph Scope: Directory.Read.All, RoleManagement.Read.Directory, PrivilegedEligibilitySchedule.Read.AzureADGroup, PrivilegedAssignmentSchedule.Read.AzureADGroup
                    Author: OAG

    Stage 5: Export Microsoft Entra Conditional Access, Authentication Method and Authentication Strength policies
               Graph Scope: Policy.Read.AuthenticationMethod, Policy.Read.All, AuthenticationContext.Read.All, Policy.Read.ConditionalAccess, RoleManagement.Read.Directory, Application.Read.All, Directory.Read.All
                    Author: OAG

    Stage 6: Export Microsoft Entra Tenant Sign-in Logs (optional)
               Graph Scope: AuditLog.Read.All, Directory.Read.All, IdentityRiskyUser.Read.All, IdentityRiskyServicePrincipal.Read.All
                     Roles: (ExchangeOnline) See below for roles required
                    Author: OAG

    Stage 7: Export Exchange Online Protection and Defender for Office threat protection policies
                     Roles: (ExchangeOnline) See below for roles required
                    Author: OAG

    For the Graph API, the script will request the scopes defined in $script:exportM365MsGraphScopes.
    For ExchangeOnline (stages 6 & 7), the user must have the roles defined below.
        Entra: Global Admin or Global Reader; OR
        Exchange Server: Compliance Management, Hygiene Management, Organization Management, View-Only Organization Management; OR
        Exchange Online: Compliance Management, Delegated Setup, Hygiene Management, Organization Management, View-Only Organization Management
        https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps

.COMPONENT 
Requires PowerShell module Microsoft.Graph 
Requires PowerShell module ExchangeOnline

Specific requirements documented in $script:psGraphModulesRequired (ExportM365-ConnectMgGraph) and $script:psExoModulesRequired (ExportM365-ConnectIPPSSession, ExportM365-ConnectExchangeOnline) 

.INPUTS
Prompted for EntraID credentials with roles required to authorise required scopes

.OUTPUTS
CSV exports and HTML reports of 
  - Entra Users, devices and service principals
  - User MFA settings
  - Built-in role membership 
  - Conditional access, authentication method and authentication strength policies
  - Audit log samples
  - Exchange Online Protection and Defender Threat Protection policies

.PARAMETER Output
The name of the folder to output exported information to. Defaults to the running scripts folder ($PSScriptRoot).

.PARAMETER LogFilter
The type of filter to apply when exporting log entries. Defaults to the top 10,000

.PARAMETER RunAll
Should all stages be run. Defaults to $true. Will be set $false if RunStages are specified.

.PARAMETER RunStages
Run specific stages only

.EXAMPLE
Export-M365.ps1
Export-M365.ps1 -RunAll

By default all stages will be run, results will be saved in a subfolder under the scripts current location. The optional -RunAll parameter is $true by default.

.EXAMPLE
Export-M365.ps1 -Output C:\temp

Optionally change the output location.

.EXAMPLE
Export-M365.ps1 -LogFilter All

Optionally change the log sample filter limit.

.EXAMPLE
Export-M365.ps1 -RunStages 2,3

Optionally run specific stages instead of the default -RunAll.


.NOTES
Updated 2026-02-27
  - Updated Graph module scopes to include PrivilegedAssignmentSchedule.Read.AzureADGroup
  - Forced re-authentication on script execution
  - Included PowerShell module version numbers in output

Updated 2025-11-19
  - Major re-write. Added stages and re-numbered
  - Added additional functions for ExchangeOnlineManagement (additional logging and policies for EOP and DTPP)
  - Added stage 7. Environment wide org settings export and threat policy exports
  - Added additional log exports to stage 6
  - Added M365 devices export to stage 2
  - Updated conditional access policy exports. 
  - Recursive group membership lookup for privileged roles and conditional access policies
  - Added support for running specific stages only using RunStages parameter
  
Updated 2024-04-18
  - Emphasise focus on enabled user accounts only

Updated 2024-04-02
  - Updated MFA status check. Now uses Get-MgReportAuthenticationMethodUserRegistrationDetail
  - Removed UPN filter options incompatible with Get-MgReportAuthenticationMethodUserRegistrationDetail

Updated 2023-02-15
  - Added export of service principals and their app role assignments and oAuth2 grants
  - Microsoft Graph modules now itemised for clarity

Updated 2023-10-31
  - Added contrained language mode check
  - Conditional Access Policy report formating
  
Updated 2023-09-27
  - Added HTML output of conditional access policies
  - Added initial warning splash screen
  - Removed use of Microsoft Graph PowerShell module beta version
  - Removed script require statements

Updated 2023-06-08
 - Changed MFA lookup to use ID not UserPrincipalName

Uses Microsoft.Graph PowerShell module:
    # Global
    Connect-MgGraph
    Disconnect-MgGraph
    Get-MgContext
    Get-MgDirectoryObject

    # Stage 1: Environment
    Get-MgOrganization
    Get-MgDomain
    Get-MgDirectoryOnPremiseSynchronization
    Get-MgDomainFederationConfiguration

    # Stage 2: Users, Devices and Services
    Get-MgUser
    Get-MgDevice
    Get-MgServicePrincipal
    Get-MgServicePrincipalAppRoleAssignment
    Get-MgServicePrincipalOauth2PermissionGrant

    # Stage 3: User MFA
    Get-MgReportAuthenticationMethodUserRegistrationDetail

    # Stage 4: Roles
    Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule
    Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule
    Get-MgRoleManagementDirectoryRoleAssignmentSchedule
    Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance
    Get-MgRoleManagementDirectoryRoleEligibilitySchedule
    Get-MgRoleManagementDirectoryRoleAssignment
    Get-MgGroupTransitiveMember

    # Stage 5: Conditional Access Policies
    Get-MgPolicyAuthenticationMethodPolicy
    Get-MgPolicyAuthenticationStrengthPolicy
    Get-MgIdentityConditionalAccessPolicy
    Get-MgIdentityConditionalAccessNamedLocation
    Get-MgIdentityConditionalAccessAuthenticationContextClassReference
    Get-MgDirectoryRoleTemplate
    Get-MgServicePrincipal

    # Stage 6: Audit logs
    Get-MgAuditLogSignIn
    Get-MgAuditLogProvisioning
    Get-MgAuditLogDirectoryAudit

Uses ExchangeOnlineManagement PowerShell module:
    # Global
    Connect-IPPSSession
    Connect-ExchangeOnline
    Disconnect-ExchangeOnline

    # Stage 6: Audit logs
    Get-AdminAuditLogConfig
    Get-UnifiedAuditLogRetentionPolicy

    # Stage 7: Exchange Online Protection and Defender Threat Policies
    Get-MalwareFilterPolicy
    Get-MalwareFilterRule
    Get-AntiPhishingPolicy
    Get-AntiPhishingRule
    Get-HostedConnectionFilterPolicy
    Get-SafeAttachmentsPolicy
    Get-SafeAttachmentsRule
    Get-SafeLinksPolicy
    Get-SafeLinksRule
    Get-QuarantinePolicy
  
.LINK

#>
[CmdletBinding(DefaultParameterSetName="RunAll")]
Param (
	[Parameter(Mandatory = $false, ParameterSetName = 'RunAll')]
    [Parameter(Mandatory = $false, ParameterSetName = 'RunStages')][String]$Output = $PSScriptRoot,

	[Parameter(Mandatory = $false, ParameterSetName = 'RunAll')]
    [Parameter(Mandatory = $false, ParameterSetName = 'RunStages')][ValidateSet('Top10k','All','None')][AllowNull()][String]$LogFilter = "Top10k",

    ## Set $False if specifying $RunStages
	[Parameter(Mandatory = $false, ParameterSetName = 'RunAll')]
    [Parameter(Mandatory = $false, ParameterSetName = 'RunStages')][Bool]$RunAll = $True,    

	[Parameter(Mandatory = $false, ParameterSetName = 'RunStages')][ValidateSet(1,2,3,4,5,6,7)][int[]]$RunStages
)

begin {

#
# Define required graph API scopes
# Note: does not include requirements for connect-ippssession and connect-exchangeonline also required for stages 5 and 6
#
$script:exportM365MsGraphScopes = @(
    "User.Read.All", 
    "Reports.Read.All", 
    "RoleEligibilitySchedule.Read.Directory", 
    "RoleAssignmentSchedule.Read.Directory", 
    "CrossTenantInformation.ReadBasic.All", 
    "AuditLog.Read.All", 
    "Application.Read.All", 
    "Group.Read.All", 
    "Policy.Read.All", 
    "RoleManagement.Read.Directory", 
    "AdministrativeUnit.Read.All",
    "Organization.Read.All",
    "Device.Read.All",
    "Domain.Read.All",
    "Directory.Read.All",
    "PrivilegedEligibilitySchedule.Read.AzureADGroup",
    "PrivilegedAssignmentSchedule.Read.AzureADGroup",
    "OnPremDirectorySynchronization.Read.All",
    "Policy.Read.AuthenticationMethod",
    "IdentityRiskyUser.Read.All",
    "IdentityRiskyServicePrincipal.Read.All"
)

function ExportM365-Exception {
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
    ExportM365-Exception -message "Failed connecting Graph API to M365 tenant" -exception $e

    .NOTES
        NAME: ExportM365-Exception
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

function ExportM365-ConnectMgGraph {
    <#
    .SYNOPSIS
    Connects PowerShell to M365 Microsoft Graph 

    .DESCRIPTION
    Imports required PowerShell modules, if required, and then connects the Graph PowerShell module.  

    .EXAMPLE
    ExportM365-ConnectMgGraph

    .NOTES
        NAME: ExportM365-ConnectMgGraph
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Connect-MgGraph
           Get-MgContext

        CHANGELOG:
          2025-11-18: Updated in include modules only when required. Part of allowing stages to be run independently.
    #>

    try {
        Write-Host "`nMicrosoft Graph required. Please login if prompted" -ForegroundColor Cyan
        sleep 2

        if (-not (Get-Module -Name $script:psGraphModulesRequired)) {
            Write-Host " - Importing PowerShell Graph modules..." -ForegroundColor Cyan
            sleep 2
            $script:psGraphModulesRequired | Import-Module -Verbose
        }
        
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -ContextScope Process -Scopes $script:exportM365MsGraphScopes    
            Write-Host " - Connecting to Microsoft Graph...`nScopes: $($script:exportM365MsGraphScopes)" -ForegroundColor Cyan
        }

        if ((Get-MgContext -ErrorAction SilentlyContinue)) {
            Write-Host " - Connected to Microsoft Graph PowerShell using $((Get-MgContext).Account) account" -ForegroundColor Cyan
        } else {
            throw "Authentication needed. Failed connecting to Microsoft Graph with scopes `$($script:exportM365MsGraphScopes -join ", ")` "
            pause
            return $false
        }

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed connecting to Microsoft Graph PowerShell" -exception $_
        Write-Host $exception -ForegroundColor Red
        pause
        return $false
    }
}

function ExportM365-ConnectIPPSSession {
    <#
    .SYNOPSIS
    Connects PowerShell to M365 Microsoft Exchange Online (Security and Compliance) 

    .DESCRIPTION
    Imports required PowerShell modules, if required, and then connects the Security and Compliance PowerShell module.  

    .EXAMPLE
    ExportM365-ConnectIPPSSession

    .NOTES
        NAME: ExportM365-ConnectIPPSSession
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Connect-IPPSSession

        CHANGELOG:    
          2025-11-18: Updated in include modules only when required. Part of allowing stages to be run independently.

    #>

    try {
        Write-Host "`nMicrosoft Exchange Online (Security and Compliance) PowerShell required. Please login if prompted" -ForegroundColor Cyan

        sleep 2

        if (-not (Get-Module -Name $script:psExoModulesRequired)) {
            Write-Host " - Importing PowerShell Exchange Online modules..." -ForegroundColor Cyan
            sleep 2
            $script:psExoModulesRequired | Import-Module -Verbose
        }

        Write-Host " - Connecting to Microsoft Security & Compliance PowerShell..." -ForegroundColor Cyan
        Connect-IPPSSession -ShowBanner:$false

        $user = (Get-ConnectionInformation).UserPrincipalName
        if ($user) {
            Write-Host " - Connected to Microsoft Security & Compliance PowerShell using $($user) account" -ForegroundColor Cyan
        }

        return $true
    } catch {
        Write-Warning "If connecting to ExchangeOnline fails with the 'Error Acquiring Token' message, the version of the ExchangeOnlineManagement module may be the issue. Issues have been identified with 3.9.2"
        Write-Warning "   Last confirmed working version: Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.9.0 -Force -AllowClobber -Scope AllUsers"

        $exception = ExportM365-Exception -message "Failed connecting to Microsoft Exchange Online Managment Security & Compliance PowerShell" -exception $_
        Write-Host $exception -ForegroundColor Red
        pause
        return $false
    }
}

function ExportM365-ConnectExchangeOnline {
    <#
    .SYNOPSIS
    Connects PowerShell to M365 Exchange Online Management  

    .DESCRIPTION
    Imports required PowerShell modules, if required, and then connects the Exchange Online Management PowerShell module.  

    .EXAMPLE
    ExportM365-ConnectExchangeOnline

    .NOTES
        NAME: ExportM365-ConnectExchangeOnline
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Connect-ExchangeOnline

        CHANGELOG:    
          2025-11-18: Updated in include modules only when required. Part of allowing stages to be run independently.
    #>

    try {
        Write-Host "`nMicrosoft Exchange Online Management PowerShell required. Please login if prompted" -ForegroundColor Cyan
        sleep 2

        if (-not (Get-Module -Name $script:psExoModulesRequired)) {
            Write-Host " - Importing PowerShell Exchange Online modules..." -ForegroundColor Cyan
            sleep 2
            $script:psExoModulesRequired | Import-Module -Verbose
        }

        Write-Host " - Connecting to Microsoft Exchange Online PowerShell..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

        $user = (Get-ConnectionInformation).UserPrincipalName
        if ($user) {
            Write-Host " - Connected to Microsoft Exchange Online PowerShell using $($user) account" -ForegroundColor Cyan
        }

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed connecting to Microsoft Exchange Online PowerShell" -exception $_ 
        Write-Host $exception -ForegroundColor Red
        pause
        return $false
    }
}

function ExportM365-Disconnect {
    <#
    .SYNOPSIS
    Disconnects from M365 services

    .DESCRIPTION
    Disconnects from Connect-MgGraph and Connect-ExcangeOnline.

    .EXAMPLE
    ExportM365-Disconnect

    .NOTES
        NAME: ExportM365-Disconnect
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Disconnect-MgGraph
           Disconnect-ExchangeOnline

        CHANGELOG:    
    #>

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -ErrorAction SilentlyContinue -Confirm:$false  
    } catch {
        Write-Host "ERROR: Failed disconnecting $($_.Exception.message)" -ForegroundColor Red
    }
}

function ExportM365-Organization {
    <#
    .SYNOPSIS
    Get M365 tenant settings 

    .DESCRIPTION    
    Get M365 tenant settings 

    .EXAMPLE
    ExportM365-Organization

    .NOTES
        NAME: ExportM365-Organization
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgOrganization

        CHANGELOG:    
    #>

    try {
        Write-Host "`nM365 tenant organization settings"

        # Get organization details
        $getMgOrganization = Get-MgOrganization -All
        $getMgOrganizationCount = $getMgOrganizationCount.count
        Write-Host " - Found M365 organizations data"

        # Export to CSV
        $getMgOrganization | Select-Object -Property DisplayName, CountryLetterCode, CreatedDateTime, Id, MobileDeviceManagementAuthority, OnPremisesLastSyncDateTime, OnPremisesSyncEnabled, TenantType, @{l='TechnicalNotificationMails'; e={$_.TechnicalNotificationMails -join '; ' }}  | Export-CSV $script:exportTarget.organisations -NoTypeInformation
        Write-Host " - Exported M365 organizations to CSV $($script:exportTarget.organisations)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed getting M365 organization details" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-OrgOnPremiseSync {
    <#
    .SYNOPSIS
    Get details of Entra ID sync with on-premises AD

    .DESCRIPTION    
    Get details of Entra ID sync with on-premises AD

    .EXAMPLE
    ExportM365-OrgOnPremiseSync

    .NOTES
        NAME: ExportM365-OrgOnPremiseSync
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDirectoryOnPremiseSynchronization

        CHANGELOG:    
    #>

    try {
        Write-Host "`nM365 tenant onPremises AD sync settings"

        # Get onPremise sync status
        $getMgDirectoryOnPremiseSync = Get-MgDirectoryOnPremiseSynchronization 
        $getMgDirectoryOnPremiseSyncCount = $getMgDirectoryOnPremiseSync.count
        Write-Host " - Found $getMgDirectoryOnPremiseSyncCount M365 onPremises sync settings"

        # Export to CSV
        $getMgDirectoryOnPremiseSync | ForEach-Object {
            $onPremisesSync = $_.Features
            $onPremisesSync | Add-Member -MemberType NoteProperty -Name Id -Value $_.Id
            $onPremisesSync
        } | Select-Object -Property * -ExcludeProperty AdditionalProperties | Export-CSV $script:exportTarget.onPremisesSync -NoTypeInformation
        Write-Host " - Exported M365 onPremises sync settings to CSV $($script:exportTarget.onPremisesSync)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed getting M365 directory sync config" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-Domains {
    <#
    .SYNOPSIS
    Get domains configured within M365 tenant 

    .DESCRIPTION    
    Get domains configured within M365 tenant 

    .EXAMPLE
    ExportM365-Domains

    .NOTES
        NAME: ExportM365-Domains
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDomain

        CHANGELOG:    
    #>

    try {
        Write-Host "`nM365 tenant domains"

        # Get organization domains
        $getMgDomain = Get-MgDomain -All
        $getMgDomainCount = $getMgDomain.count
        Write-Host " - Found $getMgDomainCount M365 domains"

        $getMgDomain | Select-Object -Property Id, IsAdminManaged, IsDefault, IsInitial, IsRoot, IsVerified, AuthenticationType, AvailabilityStatus, DomainNameReferences, FederationConfiguration, PasswordNotificationWindowInDays, PasswordValidityPeriodInDays | Export-CSV $script:exportTarget.domains -NoTypeInformation
        Write-Host " - Exported M365 domains to CSV $($script:exportTarget.domains)" -ForegroundColor Green

        return $getMgDomain
    } catch {
        $exception = ExportM365-Exception -message "Failed getting M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-EmailAuth {
    <#
    .SYNOPSIS
    Get email authentication DNS records for M365 tenant domains 

    .DESCRIPTION    
    Get email authentication DNS records for M365 tenant domains 

    .PARAMETER domains
    Iterable object containing domains retrieved from Get-MgDomains

    .EXAMPLE
    ExportM365-EmailAuth -domains $object

    .NOTES
        NAME: ExportM365-EmailAuth
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgDomain
           ExportM365-DomainMx
           ExportM365-DomainSpf
           ExportM365-DomainDkim
           ExportM365-DomainDmarc

        CHANGELOG:    
    #>
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$domains
    )

    try {
        Write-Host "`nEmail authentication DNS settings for M365 domains"

        # Get MX records to confirm a mail server is available
        $getMxDnsRecord = ExportM365-DomainMx -getMgDomains $domains

        # Get SPF DNS record for each M365 domain
        Write-Host "`nRetrieving SPF records for M365 domains"
        $getSpfDnsRecord = ExportM365-DomainSpf -getMgDomains $domains
        $getSpfDnsRecordCount = $getSpfDnsRecord.count
        Write-Verbose " - SPF records for $getSpfDnsRecordCount M365 domains"

        ($getMxDnsRecord + $getSpfDnsRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainSpf -NoTypeInformation
        Write-Host " - Exported SPF records for M365 domains to CSV $($script:exportTarget.domainSpf)" -ForegroundColor Green

        # Get DKIM DNS Record for each M365 domain
        Write-Host "`nRetrieving DKIM records for M365 domains"
        $getDkimSigningRecord = ExportM365-DomainDkim -getMgDomains $domains 
        $getDkimSigningRecordCount = $getDkimSigningRecord.count
        Write-Verbose " - Found DKIM config for $getDkimSigningRecordCount domains"

        ($getMxDnsRecord + $getDkimSigningRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDkim -NoTypeInformation
        Write-Host " - Exported M365 DKIM records for M365 domains to CSV $($script:exportTarget.domainDkim)" -ForegroundColor Green

        # Get DMARC DNS Record for each M365 domain
        Write-Host "`nRetrieving DMARC records for M365 domains"
        $getDmarcDnsRecord += ExportM365-DomainDmarc -getMgDomains $domains
        $getDmarcDnsRecordCount = $getDmarcDnsRecord.count
        Write-Verbose " - DMARC records for $getDmarcDnsRecordCount M365 domains"

        ($getMxDnsRecord + $getDmarcDnsRecord) | Select-Object -Property * | Export-CSV $script:exportTarget.domainDmarc -NoTypeInformation
        Write-Host " - Exported DMARC records for M365 domains to CSV $($script:exportTarget.domainDmarc)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed getting email authentication DNS records for M365 domains" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function ExportM365-DomainMx {
    <#
    .SYNOPSIS
    Returns MX DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns MX DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    ExportM365-DomainMx -getMgDomains $object

    .NOTES
        NAME: ExportM365-DomainMx
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>    
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $getMgDomains | ForEach-Object { 
        try {
            $records = Resolve-DnsName -Type MX -Name $_.Id -ErrorAction SilentlyContinue | Where-Object { $_.NameExchange -ne ""; }
            if ($records) {
                foreach ($record in $records) {
                    Write-Host " - $($_.Id) MX found" -ForegroundColor Green

                    if ([string]::IsNullOrEmpty($record.NameExchange) -eq $false) { 
                        $strings = $record.NameExchange -join '; '
                    } elseif ([string]::IsNullOrEmpty($record.IPAddress) -eq $false) {
                        $strings = $record.IPAddress -join '; '
                    } else {
                        $strings = ""
                    }

                    [PSCustomObject]@{
                        "Domain"  = $record.Name
                        "Type"    = $record.Type
                        "TTL"     = $record.TTL
                        "Section" = $record.Section
                        "Strings" = $strings
                    }
                }
            } else {
                Write-Host " - $($_.Id) MX not found" -ForegroundColor DarkYellow
                [PSCustomObject]@{
                    "Domain"  = $_.Id
                    "Type"    = "MX"
                    "TTL"     = ""
                    "Section" = ""
                    "Strings"   = "No MX record found"
                }
            }
        }
        catch {
            Write-Host (ExportM365-Exception -message " - $($_.Id) MX error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $_.Id
                "Type"    = "MX"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving MX record; $($_.Exception.message)"
            }
        }
    } 
}


function ExportM365-DomainSpf {
    <#
    .SYNOPSIS
    Returns SPF DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns SPF DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    ExportM365-DomainSpf -getMgDomains $object

    .NOTES
        NAME: ExportM365-DomainSpf
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>
    
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $getMgDomains | ForEach-Object { 
        try {
            $records = Resolve-DnsName -Type TXT -Name $_.Id -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=spf1"; }
            if ($records) {
                foreach ($record in $records) {
                    Write-Host " - $($_.Id) SPF found" -ForegroundColor Green
                    [PSCustomObject]@{
                        "Domain"  = $record.Name
                        "Type"    = $record.Type
                        "TTL"     = $record.TTL
                        "Section" = $record.Section
                        "Strings" = $record.Strings -join ';'
                    }
                }
            } else {
                Write-Host " - $($_.Id) SPF not found" -ForegroundColor DarkYellow
                [PSCustomObject]@{
                    "Domain"  = $_.Id
                    "Type"    = "TXT"
                    "TTL"     = ""
                    "Section" = ""
                    "Strings"   = "No SPF record found"
                }
            }
        }
        catch {
            Write-Host (ExportM365-Exception -message " - $($_.Id) SPF error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $_.Id
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving SPF record; $($_.Exception.message)"
            }
        }
    } 
}

function ExportM365-DomainDkim {
    <#
    .SYNOPSIS
    Returns DKIM DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns DKIM DNS records for domains obtained from Get-MgDomains.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    ExportM365-DomainDkim -getMgDomains $object

    .NOTES
        NAME: ExportM365-DomainDkim
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>
    
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $DkimSelectors = @(
        "selector1",
        "selector2"
    )

    $getMgDomains | ForEach { 
        try {
            foreach ($DkimSelector in $DkimSelectors) {
                $dnsRecordName = "$($DkimSelector)._domainkey.$($_.Id)"
                $records = Resolve-DnsName -Type TXT -Name $dnsRecordName -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=DKIM1"; } 
                if ($records) {
                    Write-Host " - $($dnsRecordName) found" -ForegroundColor Green
                    foreach ($record in $records) {
                        [PSCustomObject]@{
                            "Domain"  = $record.Name
                            "Type"    = $record.Type
                            "TTL"     = $record.TTL
                            "Section" = $record.Section
                            "Strings" = $record.Strings -join ';'
                        }
                    }
                } else {
                    Write-Host " - $($dnsRecordName) not found" -ForegroundColor DarkYellow
                    [PSCustomObject]@{
                        "Domain"  = $dnsRecordName
                        "Type"    = "TXT"
                        "TTL"     = ""
                        "Section" = ""
                        "Strings"   = "No DKIM record found"
                    }
                }
            }
        }
        catch {
            Write-Host (ExportM365-Exception -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain"  = $dnsRecordName
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Strings"  = "Error retrieving DKIM record; $($_.Exception.message)"
            }
        }
    } 
}

function ExportM365-DomainDmarc {
    <#
    .SYNOPSIS
    Returns DMARC DNS records for domains obtained from Get-MgDomains

    .DESCRIPTION
    Returns DMARC DNS records for domains obtained from Get-MgDomains. Uses selector1 and selector2 DMARC record types.

    .PARAMETER getMgDomains
    Results of Get_MgDomains output. Iterable object with tenancy domains to identify DNS records for

    .EXAMPLE
    ExportM365-DomainDmarc -getMgDomains $object

    .NOTES
        NAME: ExportM365-DomainDmarc
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Resolve-DnsName

        CHANGELOG:    
    #>

    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][object]$getMgDomains
    )

    $getMgDomains | ForEach-Object { 
        try {
            $dnsRecordName = "_dmarc." + $_.Id
            $records = Resolve-DnsName -Type TXT -Name $dnsRecordName -ErrorAction SilentlyContinue | Where-Object { $_.Strings -match "^v=DMARC1"; } 
            if ($records) {
                Write-Host " - $($dnsRecordName) found" -ForegroundColor Green
                foreach ($record in $records) {
                    [PSCustomObject]@{
                        "Domain"  = $record.Name
                        "Type"    = $record.Type
                        "TTL"     = $record.TTL
                        "Section" = $record.Section
                        "Strings" = $record.Strings -join ';'
                    }
                }
            } else {
                Write-Host " - $($dnsRecordName) not found" -ForegroundColor DarkYellow
                [PSCustomObject]@{
                    "Domain"  = $dnsRecordName
                    "Type"    = "TXT"
                    "TTL"     = ""
                    "Section" = ""
                    "Strings"  = "No DMARC record found"
                }
            }
        }
        catch {
            Write-Host (ExportM365-Exception -message " - $($dnsRecordName) error: $($_.Exception.message)" -exception $_) -ForegroundColor Red
            [PSCustomObject]@{
                "Domain" = $dnsRecordName
                "Type"    = "TXT"
                "TTL"     = ""
                "Section" = ""
                "Value"  = "Error retrieving DMARC record; $($_.Exception.message)"
            }
        }
    } 
}

function ExportM365-EntraUsers {
    <#
    .SYNOPSIS
    Exports Microsoft Entra users

    .DESCRIPTION
    Exports Microsoft Entra users

    .EXAMPLE
    ExportM365-EntraUsers

    .NOTES
        NAME: ExportM365-EntraUsers
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
        $exception = ExportM365-Exception -message "Failed retrieving M365 users" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-EntraDevices {
    <#
    .SYNOPSIS
    Exports Microsoft Entra managed devices

    .DESCRIPTION
    Exports Microsoft Entra managed devices

    .EXAMPLE
    ExportM365-EntraDevices

    .NOTES
        NAME: ExportM365-EntraDevices
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
        $exception = ExportM365-Exception -message "Failed retrieving M365 devices" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }

}

function ExportM365-EntraServicePrincipals {
    <#
    .SYNOPSIS
    Exports Microsoft Entra service principals and their assigned rights

    .DESCRIPTION
    Exports Microsoft Entra service principals and their assigned rights from any application role assignments or oAuth grants.

    .EXAMPLE
    ExportM365-EntraServicePrincipals

    .NOTES
        NAME: ExportM365-EntraServicePrincipals
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
        $exception = ExportM365-Exception -message "Failed retrieving M365 service principals" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-EntraUserMfa {
    <#
    .SYNOPSIS
    Returns Microsoft Entra MFA settings for all users

    .DESCRIPTION
    Creates a Microsoft Entra user registration details report for each user including and SSPR, MFA settings and preferences. Returns details for all users.

    .EXAMPLE
    ExportM365-EntraUserMfa

    .NOTES
        NAME: ExportM365-EntraUserMfa
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
        $exception = ExportM365-Exception -message "Failed retrieving Entra user MFA report" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


function ExportM365-Test-Guid {
    <#
    .SYNOPSIS
    Validates a given input string and checks string is a valid GUID

    .DESCRIPTION
    Validates a given input string and checks string is a valid GUID by using the .NET method Guid.TryParse

    .PARAMETER InputObject
    The GUID(s) to test

    .EXAMPLE
    ExportM365-Test-Guid -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Test-Guid
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

function ExportM365-Resolve-MgObject {
    <#
    .SYNOPSIS
    Resolve a Microsoft Graph item GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Graph item GUID to an object containing Get-MgDirectoryObject properties including user principal name, display name and app display name

    .PARAMETER InputObject
    The GUID(s) of the directory object to resolve

    .EXAMPLE
    ExportM365-Resolve-MgObject -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Resolve-MgObject
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
        if (ExportM365-Test-Guid -InputObject $InputObject) {
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
                        $exception = ExportM365-Exception -message "Unable to find object using Get_MgDirectoryObject using provided GUID" -exception $_
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

function ExportM365-Resolve-MgNamedLocation {
    <#
    .SYNOPSIS
    Resolve a Microsoft Conditional Access Policy named location GUID to an object

    .DESCRIPTION
    Resolve a Microsoft Conditional Access Policy named location GUID to an object containing Get-MgIdentityConditionalAccessNamedLocation properties including display name and CIDR IP address details

    .PARAMETER InputObject
    The GUID(s) of the named location to resolve

    .EXAMPLE
    ExportM365-Resolve-MgNamedLocation -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Resolve-MgNamedLocation
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
        if (ExportM365-Test-Guid -InputObject $InputObject) {
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

function ExportM365-Resolve-MgAdminUnit {
    <#
    .SYNOPSIS
    Resolve an Entra Administrative Unit ID to an object

    .DESCRIPTION
    Resolve an Entra Administrative Unit ID to an object

    .PARAMETER InputObject
    The GUID(s) of the admin unit to resolve

    .EXAMPLE
    ExportM365-Resolve-MgAdminUnit -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Resolve-MgAdminUnit
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
        if (ExportM365-Test-Guid -InputObject $InputObject) {
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

function ExportM365-Resolve-MgServicePrincipal {
    <#
    .SYNOPSIS
    Resolve a service principal appId

    .DESCRIPTION
    Resolve a service principal appId to an object

    .PARAMETER InputObject
    The appId of the service principal to resolve

    .EXAMPLE
    ExportM365-Resolve-MgServicePrincipal -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Resolve-MgServicePrincipal
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
        if (ExportM365-Test-Guid -InputObject $InputObject) {
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

function ExportM365-Resolve-MgDirectoryRoleTemplate {
    <#
    .SYNOPSIS
    Resolve a directory role template 

    .DESCRIPTION
    Resolve a directory role template Id to a display name

    .PARAMETER InputObject
    The Id of the directory role template to resolve

    .EXAMPLE
    ExportM365-Resolve-MgDirectoryRoleTemplate -InputObject "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Resolve-MgDirectoryRoleTemplate
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



function ExportM365-GroupMembers {
    <#
    .SYNOPSIS
    Return the members of a group, recursively 

    .DESCRIPTION
    Return all M365 group members recursively to identify any users with direct, or indirect, membership in the group.  

    .PARAMETER groupId
    The ID of the group to obtain a membership list for

    .EXAMPLE
    ExportM365-GroupMembers -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-GroupMembers
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgGroupTransitiveMember

        CHANGELOG:
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][String]$groupId
    )

    $groupMembers = Get-MgGroupTransitiveMember -GroupId $groupId -All
    $groupMembersObjects = $groupMembers | ForEach-Object { ExportM365-Resolve-MgObject -InputObject $_.Id }
    return $groupMembersObjects
}


function ExportM365-SetPrivilegedRoleMember {
    <#
    .SYNOPSIS
    Define and return a custom object containing the details of a user assigned a built-in role 

    .DESCRIPTION
    Iterates through the provided object containing built-in role members to build an array of users with 
    the assigned rights. Iterates through any nested groups. Also includes any assigned service principals.

    .PARAMETER roleDisplayName
    The display name of the role they are granted rights to

    .PARAMETER memberPrinicpalId
    The principalId of this user/group/service principal

    .PARAMETER memberObject
    The object containing this instance of a user/group/service principal assigned a specific built-in role 

    .PARAMETER membershipType
    Direct, or eligible / permanent assigned through PIM

    .PARAMETER srcGroup
    A hierarchy of group display names showing how the user was granted the rights 

    .EXAMPLE
    ExportM365-SetPrivilegedRoleMember -roleDisplayName "Global Administrator"
                                        -memberPrinicpalId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3" 
                                        -memberObject $object
                                        -membershipType "Direct" 
                                        -srcGroup "Direct > M365 Global Admins"

    .NOTES
        NAME: ExportM365-SetPrivilegedRoleMember
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           ExportM365-Resolve-MgObject
           ExportM365-GroupMembers
           ExportM365-SetPrivilegedRoleMember

        CHANGELOG:
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[PSCustomObject]])]
    Param(
        [Parameter(Mandatory = $true)]$roleDisplayName,
        [Parameter(Mandatory = $true)]$memberPrinicpalId,
        [Parameter(Mandatory = $true)]$memberObject,
        [Parameter(Mandatory = $true)]$membershipType,
        [Parameter(Mandatory = $true)]$srcGroup
    )

    $memberDataType = $memberObject.AdditionalProperties.'@odata.type'
    $memberDisplayName = $memberObject.AdditionalProperties.displayName
    
    switch ($memberDataType) {
        '#microsoft.graph.user' {
            Write-Host "    + User: $($memberObject.AdditionalProperties.userPrincipalName)"
        }
        '#microsoft.graph.serviceprincipal' {
            Write-Host "    + Service: $($memberObject.AdditionalProperties.appDisplayName)"
        }
        '#microsoft.graph.group' {
            Write-Host "    + Group: $($memberObject.AdditionalProperties.displayName)"

            if ($membershipType -eq 'Permanent') {
                $groupMembers = ExportM365-GroupMembers -groupId $memberPrinicpalId
                Write-Host "             Found $($groupMembers.Count) members"
            } else { 
                $assignedPimGroupMember = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule -Filter "groupId eq '$memberPrinicpalId'" -ExpandProperty * -All
                $eligiblePimGroupMember = Get-MgIdentityGovernancePrivilegedAccessGroupEligibilityScheduleInstance -filter "groupId eq '$memberPrinicpalId'" -ExpandProperty *
                Write-Host "             Found $($assignedPimGroupMember.Count) PIM assigned members"
                Write-Host "             Found $($eligiblePimGroupMember.Count) PIM eligible members"
            }
        }
        default {
            Write-Host "Error. Unknown role membership object type. `n  $($memberObject | ConvertTo-Json)" -ForegroundColor Red
        }
    }

    if ($null -ne $memberObject.administrativeUnitId) {
        $adminUnit = ExportM365-Resolve-MgAdminUnit -InputObject $memberObject.administrativeUnitId
    } else {
        $adminUnit = $null
    }

    if ($null -ne $memberObject.directoryScopeId) {
        $dirScopeName = (ExportM365-Resolve-MgObject -InputObject $memberObject.directoryScopeId).AdditionalProperties.displayName
    } else {
        $dirScopeName = ""
    }

    if ($null -ne $memberObject.appScopeId) {
        $appScopeName = (ExportM365-Resolve-MgObject -InputObject $memberObject.appScopeId).AdditionalProperties.displayName
    } else {
        $appScopeName = ""
    }

    [PSCustomObject]@{
        'odata.type'        = $memberDataType
        'principalId'       = $memberPrinicpalId
        'membership'        = $membershipType
        'role'              = $roleDisplayName
        'srcGroup'          = $srcGroup
        'servicePrincpal'   = if ($memberDataType -eq '#microsoft.graph.serviceprincipal') { $memberObject.AdditionalProperties.appDisplayName } else { '' }
        'user'              = if ($memberDataType -eq '#microsoft.graph.user') { $memberObject.AdditionalProperties.userPrincipalName } else { '' }
        'group'             = if ($memberDataType -eq '#microsoft.graph.group') { $memberObject.AdditionalProperties.displayName } else { '' } 
        'scopedAdminUnitId' = if ($null -ne $memberObject.administrativeUnitId) { $memberObject.administrativeUnitId } else { '' } 
        'scopedAdminUnitDisplayName' = if ($null -ne $memberObject.administrativeUnitId) { $adminUnit.displayName } else { '' } 
        'assignedStartDateTime'   = if ($null -ne $memberObject.assignedStartDateTime) { $memberObject.assignedStartDateTime } else { '' }
        'assignedStartDateTimeTz' = if ($null -ne $memberObject.assignedStartDateTime) { $memberObject.assignedStartDateTime.Kind } else { '' }
        'assignedEndDateTime'     = if ($null -ne $memberObject.assignedEndDateTime) { $memberObject.assignedEndDateTime } else { '' }
        'assignedEndDateTimeTz'   = if ($null -ne $memberObject.assignedEndDateTime) { $memberObject.assignedEndDateTime.Kind } else { '' }
        'assignmentType'          = if ($null -ne $memberObject.assignmentType) { $memberObject.assignmentType } else { '' }
        #'eligibleStatus'          = if ($null -ne $memberObject.status) { $memberObject.status } else { '' }
        'eligibleStartDateTime'   = if ($null -ne $memberObject.eligibleStartDateTime) { $memberObject.eligibleStartDateTime } else { '' }
        'eligibleStartDateTimeTz' = if ($null -ne $memberObject.eligibleStartDateTime) { $memberObject.eligibleStartDateTime.Kind } else { '' }
        'eligibleEndDateTime'     = if ($null -ne $memberObject.eligibleEndDateTime) { $memberObject.eligibleEndDateTime } else { '' }
        'eligibleEndDateTimeTz'   = if ($null -ne $memberObject.eligibleEndDateTime) { $memberObject.eligibleEndDateTime.Kind } else { '' }
        #'eligibleMemberType'      = if ($null -ne $memberObject.memberType) { $memberObject.memberType } else { '' }
        'eligibleRecurrence'      = if ($null -ne $memberObject.scheduleInfo.recurrence) { $memberObject.scheduleInfo.recurrence } else { '' }
        'eligibleExpireType'      = if ($null -ne $memberObject.scheduleInfo.expiration.type) { $memberObject.scheduleInfo.expiration.type } else { '' }
        'eligibleExpireDuration'  = if ($null -ne $memberObject.scheduleInfo.expiration.duration) { $memberObject.scheduleInfo.expiration.duration } else { '' }
        'eligibleDirectoryScopeId'= if ($null -ne $memberObject.directoryScopeId) { $memberObject.directoryScopeId } else { '' }
        'eligibleDirectoryScopeName'= $dirScopeName
        'eligibleAppScopeId'      = if ($null -ne $memberObject.appScopeId) { $memberObject.appScopeId } else { '' }
        'eligibleAppScopeName'    = $appScopeName
        'eligibleCreatedUsing'    = if ($null -ne $memberObject.eligibleCreatedUsing) { $memberObject.eligibleCreatedUsing } else { '' }
        'eligibleCreatedDateTime' = if ($null -ne $memberObject.eligibleCreatedDateTime) { $memberObject.eligibleCreatedDateTime } else { '' }
        'eligibleModifiedDateTime'= if ($null -ne $memberObject.eligibleModifiedDateTime) { $memberObject.eligibleModifiedDateTime } else { '' }
        'eligibilityScheduleId'   = if ($null -ne $memberObject.eligibilityScheduleId) { $memberObject.eligibilityScheduleId } else { '' }
    }

    if ($memberDataType -eq '#microsoft.graph.group') {
        $srcGroup = if ($srcGroup -eq "") { $memberDisplayName } else { "$srcGroup > $memberDisplayName" }
        
        foreach ($member in $groupMembers) {
            $memberDetails = ExportM365-Resolve-MgObject -InputObject $member.id
            Write-Verbose "       + Found member: $($memberDetails.AdditionalProperties.displayName)"

            ExportM365-SetPrivilegedRoleMember `
                -memberPrinicpalId $memberDetails.id `
                -memberObject $memberDetails `
                -membershipType $membershipType `
                -srcGroup $srcGroup `
                -roleDisplayName $roleDisplayName
        }

        foreach ($member in $eligiblePimGroupMember) {
            $memberDetails = ExportM365-Resolve-MgObject -InputObject $member.PrincipalId
            Write-Verbose "       + Found PIM eligible member: $($memberDetails.AdditionalProperties.displayName)"

            $memberDetails | Add-Member -MemberType NoteProperty -Name 'eligibilityScheduleId' -Value $member.eligibilityScheduleId -Force
            $memberDetails | Add-Member -MemberType NoteProperty -Name 'eligibleStartDateTime' -Value $member.startDateTime -Force
            $memberDetails | Add-Member -MemberType NoteProperty -Name 'eligibleEndDateTime' -Value $member.endDateTime -Force

            ExportM365-SetPrivilegedRoleMember `
                -memberPrinicpalId $memberDetails.id `
                -memberObject $memberDetails `
                -membershipType "Eligible" `
                -srcGroup $srcGroup `
                -roleDisplayName $roleDisplayName
        }

        foreach ($member in $assignedPimGroupMember) {
            $memberDetails = ExportM365-Resolve-MgObject -InputObject $member.PrincipalId
            Write-Verbose "       + Found PIM assigned member: $($memberDetails.AdditionalProperties.displayName)"

            $memberDetails | Add-Member -MemberType NoteProperty -Name 'assignmentType' -Value $member.assignmentType -Force
            $memberDetails | Add-Member -MemberType NoteProperty -Name 'assignedStartDateTime' -Value $member.ScheduleInfo.StartDateTime -Force
            $memberDetails | Add-Member -MemberType NoteProperty -Name 'assignedEndDateTime' -Value $member.ScheduleInfo.Expiration.EndDateTime -Force

            ExportM365-SetPrivilegedRoleMember `
                -memberPrinicpalId $memberDetails.id `
                -memberObject $memberDetails `
                -membershipType $membershipType `
                -srcGroup $srcGroup `
                -roleDisplayName $roleDisplayName
        }
    }
}

function ExportM365-RoleAdminUnits {
    <#
    .SYNOPSIS
    Export any administrative units used to manage role allocations

    .DESCRIPTION
    Export any administrative units used to manage role allocations

    .EXAMPLE
    ExportM365-RoleAdminUnits

    .NOTES
        NAME: ExportM365-RoleAdminUnits
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
        $exception = ExportM365-Exception -message "Failed retrieving admin units" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}




function ExportM365-Pim-RoleAssignments {
    <#
    .SYNOPSIS
    Return the requested Entra PIM role assignment data

    .DESCRIPTION
    Return the requested Entra PIM role assignment data

    .PARAMETER type
    The type of object to return. Either activated or schedule

    .EXAMPLE
    ExportM365-Pim-RoleAssignments -type "activated"

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
            $exception = ExportM365-Exception -message "Failed retrieving PIM role assignment schedules" -exception $_
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

function ExportM365-Pim-GroupAssignments {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group assignment data

    .DESCRIPTION
    Return the requested Entra PIM group assignment data

    .PARAMETER groupId
    The ID of the group to return assignments for

    .EXAMPLE
    ExportM365-Pim-GroupAssignments -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Pim-GroupAssignments
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
            $exception = ExportM365-Exception -message "Failed retrieving PIM group assignment schedules"  -exception $_
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


function ExportM365-Pim-RoleEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM role eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM role eligibilitiy data

    .EXAMPLE
    ExportM365-Pim-RoleEligibility

    .NOTES
        NAME: ExportM365-Pim-RoleEligibility
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRoleManagementDirectoryRoleEligibilitySchedule

        CHANGELOG:    
    #>    
	[Cmdletbinding()]param()
	
	process {
		try {
			$eligible = (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty principal,roleDefinition -All -ErrorAction Stop | Select-Object id, principalId, directoryScopeId, roleDefinitionId, status, principal, roleDefinition)
			Write-Host " - Lookup: Found $(($eligible | Measure-Object).Count) PIM role eligibility schedules"
			return $eligible
		} catch {
            $exception = ExportM365-Exception -message "Failed retrieving PIM role eligibility schedules" -exception $_
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

function ExportM365-Pim-GroupEligibility {
    <#
    .SYNOPSIS
    Return the requested Entra PIM group eligibilitiy data

    .DESCRIPTION
    Return the requested Entra PIM group eligibilitiy data

    .PARAMETER groupId
    The ID of the group to return eligibilitiy for

    .EXAMPLE
    ExportM365-Pim-GroupEligibility -groupId "3cb87a8f-0a41-4ca8-8910-e56cc00114a3"

    .NOTES
        NAME: ExportM365-Pim-GroupEligibility
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
            $exception = ExportM365-Exception -message "Failed retrieving PIM group eligibility schedules" -exception $_
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



function ExportM365-Pim-Roles {
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
    ExportM365-Pim-Roles

    .NOTES
        NAME: ExportM365-Pim-Roles
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            Get-MgRoleManagementDirectoryRoleAssignment
            ExportM365-Pim-RoleAssignments
            ExportM365-Pim-RoleEligibility
            Get-MgGroupTransitiveMember
            ExportM365-Pim-GroupAssignments
            ExportM365-Pim-GroupEligibility
            ExportM365-Pim-RolesBuild
           
        CHANGELOG:
          2025-12-04: Rewrite of PIM related functions used in Stage 4 (built-in role memberships)
    #>
	[Cmdletbinding()]param ()

    Write-Host "`nEntra membership of built-in and privileged roles"

	# Data retrieval: Get all directory roles
	$pimRoleMembers = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty principal -All
	$rolesDefinition = Get-MgRoleManagementDirectoryRoleAssignment -ExpandProperty roleDefinition -All 
	$assignmentSchedules = ExportM365-Pim-RoleAssignments -type 'schedule'

	# Data retrieval: For each directory role, add the role definition and assignment schedule properties
	foreach ($role in $pimRoleMembers) {
		$newRoleDefinition = ($rolesDefinition | Where-Object { $_.id -eq $role.id }).roleDefinition
		$role.RoleDefinition = $newRoleDefinition
		
		if ($null -ne $assignmentSchedules) {
			$assignmentSchedule = $assignmentSchedules | Where-Object { ($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.principalId -eq $role.PrincipalId) }
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $assignmentSchedule
		} else {
			Add-Member -InputObject $role -MemberType NoteProperty -Name "assignmentSchedule" -Value $null
		}
	}

	# Data retrieval: Add eligible PIM roles to roles array
	$eligibleRoles = ExportM365-Pim-RoleEligibility
	if ($null -ne $eligibleRoles) {
		$pimRoleMembers += $eligibleRoles
	}
    Write-Host " - Found $(($pimRoleMembers | Measure-Object).count) direct members of roles"
    

	# Data retrieval: Identify duplicates in roles array based on active and eligible
	$pimRoleActivations = ExportM365-Pim-RoleAssignments -type 'activated'

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

		$activeGroupMembers = Get-MgGroupTransitiveMember -GroupId $_.principalId -All -Property id, displayName, userPrincipalName
		
		$assignmentSchedules = ExportM365-Pim-GroupAssignments -groupId $_.principalId		
		$eligibleRoleMembers = ExportM365-Pim-GroupEligibility -groupId $_.principalId

        Write-Host " + Resolving members of $($_.Principal.AdditionalProperties.displayName) group ($($_.principalId))"

		# Add active role members from group, include active schedule and eligibilty info 
        if ($null -ne $activeGroupMembers) {
            foreach ($member in $activeGroupMembers) {
                Write-Host "      Adding active member $($member.AdditionalProperties.displayName)"
                
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
                    "AssignmentType" = "Eligible via Role"
                    "AssignmentStartDateTime" = ""
                    "AssignmentEndDateTime" = ""
                    "GroupName" = $_.Principal.AdditionalProperties.displayName
                    "GroupId" = $_.principalId
                    "Status" = $member.Status
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
	
	$privilegedUsers = ExportM365-Pim-RolesBuild -pimRoleMembers $pimRoleMembers -pimRoleActivations $pimRoleActivations
	return $privilegedUsers
}


function ExportM365-Pim-RolesBuild {
    <#
    .SYNOPSIS
    Build an array of privileged role members from data provided by ExportM365-Pim-Roles

    .DESCRIPTION
    Build an array of privileged role members from data provided by ExportM365-Pim-Roles

    .EXAMPLE
    ExportM365-Pim-RolesBuild

    .NOTES
        NAME: ExportM365-Pim-RolesBuild
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
            ExportM365-Pim-Roles (called by, requires data from)
           
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
                $tmpStartDateTime = $roleActivation.startDateTime | Select -Unique | Sort-Object | Select -First 1 
                $tmpAssignmentType = "Eligible (Active)" 
            } else { 
                $tmpStartDateTime = $null 
                $tmpAssignmentType = "Eligible" 
            }

            if ($roleActivation.endDateTime) { 
                $tmpEndDateTime = $roleActivation.endDateTime | Select -Unique | Sort-Object -Descending | Select -First 1 
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
			"InheritedFrom" = ""
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
					"InheritedFrom" = $activeMember.GroupName
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
					"InheritedFrom" = $eligibleMember.GroupName
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


function ExportM365-CapReportLocation {
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
    ExportM365-CapReportLocation -locations $object -type 'IncludeLocation'

    .NOTES
        NAME: ExportM365-CapReportLocation
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           ExportM365-Resolve-MgNamedLocation
           
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
        $thisLocation =  ExportM365-Resolve-MgNamedLocation -InputObject $location

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

function ExportM365-CapFormatMgObjects {
    <#
    .SYNOPSIS
    Outputs a list of objects that can be resolved by ExportM365-Resolve-MgObject (user, group, etc.) based on their GUID from a conditional access policy

    .DESCRIPTION
    Outputs a list of objects that can be resolved by ExportM365-Resolve-MgObject (user, applications, etc.) based on their GUID from a conditional access policy

    .PARAMETER srcObject
    Iterable source object containing GUIDs. Uses ForEach-Object pipeline

    .PARAMETER getAttrib
    Object attribute to return in list (e.g. displayName or userPrinicpalName)

    .EXAMPLE
    ExportM365-CapFormatMgObjects -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"

    .NOTES
        NAME: ExportM365-CapFormatMgObjects
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           ExportM365-Resolve-MgObject
           
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
        if (ExportM365-Test-Guid -InputObject $PSItem) {
            $object = ExportM365-Resolve-MgObject -InputObject $PSItem
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

function ExportM365-CapFormatMgServicePrincipals {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of service principal objects

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of service principal objects

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    ExportM365-CapFormatMgServicePrincipals -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals

    .NOTES
        NAME: ExportM365-CapFormatMgServicePrincipals
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           ExportM365-Resolve-MgServicePrincipal
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((ExportM365-Resolve-MgServicePrincipal -InputObject $PSItem)) | Out-Null
    }

    return $list
}


function ExportM365-CapFormatMgDirectoryRoleTemplates {
    <#
    .SYNOPSIS
    Resolves and returns the display names for an iterable list of directory roles

    .DESCRIPTION
    Resolves and returns the display names for an iterable list of directory roles

    .PARAMETER srcObject
    Iterable source object containing ID. Uses ForEach-Object pipeline

    .EXAMPLE
    ExportM365-CapFormatMgDirectoryRoleTemplates -srcObject $policy.Conditions.Users.IncludeRoles

    .NOTES
        NAME: ExportM365-CapFormatMgDirectoryRoleTemplates
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           ExportM365-Resolve-MgDirectoryRoleTemplate
           
        CHANGELOG:
          2025-11-21: Implemented resolver function for looking up directory objects
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowNull()]$srcObject
    )

    $list = [System.Collections.Generic.List[Object]]::new()
    $srcObject | ForEach-Object {
        $list.Add((ExportM365-Resolve-MgDirectoryRoleTemplate -InputObject $PSItem)) | Out-Null
    }

    return $list
}

function ExportM365-CapReport {
    <#
    .SYNOPSIS
    Build an HTML report from a conditional access policy

    .DESCRIPTION
    Build an HTML report from a conditional access policy

    .PARAMETER policy
    Conditional access policy object obtained from Microsoft Graph function Get-MgIdentityConditionalAccessPolicy

    .EXAMPLE
    ExportM365-CapReport -policy $policy

    .NOTES
        NAME: ExportM365-CapReport
        VERSION: 1.2

        FUNCTIONS & PERMISSIONS:
           ExportM365-CapFormatMgObjects
           ExportM365-CapFormatMgServicePrincipals
           ExportM365-CapReportLocation
           ExportM365-CapFormatMgObjects
           ExportM365-CapFormatMgDirectoryRoleTemplates
           ExportM365-Test-Guid
           ExportM365-Resolve-MgObject
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
$includeApps = ExportM365-CapFormatMgObjects -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"
$excludeApps = ExportM365-CapFormatMgObjects -srcObject $policy.Conditions.Applications.ExcludeApplications -getAttrib "displayName"

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

$includeServicePrincipals = ExportM365-CapFormatMgServicePrincipals -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals
$excludeServicePrincipals = ExportM365-CapFormatMgServicePrincipals -srcObject $policy.Conditions.ClientApplications.ExcludeServicePrincipals

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
    $outHtml += ExportM365-CapReportLocation -type "IncludeLocation" -locations $policy.conditions.Locations.IncludeLocations 
    $outHtml += ExportM365-CapReportLocation -type "ExcludeLocation" -locations $policy.conditions.Locations.ExcludeLocations 
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


$includeUsers = ExportM365-CapFormatMgObjects -srcObject $policy.Conditions.Users.IncludeUsers -getAttrib "userPrincipalName"
$excludeUsers = ExportM365-CapFormatMgObjects -srcObject $policy.Conditions.Users.ExcludeUsers -getAttrib "userPrincipalName"

# Resolve object IDs of included groups
$includeGroups = [System.Collections.ArrayList]::new()
$includeGroupsMembers = @{}
$policy.Conditions.Users.IncludeGroups | ForEach-Object {
    if (ExportM365-Test-Guid -InputObject $PSItem) {
        $group = ExportM365-Resolve-MgObject -InputObject $PSItem
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
    if (ExportM365-Test-Guid -InputObject $PSItem) {
        $group = ExportM365-Resolve-MgObject -InputObject $PSItem
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

$includeRoles = ExportM365-CapFormatMgDirectoryRoleTemplates -srcObject $policy.Conditions.Users.IncludeRoles
$excludeRoles = ExportM365-CapFormatMgDirectoryRoleTemplates -srcObject $policy.Conditions.Users.ExcludeRoles

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

function ExportM365-CapSummary {
    <#
    .SYNOPSIS
    Build object summarising conditional access policy properties

    .DESCRIPTION
    Build object summarising conditional access policy properties. Used when exporting summary of all policies to CSV.

    .PARAMETER policy
    Object containing conditional access policy.

    .EXAMPLE
    ExportM365-CapSummary -policy $policy

    .NOTES
        NAME: ExportM365-CapSummary
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


function ExportM365-Caps {
    <#
    .SYNOPSIS
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .DESCRIPTION
    Obtain and iterate through all conditional access policies and build a HTML report for each

    .PARAMETER OutputFolder
    The folder to create the individual HTML files in. Defined in the $script:exportTargetFolder variable.

    .EXAMPLE
    ExportM365-Caps -OutputFolder $OutputFolder

    .NOTES
        NAME: ExportM365-Caps
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgIdentityConditionalAccessPolicy
           ExportM365-CapReport
           
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

        $capSummary += ExportM365-CapSummary -policy $policy

        try {
            $capExportFileName = ($policy.DisplayName) -replace '[^a-zA-Z0-9-_ ]', ' '
            $outHtml = ExportM365-CapReport -policy $policy 
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

function ExportM365-AuthMethodPolicy {
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
        $exception = ExportM365-Exception -message "Failed retrieving Entra authentication method policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-AuthStrengthPolicy {
    <#
    .SYNOPSIS
    Exports Microsoft Entra authentication strength policies

    .DESCRIPTION
    Exports Microsoft Entra authentication strength policies

    .EXAMPLE
    ExportM365-AuthStrengthPolicy

    .NOTES
        NAME: ExportM365-AuthStrengthPolicy
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
        $exception = ExportM365-Exception -message "Failed retrieving Entra authentication strength policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

$script:isUnifiedAuditLogEnabled = $null

function ExportM365-AuditLogSettings {
    <#
    .SYNOPSIS
    Exports Microsoft Security and Compliance audit log settings

    .DESCRIPTION
    Exports Microsoft Security and Compliance audit log settings

    .EXAMPLE
    ExportM365-AuditLogConfig

    .NOTES
        NAME: ExportM365-AuditLogConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-AdminAuditLogConfig

        CHANGELOG:    
    #>

    try {
        Write-Host "`nMicrosoft Security and Compliance audit log settings"
        
        # Get audit log settings
        $getAdminAuditLogConfig = Get-AdminAuditLogConfig 

        # Export to CSV
        $getAdminAuditLogConfig | Export-CSV $script:exportTarget.auditLogConfig -NoTypeInformation
        Write-Host " - Exported $($getAdminAuditLogConfig.count) M365 audit log settings to $($script:exportTarget.auditLogConfig)" -ForegroundColor Green

        if ($getAdminAuditLogConfig.count -gt 0) {
            Write-Host " - Failed to retrieve M365 audit log settings" -ForegroundColor Red
        }

        if ($getAdminAuditLogConfig.UnifiedAuditLogIngestionEnabled -eq $true) {
            Write-Host " - Unified audit logging enabled" -ForegroundColor Green
            $script:isUnifiedAuditLogEnabled = $true
        } else {
            Write-Host " - Unified audit logging disabled (set: $($getAdminAuditLogConfig.UnifiedAuditLogIngestionEnabled))" -ForegroundColor DarkYellow
            $script:isUnifiedAuditLogEnabled = $false
        }

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving Microsoft Security and Compliance audit log settings" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-AuditLogRetention {
    <#
    .SYNOPSIS
    Export unified audit log retention policies.

    .DESCRIPTION
    Export unified audit log retention policies. Licensing requirement.

    .EXAMPLE
    ExportM365-AuditLogRetention 

    .NOTES
        NAME: ExportM365-AuditLogRetention
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-UnifiedAuditLogRetentionPolicy
           
        CHANGELOG:
    #>

    try {
        Write-Host "`nMicrosoft Security and Compliance audit log retention settings"

        # Get audit log retention policy
        $getUnifiedAuditLogRetentionPolicy = Get-UnifiedAuditLogRetentionPolicy
        if ($getUnifiedAuditLogRetentionPolicy.count -gt 0) {
            Write-Host " - Found $($getUnifiedAuditLogRetentionPolicy.count) M365 unified audit log retention policies"

            # Export to CSV
            $getUnifiedAuditLogRetentionPolicy | Export-CSV $script:exportTarget.auditLogRetention -NoTypeInformation
            Write-Host " - M365 audit log retention policy to $($script:exportTarget.auditLogRetention)" -ForegroundColor Green
        } else {
            Write-Host " - Found 0 M365 unified audit log retention policies" -ForegroundColor DarkYellow
        }
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving Microsoft Security and Compliance audit log retention policies.  Unified audit logging may not be a licensed and configured feature for this tenant." -exception $_
        Write-Host $exception -ForegroundColor DarkYellow
        return $false
    }
}


$script:unifiedAuditLogRecordTypes = @(
    # Privileged Actions
    'ExchangeAdmin',
    'AppAdminActivity',
    'AppSettingsAdminActivity',

    # IAM
    'SecurityComplianceRBAC',
    'SecurityComplianceUserChange',
    'URBACAssignment',
    'URBACRole',
    'URBACEnableState',
    'AzureActiveDirectory',
    'AzureActiveDirectoryAccountLogon',
    'AzureActiveDirectoryStsLogon',

    # Changes to auditing & policy
    'AuditRetentionPolicy',
    'AuditConfig',
    'DataGovernance',
    'ComplianceSettingsChange',
    'CrossTenantAccessPolicy',
    'Microsoft365BackupBackupPolicy',

    # DLP
    'ComplianceDLPSharePoint',
    'ComplianceDLPExchange',
    'ComplianceDLPSharePointClassification',
    'ComplianceDLPExchangeClassification',
    'SensitivityLabelAction',
    'SensitivityLabeledFileAction',

    # AI
    'EnableCopilotPlugin',
    'DisableCopilotPlugin',
    'UpdateCopilotPlugin',
    'CreateCopilotPlugin',
    'DeleteCopilotPlugin',
    'UpdateCopilotSettings',
    'ConnectedAIAppInteraction',

    # Threat detection
    'SecurityComplianceCenterEOPCmdlet',
    'ThreatIntelligence',
    'ThreatFinder',
    'SecurityComplianceAlerts',
    'ThreatIntelligenceAtpContent',
    'AadRiskDetection',
    'PurviewInsiderRiskAlerts',

    # Accessing & sharing resources
    'SharePointFileOperation',
    'SharePointSharingOperation',
    'OneDrive'
)


function ExportM365-AuditLogRisky {
    <#
    .SYNOPSIS
    Export risky users and service principals records

    .DESCRIPTION
    Export risky users and service principals records

    .EXAMPLE
    ExportM365-AuditLogRisky 

    .NOTES
        NAME: ExportM365-AuditLogRisky
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgRiskyUser
           Get-MgRiskyServicePrincipal
           
        CHANGELOG:
    #>
    [CmdletBinding()]
    Param(
    )

    try {
        Write-Host "`nRisky users and service principals log records"

        $getRiskySp += Get-MgRiskyServicePrincipal
        if ($getRiskySp.count -gt 0) {
            Write-Host " - Found $($getRiskySp.count) M365 risky service principal logs"

            # Export to CSV
            $getRiskySp | Export-CSV $script:exportTarget.riskyServicePrincipals -NoTypeInformation
            Write-Host " - M365 risky service principal logs to $($script:exportTarget.riskyServicePrincipals)" -ForegroundColor Green


            $getRiskySpHistory = @() 
            $getRiskySp | ForEach-Object {
                $getRiskySpHistory += Get-MgRiskyServicePrincipalHistory -RiskyServicePrincipalId $_.Id 
                Write-Host " - Getting history for $($_.displayName)"
            }

            if ($getRiskySpHistory.count -gt 0) {
                Write-Host " - Found $($getRiskySpHistory.count) M365 risky service principal history logs"

                # Export to CSV
                $getRiskySpHistory | Export-CSV $script:exportTarget.riskySpHistory -NoTypeInformation
                Write-Host " - M365 risky service principal history logs to $($script:exportTarget.riskySpHistory)" -ForegroundColor Green
            } else {
                Write-Host " - Found 0 M365 risky service principal history logs" -ForegroundColor DarkYellow
            }
        } else {
            Write-Host " - Found 0 M365 risky service principal logs" -ForegroundColor DarkYellow
        }
 
        
        $getRiskyUser += Get-MgRiskyUser -All
        if ($getRiskyUser.count -gt 0) {
            Write-Host " - Found $($getRiskyUser.count) M365 risky user logs"

            # Export to CSV
            $getRiskyUser | Export-CSV $script:exportTarget.riskyUsers -NoTypeInformation
            Write-Host " - M365 risky user logs to $($script:exportTarget.riskyUsers)" -ForegroundColor Green
        } else {
            Write-Host " - Found 0 M365 risky user logs" -ForegroundColor DarkYellow
        }

    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving risky user and service principal records" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-AuditLogUnified {
    <#
    .SYNOPSIS
    Export a sample of the unified audit log records

    .DESCRIPTION
    Export a sample of the unified audit log records for each of the RecordTypes defined in $script:unifiedAuditLogRecordTypes

    .PARAMETER limit 
    Number of records to retrieve. Defaults to 100 per RecordType (defined in $script:unifiedAuditLogRecordTypes)

    .EXAMPLE
    ExportM365-AuditLogUnified 

    .NOTES
        NAME: ExportM365-AuditLogUnified
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Search-UnifiedAuditLog
           
        CHANGELOG:
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][Int]$limit = 100
    )

    try {
        Write-Host "`nMicrosoft Exchange Online unified audit log records"

        $allRecords = @() 

        # Get 100 records as a sample
        $script:unifiedAuditLogRecordTypes | ForEach-Object {
            $records = Search-UnifiedAuditLog -StartDate ((Get-Date).AddDays(-31)) -EndDate (Get-Date) -RecordType $_ -ResultSize $limit

            Write-Host " - $($_): Found $($records.count)"
            if ($records.count -gt 0) {
                $allRecords += $records
            }
        } 

        # Export to CSV
        if ($allRecords.count -gt 0) {
            $allRecords | ConvertTo-Json -depth 50 | Out-File $script:exportTarget.unifiedAuditLog
            Write-Host " - Exported M365 unified audit log to $($script:exportTarget.unifiedAuditLog)" -ForegroundColor Green
        } else {
            Write-Host " - Found 0 M365 unified audit log entries" -ForegroundColor DarkYellow
        }

    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving unified audit log records" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-AuditLogProvisioning {
    <#
    .SYNOPSIS
    Export a sample of audit log provisioning records

    .DESCRIPTION
    Export a sample of audit log provisioning records

    .PARAMETER limit 
    Number of records to retrieve. Defaults to 1000

    .EXAMPLE
    ExportM365-AuditLogProvisioning 

    .NOTES
        NAME: ExportM365-AuditLogProvisioning
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgAuditLogProvisioning
           
        CHANGELOG:
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][Int]$limit = 1000
    )

    try {
        Write-Host "`nMicrosoft Security and Compliance audit log provisioning records"

        # Get provisioning audit log sample
        $getMgAuditLogProvisioning = Get-MgAuditLogProvisioning -Sort "activityDateTime DESC" -Top $limit
        if ($getMgAuditLogProvisioning.count -gt 0) {
            Write-Host " - Found $($getMgAuditLogProvisioning.count) M365 provisioning audit log entries"

            # Export to CSV
            $getMgAuditLogProvisioning | Export-CSV $script:exportTarget.auditLogProvisioning -NoTypeInformation
            Write-Host " - Exported M365 provisioning audit log to $($script:exportTarget.auditLogProvisioning)" -ForegroundColor Green
        } else {
            Write-Host " - Found 0 M365 provisioning audit log entries" -ForegroundColor DarkYellow
        }
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving audit log provisioning records" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-AuditLogDirectory {
    <#
    .SYNOPSIS
    Export a sample of audit log directory records

    .DESCRIPTION
    Export a sample of audit log directory records

    .PARAMETER limit 
    Number of records to retrieve. Defaults to 1000

    .EXAMPLE
    ExportM365-AuditLogDirectory 

    .NOTES
        NAME: ExportM365-AuditLogDirectory
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgAuditLogDirectoryAudit
           
        CHANGELOG:
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][Int]$limit = 1000
    )

    try {
        Write-Host "`nMicrosoft Security and Compliance audit log directory records"

        # Get directory audit log sample
        $getMgAuditLogDirectoryAudit = Get-MgAuditLogDirectoryAudit -Sort "activityDateTime DESC" -Top 1000
        if ($getMgAuditLogDirectoryAudit.count -gt 0) {
            Write-Host " - Found $($getMgAuditLogDirectoryAudit.count) M365 directory audit log entries"

            $getMgAuditLogDirectoryAudit | Select-Object -Property `
                    Id, `
                    @{l='ActivityDateTime'; e={ $_.ActivityDateTime.Date }}, `
                    ActivityDisplayName, Category, LoggedByService, OperationType, Result, ResultReason, `
                    @{l='TargetResources'; e={ $_.TargetResources | ConvertTo-Json -Depth 4 }}, `
                    CorrelationId, `
                    @{l='InitiatedByAppDisplayName'; e={$_.InitiatedBy.App.displayName}}, `
                    @{l='InitiatedByAppServicePrincipalName'; e={$_.InitiatedBy.App.servicePrincipalName}}, `
                    @{l='InitiatedByUserDisplayName'; e={$_.InitiatedBy.User.displayName}}, `
                    @{l='InitiatedByUserUpn'; e={$_.InitiatedBy.User.userPrincipalName}}, `
                    @{l='InitiatedByUserIpAddress'; e={$_.InitiatedBy.User.IPAddress}}, `
                    @{l='InitiatedByUserAdditional'; e={$_.InitiatedBy.User.AdditionalProperties | ConvertTo-Json -Depth 4}}, `
                    @{l='AdditionalProperties'; e={ $_.AdditionalProperties | ConvertTo-Json -Depth 4 }} `
               | Export-CSV $script:exportTarget.auditLogDirectoryAudit -NoTypeInformation
            Write-Host " - Exported M365 directory audit log to $($script:exportTarget.auditLogDirectoryAudit)" -ForegroundColor Green
        } else {
            Write-Host " - Found 0 M365 directory audit log entries" -ForegroundColor DarkYellow
        }
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving audit log directory records" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }

}

function ExportM365-AuditLogSignIn {
    <#
    .SYNOPSIS
    Obtain and iterate through sign-in logs building a summary for export

    .DESCRIPTION
    Obtain and iterate through sign-in logs building a summary for export

    .PARAMETER LogFilter
    Specifies the filter to be used when obtaining the logs from Get-MgAuditLogSignIn. Screened 
    through switch statement. Valid set includes Top10k, oldest, all.

    .EXAMPLE
    ExportM365-AuditLogSignIn OutputFolder $OutputFolder

    .NOTES
        NAME: ExportM365-AuditLogSignIn
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MgAuditLogSignIn
           
        CHANGELOG:
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][ValidateSet('Top10k','oldest','All','None')][String]$LogFilter,
        [Parameter(Mandatory = $true)][String]$logFile
    )

    try {
        Write-Host "`nAudit Sign-in Logs"

        [array]$signInLogs = @()
        switch ($LogFilter) {
            "Top10k" {
                $logs = Get-MgAuditLogSignIn -Filter "status/errorCode eq 0" -Sort "createdDateTime DESC" -Top 10000
                $desc = "retrieving top 10k logs"
            }
            "oldest" {
                $logs = Get-MgAuditLogSignIn -Sort "createdDateTime ASC" -Top 100
                $desc = "retrieving sample of oldest logs"
            }

            # @todo BETA api only. Review
            "nonInteractiveUser" {
                $logs = Get-MgAuditLogSignIn -Filter "signInEventTypes/any(t: t eq 'nonInteractiveUser')" -Top 2000
                $desc = "retrieving sample of non-interactive user logs"
            }

            # @todo BETA api only. Review
            "interactiveUser" {
                $logs = Get-MgAuditLogSignIn -Filter "signInEventTypes/any(t: t eq 'interactiveUser')" -Top 2000
                $desc = "retrieving sample of interactive user logs"
            }

            # @todo BETA api only. Review
            "servicePrincipal" {
                $logs = Get-MgAuditLogSignIn -Filter "signInEventTypes/any(t: t eq 'servicePrincipal')" -Top 2000
                $desc = "retrieving sample of service principal logs"
            }

            # @todo BETA api only. Review
            "managedIdentity" {
                $logs = Get-MgAuditLogSignIn -Filter "signInEventTypes/any(t: t eq 'managedIdentity')" -Top 2000
                $desc = "retrieving sample of managed identity logs"
            }

            "all" {
                $logs = Get-MgAuditLogSignIn -Filter "status/errorCode eq 0" -All $true
                $desc = "retrieving all logs"
            }

            default {
                Write-Host "Skipping log export (LogFilter: $($LogFilter))"
                return $null
            }
        }

        Write-Host " - Found $($logs.Count) sign in logs - $desc "
        Write-Progress -PercentComplete -1 -Activity "Fetching audit log sign-in history from Graph API"

        foreach ($log in $logs) {
            # Display some progress (based on log count)
            $currentIndex = $logs.indexOf($log) + 1

            $progress = @{
                Activity         = "Exporting Sign-in Logs..."
                PercentComplete  = [Decimal]::Divide($currentIndex, $logs.Count) * 100
                CurrentOperation = "Processing `"$(($log.CreatedDateTime).ToString("d MMMM yyyy"))`""
            }
            if ($currentIndex -eq $logs.Count) { $progress.Add("Completed", $true) }

            Write-Progress @progress

            if ($script:LegacyAuthClients -contains $log.ClientAppUsed) {
                Write-Host " - Legacy auth client $($log.ClientAppUsed) ($($log.CreatedDateTime))" -ForegroundColor DarkYellow
                Write-Host "     + $($log.AppDisplayName)  ($($log.Id))" -ForegroundColor DarkYellow
            }

            $isMfaEnforced = $false
            $log.AppliedConditionalAccessPolicies | ForEach-Object { 
                if ($_.EnforcedGrantControls -contains 'mfa' -and $_.Result -eq 'success') {
                    $isMfaEnforced = $true
                }
            }

            $line = [pscustomobject][Ordered]@{ 
                id = $log.id
                createdDateTime = $log.createdDateTime
                appId = $log.appId
                appDisplayName = $log.appDisplayName
                resourceId = $log.resourceId
                resourceDisplayName = $log.resourceDisplayName
                clientAppUsed = $log.clientAppUsed
                clientCredentialType = $log.clientCredentialType
                SignInEventTypes = $log.SignInEventTypes -join ", "
                statusErrorCode = $log.status.errorCode
                statusErrorDetail = $log.status.additionalDetails
                statusErrorFailure = $log.status.failureReason
                mfaAuthMethod = $log.MfaDetail.AuthMethod
                userDisplayName = $log.userDisplayName
                userPrincipalName = $log.userPrincipalName
                userId = $log.userId
                UserType = $log.UserType
                riskState = $log.riskState
                riskEventTypes = $log.RiskEventTypesV2 -join ", "
                riskLevelDuringSignIn = $log.riskLevelDuringSignIn
                riskLevelAggregated = $log.riskLevelAggregated
                riskDetail = $log.riskDetail
                locationCity = $log.location.city
                locationRegion = "$($log.location.state), $($log.Location.CountryOrRegion)"
                deviceId = $log.deviceDetail.deviceId
                deviceDisplayName = $log.deviceDetail.displayName
                deviceOperatingSystem = $log.deviceDetail.operatingSystem
                devicebrowser = $log.deviceDetail.browser
                deviceIsCompliant = $log.deviceDetail.isCompliant
                deviceIsManaged = $log.deviceDetail.isManaged
                deviceTrustType = $log.deviceDetail.trustType
                ipAddress = $log.ipAddress
                isInteractive = $log.isInterActive
                conditionalAccessStatus = $log.conditionalAccessStatus
                capHasEnforcedMfa = $isMfaEnforced
                appliedConditionalAccessPolicies = $log.AppliedConditionalAccessPolicies | ConvertTo-Json 
                correlationId = $log.correlationId
            }
            
            $signInLogs += $line
        }

        $signInLogs | Export-CSV $logFile -NoTypeInformation
        Write-Host " - Exported M365 sign-in audit log (oldest) to $($script:exportTarget.signInLogsOldest)" -ForegroundColor Green
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving sign-in audit log sample" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoMalwarePolicy {
    <#
    .SYNOPSIS
    Export anti-malware policy and rule threat protection policies

    .DESCRIPTION
    Obtain anti-malware policy and rule settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoMalwarePolicy

    .NOTES
        NAME: ExportM365-DfoMalwarePolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MalwareFilterPolicy
           Get-MalwareFilterRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-malware protection settings"

        # Get EOP malware policy settings
        $getMalwareFilterPolicy = Get-MalwareFilterPolicy 
        $getMalwareFilterPolicy | Export-CSV $script:exportTarget.eopMalwarePolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection policy settings to $($script:exportTarget.eopMalwarePolicy)" -ForegroundColor Green

        $getMalwareFilterRule = Get-MalwareFilterRule 
        $getMalwareFilterRule | Export-CSV $script:exportTarget.eopMalwareRule -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection rules to $($script:exportTarget.eopMalwareRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving EOP anti-malware policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoPhishingPolicy {
    <#
    .SYNOPSIS
    Export anti-malware policy and rule threat protection policies

    .DESCRIPTION
    Export anti-malware policy and rule settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoPhishingPolicy

    .NOTES
        NAME: ExportM365-DfoPhishingPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-AntiPhishPolicy
           Get-AntiPhishRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-phishing protection settings"

        # Get EOP phishing policy settings
        $getAntiPhishPolicy = Get-AntiPhishPolicy 
        $getAntiPhishPolicy | Export-CSV $script:exportTarget.eopPhishPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection policy settings to $($script:exportTarget.eopPhishPolicy)" -ForegroundColor Green

        # Export to CSV
        $getAntiPhishRule = Get-AntiPhishRule 
        $getAntiPhishRule | Export-CSV $script:exportTarget.eopPhishRule -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection rule to $($script:exportTarget.eopPhishRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving EOP anti-phishing policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoSpamPolicy {
    <#
    .SYNOPSIS
    Export anti-spam policy threat protection policies

    .DESCRIPTION
    Export anti-spam policy settings from Exchange Online / Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoSpamPolicy

    .NOTES
        NAME: ExportM365-DfoSpamPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-HostedConnectionFilterPolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-spam protection settings"

        # Get EOP spam policy settings
        $getHostedConnectionFilterPolicy = Get-HostedConnectionFilterPolicy 
        $getHostedConnectionFilterPolicy | Export-CSV $script:exportTarget.eopSpamPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-spam protection policy settings to $($script:exportTarget.eopSpamPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving EOP anti-spam policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoSafeLinksPolicy {
    <#
    .SYNOPSIS
    Export safe links policy threat protection policies

    .DESCRIPTION
    Export safe links policy settings from Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoSafeLinksPolicy

    .NOTES
        NAME: ExportM365-DfoSafeLinksPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeLinksPolicy
           Get-SafeLinksRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nDefender Threat Protection safe links policy settings"

        # Get DTP safe links settings
        $getSafeLinksPolicy = Get-SafeLinksPolicy 
        $getSafeLinksPolicy | Export-CSV $script:exportTarget.dtppSafeLinksPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe links policy settings to $($script:exportTarget.dtppSafeLinksPolicy)" -ForegroundColor Green

        $getSafeLinksRule = Get-SafeLinksRule 
        $getSafeLinksRule | Export-CSV $script:exportTarget.dtppSafeLinksRule -NoTypeInformation
        Write-Host " - Exported DTP safe links rule settings to $($script:exportTarget.dtppSafeLinksRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving DTP safe links policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoSafeAttachmentsPolicy {
    <#
    .SYNOPSIS
    Export safe attachment threat protection policies

    .DESCRIPTION
    Export safe attachments policy settings from Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoSafeAttachmentsPolicy

    .NOTES
        NAME: ExportM365-DfoSafeAttachmentsPolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeAttachmentPolicy
           Get-SafeAttachmentRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nDefender Threat Protection safe attachments policy settings"

        # Get DTP safe attachments settings
        $getSafeAttachmentPolicy = Get-SafeAttachmentPolicy
        $getSafeAttachmentPolicy | Export-CSV $script:exportTarget.dtppSafeAttachmentPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe attachments policy settings to $($script:exportTarget.dtppSafeAttachmentPolicy)" -ForegroundColor Green

        # Export as CSV
        $getSafeAttachmentRule = Get-SafeAttachmentRule
        $getSafeAttachmentRule | Export-CSV $script:exportTarget.dtppSafeAttachmentRule -NoTypeInformation
        Write-Host " - Exported DTP safe attachments rule to $($script:exportTarget.dtppSafeAttachmentRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving DTP safe attachments policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function ExportM365-DfoQuarantinePolicy {
    <#
    .SYNOPSIS
    Export threat protection quarantine policies

    .DESCRIPTION
    Export quarantine policy settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    ExportM365-DfoQuarantinePolicy

    .NOTES
        NAME: ExportM365-DfoQuarantinePolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-QuarantinePolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online quarantine policy settings"

        $getQuarantinePolicy = Get-QuarantinePolicy | Select-Object -Property Id, IsValid, Name, EndUserQuarantinePermissions, ESNEnabled, QuarantinePolicyType, QuarantineRetentionDays, IncludeMessagesFromBlockedSenderAddress, DistinguishedName, ObjectCategory, ObjectClass, WhenChanged, WhenCreated, WhenChangedUTC, WhenCreatedUTC, ExchangeObjectId, OrganizationalUnitRoot, OrganizationId, Guid, OriginatingServer, ObjectState, ExchangeVersion
        $getQuarantinePolicy | Export-CSV $script:exportTarget.dfoQuarantinePolicy -NoTypeInformation
        Write-Host " - Exported EOP quarantine policy settings to $($script:exportTarget.dfoQuarantinePolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = ExportM365-Exception -message "Failed retrieving EOP quarantine policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}


    # -----------------------------------
    # Define variables required by script
    # -----------------------------------

    clear

    $script:Output = ($Output | Resolve-Path).Path
    $script:execDateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")

    # Create folders to store output
	$script:exportTargetFolder = [PSCustomObject]@{
		"organisation"      = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\Organisation"
		"iam"               = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\IAM"
		"conditionalAccess" = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\ConditionalAccess"
		"auditLog"          = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\AuditLog"
		"threatProtection"  = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\ThreatProtection"
	}

    # Define where exports will be saved
    $script:exportTarget = [pscustomobject]@{
        "organisations"             = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "organisations.csv"
        "onPremisesSync"            = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "onPremisesSync.csv"
        "domains"                   = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "domains.csv"
        "domainDkim"                = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "domainDkim.csv"
        "domainSpf"                 = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "domainSpf.csv"
        "domainDmarc"               = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "domainDmarc.csv"
        "domainFederationConfig"    = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "domainFederationConfig.csv"
        "authMethodPolicy"          = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "authMethodPolicies.csv"
        "authStrengthPolicy"        = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "authStrengthPolicies.csv"
        "directorySettings"         = Join-Path -Path $script:exportTargetFolder.organisation -ChildPath "directorySettings.csv"
        "signInLogs"                = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogs.csv"
        "signInLogsOldest"          = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogsOldest.csv"
        "signInLogsNonInteractive"  = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogsNonInteractive.csv"
        "signInLogsInteractive"     = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogsInteractive.csv"
        "signInLogsServicePrincipal" = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogsServicePrincipal.csv"
        "signInLogsManagedIdentity" = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "signInLogsManagedIdentity.csv"        
        "auditLogConfig"            = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "auditLogConfig.csv"
        "auditLogDirectoryAudit"    = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "auditLogDirectoryAudit.csv"
        "auditLogRetention"         = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "auditLogRetention.csv"
        "auditLogProvisioning"      = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "auditLogProvisioning.csv"
        "riskyServicePrincipals"    = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "riskyServicePrincipal.csv"
        "riskySpHistory"            = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "riskyServicePrincipalHistory.csv"
        "riskyUsers"                = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "riskyUser.csv"
        "unifiedAuditLog"           = Join-Path -Path $script:exportTargetFolder.auditLog -ChildPath "unifiedAuditLog.json"
        "builtInRoleMembers"        = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "builtInRoleMembers.csv"
        "roleAdminUnits"            = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "adminUnits.csv"
        "userMfa"                   = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "userMFA.csv"
        "users"                     = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "users.csv"
        "servicePrincipals"         = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "servicePrincipals.csv"
        "entraDevices"              = Join-Path -Path $script:exportTargetFolder.iam -ChildPath "entraDevices.csv"
        "dtppSafeLinksPolicy"       = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeLinksPolicy.csv"
        "dtppSafeLinksRule"         = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeLinksRule.csv"
        "dtppSafeAttachmentPolicy"  = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeAttachmentPolicy.csv"
        "dtppSafeAttachmentRule"    = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeAttachmentRule.csv"
        "eopMalwarePolicy"          = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopMalwarePolicy.csv"
        "eopMalwareRule"            = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopMalwareRule.csv"
        "eopPhishPolicy"            = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopPhishPolicy.csv"
        "eopPhishRule"              = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopPhishRule.csv"
        "eopSpamPolicy"             = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopSpamPolicy.csv"
        "dfoQuarantinePolicy"       = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dfoQuarantinePolicy.csv"
        "capSummary"                = Join-Path -Path $script:exportTargetFolder.conditionalAccess -ChildPath "capSummary.csv"
    }

    $script:psGraphModulesRequired = @(
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.DirectoryObjects',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Identity.Governance',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Reports'
    )

    $script:psExoModulesRequired = @(
        'ExchangeOnlineManagement'
    )


    $script:LegacyAuthClients = @(
        'Authenticated SMTP',
        'Autodiscover',
        'Exchange ActiveSync',
        'Exchange Online Powershell',
        'Exchange Web Services',
        'IMAP',
        'MAPI Over HTTP',
        'Offline Address Book',
        'Outlook Anywhere (RPC over HTTP)',
        'Outlook Service',
        'POP3',
        'Reporting Web Services',
        'Universal Outlook',
        'Other Clients'
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
 


    # --------------------------------------------------------------------------------------------------------------
    # Initialise script. Start transcript, check for required modules and safe mode. Output parameters for audit log
    # --------------------------------------------------------------------------------------------------------------

    Start-Transcript -Path (Join-Path -Path $Output -ChildPath "$($script:execDateTime)_ExportM365_transcript.txt")
    Write-Host "Starting $(Get-Date)"

    Write-Host "`n----------------------------------------------"
    Write-Host " This script will output audit evidence to a series of subfolders under the location '$Output'"
    if ((Get-ChildItem $Output  | Measure-Object ).count -gt 4) {
        Write-Host "   - The script should be run from an empty folder to avoid confusion when providing the exported evidence" -ForegroundColor Yellow
    }

    ($script:psGraphModulesRequired + $script:psExoModulesRequired) | foreach {
        $getModuleResult = Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue 
        if ($getModuleResult) {
            $moduleVersionNo = ($getModuleResult | Sort-Object Version -Descending).Version | Join-String -Separator ", "
            Write-Host "   - Found required PowerShell module '$_' [$moduleVersionNo]. If exceptions occur check that module is up-to-date." -ForegroundColor Green
        } else {
            Write-Host "   - Missing required PowerShell module '$_'. Abort now." -ForegroundColor Red
        }
    }

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
    Pause

    if ((-not [string]::IsNullOrEmpty($RunStages))) {
        $RunAll = $false
    }

    Write-Host "`nScript Parameters:" -ForegroundColor Cyan
    Write-Host "  Output Folder: $Output"
    Write-Host "  Run all stages: $RunAll"
    Write-Host "  Run specific stages: $RunStages"
    Write-Host "  Log Filter: $LogFilter"
    Write-Host "`n`n"
}

process {
    #
    # Run Stage 1: Export tenancy settings and policies
    #    @todo Get-MgOrganization: review attributes available in differing tenant configurations
    #    @todo Get-MgDomain: review attributes available in differing tenant configurations
    #    @todo Get-MgBetaDirectorySetting: Requires beta API. Monitor for production API version
    #    @todo Get-MgDomainFederationConfiguration: Test

    if ($RunAll -or $RunStages.Contains(1)) {
        Write-Host "`nSTAGE 1: Getting M365 tenant configuration ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null

        ExportM365-Organization | Out-Null

        ExportM365-OrgOnPremiseSync | Out-Null

        $getMgDomains = ExportM365-Domains
        ExportM365-EmailAuth -domains $getMgDomains | Out-Null
        

        # Get domain password protection settings
        # @todo Test use of beta API

        # $getMgBetaDirectorySetting = Get-MgBetaDirectorySetting
        # $getMgBetaDirectorySettingCount = $getMgBetaDirectorySetting.count
        # Write-Host "Found $getMgBetaDirectorySettingCount M365 directory (password protection) settings"

        # $getMgBetaDirectorySetting | Select-Object -Property * | Export-CSV $script:exportTarget.directorySettings -NoTypeInformation
        # Write-Host "Exported M365 directory settings to CSV $($script:exportTarget.directorySettings)" -ForegroundColor Green


        # Get details of any federated domain authentication
        # @todo Test

        # $getMgDomainFederationConfig = $getMgDomain | ForEach { $_; Get-MgDomainFederationConfiguration -DomainId $_.Id; }
        # $getMgDomainFederationConfigCount = $getMgDomainFederationConfig.count
        # Write-Host "Found M365 domain federation config for $getMgDomainFederationConfigCount domains"

        # $getMgDomainFederationConfig | Select-Object -Property * | Export-CSV $script:exportTarget.domainFederationConfig -NoTypeInformation
        # Write-Host "Exported M365 domain federation config to CSV $($script:exportTarget.domainFederationConfig)" -ForegroundColor Green
    }


    #
    # Stage 2: Export Microsoft Entra user, device and serivce principal information
    #
    if ($RunAll -or $RunStages.Contains(2)) {
        Write-Host "`nSTAGE 2: Getting M365 users and service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null

        ExportM365-EntraUsers | Out-Null

        ExportM365-EntraDevices | Out-Null

        ExportM365-EntraServicePrincipals | Out-Null
    }


    #
    # Stage 3: Export enabled Microsoft Entra MFA settings
    #
    if ($RunAll -or $RunStages.Contains(3)) {
        Write-Host "`nSTAGE 3: Retrieving M365 user MFA settings ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null

        ExportM365-EntraUserMfa | Out-Null
    }


    #
    # Stage 4: Export Microsoft 365 built-in role memberships
    #
    if ($RunAll -or $RunStages.Contains(4)) {
        Write-Host "`nSTAGE 4: Retrieving M365 user built-in role memberships ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null

        ExportM365-RoleAdminUnits | Out-Null

        ExportM365-Pim-Roles | Export-CSV $script:exportTarget.builtInRoleMembers -NoTypeInformation
        Write-Host "Exported built-in role assignments to $($script:exportTarget.builtInRoleMembers)" -ForegroundColor Green
    }


    #
    # Stage 5: Export Microsoft Entra conditional access policies
    #
    if ($RunAll -or $RunStages.Contains(5)) {
        Write-Host "`nSTAGE 5: Export Microsoft Entra conditional access policies ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null
        
        # Get conditional access policies, parse and output as HTML reports
        ExportM365-Caps -OutputFolder $script:exportTargetFolder.conditionalAccess

        # Export Authentication Method and Strength Policies
        ExportM365-AuthMethodPolicy | Out-Null
        ExportM365-AuthStrengthPolicy | Out-Null
    }


    #
    # Stage 6: Export Microsoft 365 audit and sign-in log history
    #
    if ($RunAll -or $RunStages.Contains(6)) {
        Write-Host "`nSTAGE 6: Export Microsoft 365 Audit and Sign-in log history ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

        ExportM365-ConnectMgGraph | Out-Null

        if (ExportM365-ConnectIPPSSession) {
            ExportM365-AuditLogSettings | Out-Null
            ExportM365-AuditLogRetention | Out-Null
        }

        if (ExportM365-ConnectExchangeOnline) {
            ExportM365-AuditLogUnified -limit 100 | Out-Null
        }

        ExportM365-AuditLogRisky

        ExportM365-AuditLogProvisioning -limit 1000
 
        ExportM365-AuditLogDirectory -limit 1000

        # Get oldest log to show retention
        ExportM365-AuditLogSignIn -LogFilter "oldest" -logFile $script:exportTarget.signInLogsOldest
        
        # Get M365 sign-in logs sample
        ExportM365-AuditLogSignIn -LogFilter $LogFilter -logFile $script:exportTarget.signInLogs
        

        # Get non-interactive log to show retention
        # @todo BETA API only. Review
       
        # ExportM365-AuditLogSignIn -LogFilter "nonInteractiveUser" | Export-CSV "$($script:exportTarget.signInLogsNonInteractive)" -NoTypeInformation
        # Write-Host "Exported M365 sign-in audit log (non-interactive) to $($script:exportTarget.signInLogsNonInteractive)" -ForegroundColor Green

        # Get interactive log to show retention
        # @todo BETA API only. Review
       
        # ExportM365-AuditLogSignIn -LogFilter "interactiveUser" | Export-CSV "$($script:exportTarget.signInLogsInteractive)" -NoTypeInformation
        # Write-Host "Exported M365 sign-in audit log (interactive) to $($script:exportTarget.signInLogsInteractive)" -ForegroundColor Green

        # Get servicePrincipal log to show retention
        # @todo BETA API only. Review
        
        #ExportM365-AuditLogSignIn -LogFilter "servicePrincipal" | Export-CSV "$($script:exportTarget.signInLogsServicePrincipal)" -NoTypeInformation
        #Write-Host "Exported M365 sign-in audit log (service principal) to $($script:exportTarget.signInLogsServicePrincipal)" -ForegroundColor Green

        # Get managedIdentity log to show retention
        # @todo BETA API only. Review
        
        # ExportM365-AuditLogSignIn -LogFilter "managedIdentity" | Export-CSV "$($script:exportTarget.signInLogsManagedIdentity)" -NoTypeInformation
        # Write-Host "Exported M365 sign-in audit log (managed identity) to $($script:exportTarget.signInLogsManagedIdentity)" -ForegroundColor Green
    }

    #
    # Stage 7: Export Microsoft Exchange Online Protection and Defender Threat Protection Policies
    #
    if ($RunAll -or $RunStages.Contains(7)) {
        Write-Host "`nSTAGE 7: Export Exchange Online Protection and Defender Threat Protection Policies" -ForegroundColor Cyan
        
        if (ExportM365-ConnectExchangeOnline) {
            ExportM365-DfoMalwarePolicy | Out-Null
            
            ExportM365-DfoPhishingPolicy | Out-Null

            ExportM365-DfoSpamPolicy | Out-Null

            ExportM365-DfoSafeLinksPolicy | Out-Null
 
            ExportM365-DfoSafeAttachmentsPolicy | Out-Null

            ExportM365-DfoQuarantinePolicy | Out-Null
        }
    }
}

end {
    Write-Host "`n`nCompleted $(Get-Date). Disconnecting from Microsoft Graph and Exchange Online..."
    ExportM365-Disconnect

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
# MIIVwAYJKoZIhvcNAQcCoIIVsTCCFa0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf9nHHodJE81MZHTATea5YFoM
# l4KgghIgMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGizCCBPOg
# AwIBAgIRAMPvgP+UjhAauZR/ln8blhIwDQYJKoZIhvcNAQEMBQAwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNjAeFw0yNjAyMjQwMDAwMDBaFw0y
# NzAyMjQyMzU5NTlaMHkxCzAJBgNVBAYTAkFVMRowGAYDVQQIDBFXZXN0ZXJuIEF1
# c3RyYWxpYTEmMCQGA1UECgwdT2ZmaWNlIG9mIHRoZSBBdWRpdG9yIEdlbmVyYWwx
# JjAkBgNVBAMMHU9mZmljZSBvZiB0aGUgQXVkaXRvciBHZW5lcmFsMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+YIsf5QEHPChR8q7XiDlNhV0P1bd5SfW
# x++KrzoHMTf6v612Seg0vZ0Kn+WjG7jhD/gC7HGTPk5zTUBIJZdsdJZwkXIIHwNT
# nXfr5gxJDZTql/chL7jaiVlILlUGesTC6mUCl4/d5Bq9ehTAG9dcSu4jM4WEqVoN
# eQQXE5ry244xbkYWYufZdCSdlyQIrbu7Ixi63XFMTd/X8wZ72ZXlhi376xHkatl0
# 0uni9yess3AKPFxPR9NeBGmWGnFQoXa3w/lX314YQY7/Bng1c4LBJUsNgxlu2Geu
# gvWF/xd2vGZ2zF0hhE7950UMQSvouKL+YVXcDWVHPIHAOkntN6AWUiRSQhx1QwJw
# NIA8bok71LfTBG5D9scg5oQ8vlEwYQaweR4u2PJJ9lkMbPg0X3nhDGnTxsBZIVHB
# RboPBHrp9UYHdxAZC34UoNWRSa4I169s1q5wEsEvwKDQlPSVFSxz3sbIkC29L1Z7
# 7T1PjLpoyXRdi/mLeuhoqVAlISv5IE1BKQX9S0tvq0T2FqKZAZH3W+/kRcfnfOXE
# Q5Uqa2BUpK1Lb+b7/EKNAoViqvNazxiOpvOZvyGi6DYngHWBG4g+bgqAuQt63J1U
# 659c3C/pOo/FNqTIPGrXvs8SL7nK6Los0PgEOx6aB0jDL7E8XuRyts8CXvdA81NS
# FUCzKSg0CwcCAwEAAaOCAbEwggGtMB8GA1UdIwQYMBaAFA8qyyCHKLjsb0iuK1Sm
# KaoXpM0MMB0GA1UdDgQWBBS/UhfYBQcgqF3O0E4jErevoLqJgjAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzBKBgNVHSAE
# QzBBMDUGDCsGAQQBsjEBAgEDAjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3Rp
# Z28uY29tL0NQUzAIBgZngQwBBAEwSQYDVR0fBEIwQDA+oDygOoY4aHR0cDovL2Ny
# bC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcmww
# eQYIKwYBBQUHAQEEbTBrMEQGCCsGAQUFBzAChjhodHRwOi8vY3J0LnNlY3RpZ28u
# Y29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNydDAjBggrBgEFBQcw
# AYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wJgYDVR0RBB8wHYEbU2VydmljZURl
# c2tAYXVkaXQud2EuZ292LmF1MA0GCSqGSIb3DQEBDAUAA4IBgQBNeuToz2VT7Zhl
# 5/3bOAe9QX2c9oLFykvP/OVqzObznKfY+ELIl1Sj0JfHU3Y1VJlpPdG31v8t8yC3
# WV8tWa1feKwHnecpSTYK1e/6jkTCuT00MsPAm/m4ki6FEJjfcAqc4Mu+cH6he6AL
# IPMRbf0IgaElJWF1I/EgukC6+BFNfZz47Hu/7Pe6dk4e5rcz7wtP2EBxKppcnLnI
# 8HN4Jp6srDMgBloN0fhhNNHsAyJ2joGjFLrReaBBruynjd69D6phpYKe4h1GO3RM
# 2R4O3ZdRBTMRVf2sI0FcKikO6L+d9KMCuKTC/9NMsfHS5rNTQRKM2eyuIlSG8gfv
# WQJtKXnzRtAcCNgD1t+jA1y/iSvSraspgg2D5pYgbVlMFykujRrOn4FFVGUK1rem
# 26gTCfq9WpF7MulqmMLpt6I8FvFFTfnvVfouc+nwLDCm6NBNuOHgeqz33hmxkv9r
# H8TIFWT1dTOC8Lu16s7PM0fGE0DTaIt5eeRHkutYrWTcbrcMFrYxggMKMIIDBgIB
# ATBpMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYCEQDD74D/
# lI4QGrmUf5Z/G5YSMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSsLI6l6pk/N82qCiDBC/QhH5nd
# +zANBgkqhkiG9w0BAQEFAASCAgD1lSihWqoPFwaVxVSR+8ms6n/ALJbvTRE/eHod
# SS/F83j9PMyaolujO1vL3BD9Jy+sKvFXTtC1d+BxUwRv1R5mEAgCmRGiWfo6S2G6
# eJMAbcCIt7KU6Sz7vXm2ACmrQtMWi/LeUBKBceAhVXiWr7aZNfT1FLfpVZc/bfNL
# 8koMKB5voB5CZWmDFTKXlTXeYIxFxjU11cRHDbikVXq3IAzZ3fzJjyLs57lVDR9T
# aGIyZli3UL2tOq/2RrrDq309OEZ+znZUuDx4nRS4hArTjqZtglODuvFbBuZp0uC2
# 9jiO85MN4oIp9ZS1bDGNfVlYKQ27f1qHEvRczJajOij4BZMoyeNEfCpl3L6zQ6SP
# Afe3+j/WuB6X7emq4d2/LanNy0uLMngn7Dghdnkfxp1A1JLb/r0d5R5jdFO1rSQP
# 6/FC3kYGx3aiyuqgB2xqT1hzTUc16MwNNUYABz9qpMUXOQhcI2VbZ8OCxpa+IZ1g
# 37y6+JK8wXP0DJGyyfTLJT1rBIEtn9TuanpP+U7zC2MbvFfTgf4NTulABsDrCgDy
# u/Bwx1kXo4eoVQrA/Fc+cJX7o27qYsnzt/MaFNQ5/d2ib5yQtGBwHn6o8xD6350S
# EandLKzKvEUXUvTJCDzZD0+AurrNKSSMvg+azqxLgNVihk2M1AJTgQ1pbV0/D6yr
# Pb/MWA==
# SIG # End signature block
