<#
.SYNOPSIS
OAG export of Microsoft 365 users, rights, authentication settings and audit log history for testing

.DESCRIPTION
Uses the Microsoft.Graph PowerShell module to export Azure AD users, user MFA settings, Azure AD sign-in history, conditional access policies, and privileged users who are members of built-in roles.

Based on the following:
    Stage 1: Export details of enabled Microsoft Azure AD users
              Scopes: AuditLog.Read.All, Organization.Read.All
              Author: Mike Crowley
                 URL: https://mikecrowley.us/2021/10/28/azure-ad-sign-in-activity-report-via-get-mguser/

             Export details of Microsoft Azure AD Service Principals with application role assignments or oAuth2 grants
              Scopes: Application.Read.All
              Author: OAG

    Stage 2: Export MFA status of enabled Microsoft Azure AD users
              Scopes: Reports.Read.All
              Author: OAG
                 URL: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.reports/get-mgreportauthenticationmethoduserregistrationdetail

    Stage 3: Export Microsoft 365 built-in role memberships
              Scopes: RoleEligibilitySchedule.Read.Directory, RoleAssignmentSchedule.Read.Directory, CrossTenantInformation.ReadBasic.All, AuditLog.Read.All, User.Read.All
              Author: Christoph Burmeister
                 URL: https://itinsights.org/azure-ad-list-role-assignments/

    Stage 4: Export Microsoft Azure AD Conditional Access Policies
               Scope: Application.Read.All, Group.Read.All, Policy.Read.All, RoleManagement.Read.Directory, User.Read.All
              Author: Nicola Suter, OAG
                 URL: https://github.com/nicolonsky/ConditionalAccessDocumentation

    Stage 5: Export Microsoft Azure AD Tenant Sign-in Logs (optional)
               Scope: AuditLog.Read.All
              Author: OAG

.COMPONENT 
Requires Module Microsoft.Graph

.INPUTS
AzureAD credentials to authorise required API scopes

.OUTPUTS
CSV exports of AzureAD Users, built-in roles and conditional access policies

.EXAMPLE
Export-M365.ps1 

.EXAMPLE
Export-M365.ps1 -Output C:\temp

.EXAMPLE
Export-M365.ps1 -LogFilter All

.PARAMETER Output
The name of the folder to output exported information to. Defaults to the running scripts folder ($PSScriptRoot).

.PARAMETER LogFilter
The type of filter to apply when exporting log entries. Defaults to the top 10,000

.NOTES
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
  Get-MgUser
  Get-MgReportAuthenticationMethodUserRegistrationDetail
  Get-MgGroup
  Get-MgServicePrincipal
  Get-MgServicePrincipalAppRoleAssignment
  Get-MgServicePrincipalOauth2PermissionGrant
  Get-MgDirectoryObject
  Get-MgRoleManagementDirectoryRoleDefinition
  Get-MgRoleManagementDirectoryRoleEligibilitySchedule
  Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance
  Get-MgIdentityConditionalAccessPolicy
  Get-MgIdentityConditionalAccessNamedLocation
  Get-MgDirectoryRoleTemplate
  Get-MgAuditLogSignIn

.LINK
Based on:
   https://mikecrowley.us/2021/10/28/azure-ad-sign-in-activity-report-via-get-mguser/
   https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell
   https://itinsights.org/azure-ad-list-role-assignments/
   https://github.com/nicolonsky/ConditionalAccessDocumentation
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false)][String]$Output = $PSScriptRoot,
	[Parameter(Mandatory = $false)][ValidateSet('Top10k','All','None')][AllowNull()][String]$LogFilter = "Top10k"
)

begin {
    clear

    $Output = ($Output | Resolve-Path).Path
    $execDateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")

    $OutputFolder = New-Item -ItemType Directory -Force -Path "$($Output)\M365\$($execDateTime)"
    $csvSignInLogs = Join-Path -Path $OutputFolder -ChildPath "m365_SignInLogs.csv"
    $csvBuiltInRoleMembers = Join-Path -Path $OutputFolder -ChildPath "m365_BuiltInRoleMembers.csv"
    $csvUserMFA = Join-Path -Path $OutputFolder -ChildPath "m365_UserMFA.csv"
    $csvUsers = Join-Path -Path $OutputFolder -ChildPath "m365_Users.csv"
    $csvServicePrincipals = Join-Path -Path $OutputFolder -ChildPath "m365_ServicePrincipals.csv"

    $psModulesRequired = @(
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

    $LegacyAuthClients = @(
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



    Start-Transcript -Path (Join-Path -Path $Output -ChildPath "$($execDateTime)_Export-M365_transcript.txt")
    Write-Host "Starting $(Get-Date)"

    Write-Host "`n----------------------------------------------"
    Write-Host " This script will output audit evidence to a series of subfolders under the location '$Output'"
    if ((Get-ChildItem $Output  | Measure-Object ).count -gt 4) {
        Write-Host "   - The script should be run from an empty folder to avoid confusion when providing the exported evidence" -ForegroundColor Yellow
    }

    $psModulesRequired | foreach {
        if (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue) {
            Write-Host "   - Found required PowerShell module '$_'. If exceptions occur check that module is up-to-date." -ForegroundColor Green
        } else {
            Write-Host "   - Missing required PowerShell module '$_'. Abort now." -ForegroundColor Red
        }    
    }
    	
    if ($ExecutionContext.SessionState.LanguageMode -eq "ConstrainedLanguage") {
	    Write-Host "   - ERROR. This script cannot run in Constrained Language Mode. Abort now." -ForegroundColor Red
    } elseif ($ExecutionContext.SessionState.LanguageMode -eq "FullLanguage") {
	    Write-Host "   - Running in Full Language Mode" -ForegroundColor Green
    } else {		
	    Write-Host "   - Caution, '$($ExecutionContext.SessionState.LanguageMode)' Language Mode detected" -ForegroundColor DarkYellow
    }

    Write-Host "----------------------------------------------`n"
    Pause


    Write-Host "`nScript Parameters:" -ForegroundColor Cyan
    Write-Host "  Output Folder: $Output"
    Write-Host "  Log Filter: $LogFilter"
    Write-Host "`n`n"





    function Connect_MgGraph {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
	   
        Write-Host "Connecting to Microsoft Graph..."
        Connect-MgGraph -ContextScope Process -Scopes "User.Read.All", "Reports.Read.All", "RoleEligibilitySchedule.Read.Directory", "RoleAssignmentSchedule.Read.Directory", "CrossTenantInformation.ReadBasic.All", "AuditLog.Read.All", "Application.Read.All", "Group.Read.All", "Policy.Read.All", "RoleManagement.Read.Directory", "Organization.Read.All"
    }

    Connect_MgGraph

    if ((Get-MgContext) -ne "") {
        Write-Host Connected to Microsoft Graph PowerShell using (Get-MgContext).Account account -ForegroundColor Yellow
    }
}

process {
#
# Stage 1: Export Microsoft Azure AD user information
#
Write-Host "`nSTAGE 1: Getting M365 users and service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan
$users = Get-MgUser -All -Filter 'accountEnabled eq true' -ConsistencyLevel eventual -Property @("DisplayName", "UserPrincipalName", "SignInActivity", "UserType", "Mail", "AssignedLicenses", "AccountEnabled", "EmployeeId", "UsageLocation", "OnPremisesLastSyncDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "LastPasswordChangeDateTime", "PasswordPolicies", "CreatedDateTime", "CreationType", "DeletedDateTime", "id")

$userCount = $users.count
Write-Host "Found $userCount M365 users"

$users | Select-Object @("DisplayName", "UserPrincipalName", @{n='LastSignInDateTime'; e={[datetime]$_.SignInActivity.LastSignInDateTime}}, @{n='lastNonInteractiveSignInDateTime'; e={[datetime]$_.SignInActivity.AdditionalProperties.lastNonInteractiveSignInDateTime}}, "UserType", "Mail", "AccountEnabled", "EmployeeId", "UsageLocation", "OnPremisesLastSyncDateTime", "OnPremisesDistinguishedName", "LastPasswordChangeDateTime", "PasswordPolicies", "CreatedDateTime", "CreationType", "DeletedDateTime", "id") |
	Export-CSV $csvUsers -NoTypeInformation
Write-Host "Exported M365 users to CSV $csvUsers" -ForegroundColor Green


function Export-M365-ServicePrincipalRights {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][String]$csvServicePrincipals
	)

    $pCache = @{}
    $scopeCache = @{}
    $spCache = @{}
    $out = @()
    $servicePrincipals = Get-MgServicePrincipal -All
    $spCount = $servicePrincipals.Count
    $spProcessed = 0
    Write-Host "`nFound $($spCount) service principals ($((Get-Date).ToString("yyyyMMdd_HHmmss")))"

    $servicePrincipals | ForEach-Object {
        $sp = $_
    
        Write-Verbose " - Processing assigned app roles and permission grants for $($sp.DisplayName) ($($sp.Id))"

        $spProcessed++
        $completed = [math]::Ceiling(($spProcessed / $spCount) * 100)
        Write-Progress -Activity "Processing $spProcessed of $spCount '$($sp.DisplayName)'" -PercentComplete $completed 

        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -All -ServicePrincipalId $_.Id
        $appRolesCount = $appRoleAssignments.Count
        $appRolesProcessed = 0
        if ($appRolesCount -ge 1) {
            Write-Host "    + Found $($appRolesCount) application role assignments for $($($sp.DisplayName))"
        }
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

            $out += $item
        }

        $grants = Get-MgServicePrincipalOauth2PermissionGrant -All -ServicePrincipalId $_.Id
        $grantsCount = $grants.Count
        $grantsProcessed = 0
        if ($grantsCount -ge 1) {
            Write-Host "    + Found $($grantsCount) permission grants for $($($sp.DisplayName))"
        }
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
                    $out += $item    
                }
            }
        }
    }

    $out | Export-CSV $csvServicePrincipals -NoTypeInformation
    Write-Host "Exported M365 service principals to CSV $csvServicePrincipals" -ForegroundColor Green
}
Export-M365-ServicePrincipalRights -csvServicePrincipals $csvServicePrincipals


#
# Stage 2: Export enabled Microsoft Azure AD user MFA status
#
Write-Host "`nSTAGE 2: Retrieving M365 user MFA settings ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

function Export-M365-MFA {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][String]$csvUserMFA
	)

    $Result = Get-MgReportAuthenticationMethodUserRegistrationDetail -All:$true -Property * |
        Select-Object UserDisplayName, UserPrincipalName, UserType, `
             IsAdmin, IsMfaCapable, IsMfaRegistered, IsPasswordlessCapable, IsSsprCapable, IsSsprEnabled, IsSsprRegistered, IsSystemPreferredAuthenticationMethodEnabled, `
             @{n='MethodsRegistered';e={($_.MethodsRegistered -join ", ")}}, `
             @{n='SystemPreferredAuthenticationMethods';e={($_.SystemPreferredAuthenticationMethods -join ", ")}}, @{n='UserPreferredMethodForSecondaryAuthentication';e={($_.UserPreferredMethodForSecondaryAuthentication -join ", ")}}, `
             LastUpdatedDateTime, Id

	Write-Host "Found MFA settings for $(($Result | Measure-Object).Count) users"

    $Result | Export-Csv -Path $csvUserMFA -NoTypeInformation -Append
	
	Write-Host "Exported MFA settings to $csvUserMFA" -ForegroundColor Green
}

Export-M365-MFA -csvUserMFA $csvUserMFA





# Stage 3 and 4 supporting functions

function Test-Guid {
    <#
    .SYNOPSIS
    Validates a given input string and checks string is a valid GUID
    .DESCRIPTION
    Validates a given input string and checks string is a valid GUID by using the .NET method Guid.TryParse
    .EXAMPLE
    Test-Guid -InputObject "3363e9e1-00d8-45a1-9c0c-b93ee03f8c13"
    .NOTES
    Uses .NET method [guid]::TryParse()
    #>
    [Cmdletbinding()]
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )
    process {
        return [guid]::TryParse($InputObject, $([ref][guid]::Empty))
    }
}

$displayNameCache = @{}

function Resolve-MgObject {
    <#
    .SYNOPSIS
    Resolve a Microsoft Graph item to display name
    .DESCRIPTION
    Resolves a Microsoft Graph Directory Object to a Display Name when possible
    .EXAMPLE
    
    .NOTES
    
    #>
    [Cmdletbinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )
    process {
        if (Test-Guid -InputObject $InputObject) {
            try {
                # use hashtable as cache to limit API calls
                if ($displayNameCache.ContainsKey($InputObject)) {
                    Write-Debug "Cached display name for `"$InputObject`""
                    Write-Verbose " - GUID Lookup: $(($displayNameCache[$InputObject]).AdditionalProperties.displayName)"
                    return $displayNameCache[$InputObject]
                }
                else {
                    $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $InputObject -ErrorAction Stop
                    $displayName = $directoryObject.AdditionalProperties["displayName"]
                    $displayNameCache[$InputObject] = $directoryObject
                    Write-Verbose " - GUID Lookup: $(($displayNameCache[$InputObject]).AdditionalProperties.displayName)"
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




#
# Stage 3: Export Microsoft 365 built-in role memberships
#
Write-Host "`nSTAGE 3: Retrieving M365 user built-in role memberships ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

function Export-M365-Roles {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][String]$csvBuiltInRoleMembers,
		[Parameter(Mandatory = $true)]$users
	)
	
	Write-Host "Found $(($users | Measure-Object).count) users"

	# get all groups to resolve IDs
	$groups = Get-MgGroup -All
	Write-Host "Found $(($groups | Measure-Object).count) groups"
	
	# get all Azure AD role definitions to resolve IDs
	$roles = Get-MgRoleManagementDirectoryRoleDefinition
	Write-Host "Found $(($roles | Measure-Object).count) roles"
	
	[System.Collections.ArrayList]$resolved_assignments = @()

	# get all eligible role assignments
	$eligible_role_assignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All:$true	
	$ProcessedRoleCount = 0
	$RolesToProcess = ($eligible_role_assignments | Measure-Object).Count
	Write-Host "`nFound $RolesToProcess eligible for admin roles"
	foreach ($assignment in $eligible_role_assignments) {
        $object = Resolve-MgObject -InputObject $assignment.PrincipalId

        $dataType = $object.AdditionalProperties.'@odata.type'
        $displayName = $object.AdditionalProperties.displayName
        $userPrincipalName = $object.AdditionalProperties.userPrincipalName

		$obj = [pscustomobject]@{
			'odata.type'   = $dataType
			'membership'   = "eligible"
			'role'         = $roles | Where-Object { $_.id -eq $assignment.RoleDefinitionId } | Select-Object -ExpandProperty DisplayName
			'servicePrincpal'         = if ($dataType -eq '#microsoft.graph.serviceprincipal') { $object.AdditionalProperties.appDisplayName } else { "" }
			'user'         = if ($dataType -eq '#microsoft.graph.user') { $userPrincipalName } else { "" }
			'group'        = if ($dataType -eq '#microsoft.graph.group') { $displayName } else { "" }
			'enabled'      = if ($object.AdditionalProperties.accountEnabled -eq $True) { "Enabled" } elseif($object.AdditionalProperties.accountEnabled -eq $False) { "Disabled"} else { "" }
			'principalId'  = $assignment.PrincipalId
		}
		
		$resolved_assignments.Add($obj) | Out-Null
		
		$ProcessedRoleCount++
		$completed = [math]::Ceiling(($ProcessedRoleCount / $RolesToProcess) * 100)
		Write-Progress -Activity "Processed eligible roles: $ProcessedRoleCount of $RolesToProcess" -PercentComplete $completed 
	}
	
	# get all assigned role assignments
	$assigned_role_assignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All:$true
	$ProcessedRoleCount = 0
	$RolesToProcess = ($assigned_role_assignments | Measure-Object).Count
	Write-Host "`nFound $RolesToProcess assigned admin roles"
	foreach ($assignment in $assigned_role_assignments) {
        $object = Resolve-MgObject -InputObject $assignment.PrincipalId

        $dataType = $object.AdditionalProperties.'@odata.type'
        $displayName = $object.AdditionalProperties.displayName
        $userPrincipalName = $object.AdditionalProperties.userPrincipalName

		$obj = [pscustomobject]@{
			'odata.type'   = $dataType
			'membership'   = "assigned"
			'role'         = $roles | Where-Object { $_.id -eq $assignment.RoleDefinitionId } | Select-Object -ExpandProperty DisplayName
			'servicePrincpal'         = if ($dataType -eq '#microsoft.graph.serviceprincipal') { $object.AdditionalProperties.appDisplayName } else { "" }
			'user'         = if ($dataType -eq '#microsoft.graph.user') { $userPrincipalName } else { "" }
			'group'        = if ($dataType -eq '#microsoft.graph.group') { $displayName } else { "" }
			'enabled'      = if ($object.AdditionalProperties.accountEnabled -eq $True) { "Enabled" } elseif($object.AdditionalProperties.accountEnabled -eq $False) { "Disabled"} else { "" }
			'principalId'  = $assignment.PrincipalId
		}

		$resolved_assignments.Add($obj) | Out-Null

		$ProcessedRoleCount++
		$completed = [math]::Ceiling(($ProcessedRoleCount / $RolesToProcess) * 100)
		Write-Progress -Activity "Processed assigned roles: $ProcessedRoleCount of $RolesToProcess" -PercentComplete $completed 
	}
	
	# get all permanent role assignments
	$permanent_role_assignments = Get-MgRoleManagementDirectoryRoleAssignment -All:$true
	$ProcessedRoleCount = 0
	$RolesToProcess = ($permanent_role_assignments | Measure-Object).Count
	Write-Host "`nFound $RolesToProcess permanent admin role assignments"
	foreach ($assignment in $permanent_role_assignments) {
        $object = Resolve-MgObject -InputObject $assignment.PrincipalId   #Get-MgDirectoryObject -DirectoryObjectId $assignment.PrincipalId

        $dataType = $object.AdditionalProperties.'@odata.type'
        $displayName = $object.AdditionalProperties.displayName
        $userPrincipalName = $object.AdditionalProperties.userPrincipalName

		$obj = [pscustomobject]@{
			'odata.type'   = $dataType
			'membership'   = "permanent"
			'role'         = $roles | Where-Object { $_.id -eq $assignment.RoleDefinitionId } | Select-Object -ExpandProperty DisplayName 
			'servicePrincpal'         = if ($dataType -eq '#microsoft.graph.serviceprincipal') { $object.AdditionalProperties.appDisplayName } else { "" }
			'user'         = if ($dataType -eq '#microsoft.graph.user') { $userPrincipalName } else { "" }
			'group'        = if ($dataType -eq '#microsoft.graph.group') { $displayName } else { "" }
			'enabled'      = if ($object.AdditionalProperties.accountEnabled -eq $True) { "Enabled" } elseif($object.AdditionalProperties.accountEnabled -eq $False) { "Disabled"} else { "" }
			'principalId'  = $assignment.PrincipalId
		}

		$resolved_assignments.Add($obj) | Out-Null

		$ProcessedRoleCount++
		$completed = [math]::Ceiling(($ProcessedRoleCount / $RolesToProcess) * 100)
		Write-Progress -Activity "Processed assigned roles: $ProcessedRoleCount of $RolesToProcess" -PercentComplete $completed 
	}

	$resolved_assignments | Export-CSV $csvBuiltInRoleMembers -NoTypeInformation
	Write-Host "Exported built-in role assignments to $csvbuiltInRoleMembers" -ForegroundColor Green
}

Export-M365-Roles -csvBuiltInRoleMembers $csvBuiltInRoleMembers -users $users

Remove-Variable users





function BuildCapHtml {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)]$policy,
        $includeServicePrincipals,
        $excludeServicePrincipals,
        $includeGroups,
        $excludeGroups,
        $includeRoles,
        $excludeRoles,
        $includeUsers,
        $excludeUsers,
        $includeLocations,
        $excludeLocations,
        $namedLocations,
        $includeApps,
        $excludeApps,
        $includeAuthenticationContext
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
$out_includeApps = $includeApps -join $separator
$out_excludeApps = $excludeApps -join $separator

$out_includeUserActions = ""
if ($condition.Applications.includeUserActions) {
    $out_includeUserActions = $condition.Applications.includeUserActions -join $separator
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

foreach ($location in $includeLocations) {
    $thisLocation = $namedLocations | Where-Object {$_.Id -eq $location }

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
        $out_ipRanges = ($thisLocation.AdditionalProperties['ipRanges'] | Sort-Object | Out-String)
    }

    $out_countriesAndRegions = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('countriesAndRegions')) {
        $out_countriesAndRegions = ($thisLocation.AdditionalProperties['countriesAndRegions'] | Sort-Object | Out-String )
    }

    $out_includeUnknownCountriesAndRegions = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('includeUnknownCountriesAndRegions')) {
        $out_cidrAddress = ($thisLocation.AdditionalProperties['includeUnknownCountriesAndRegions'] | Sort-Object | Out-String)
    }

    $outHtml = $outHtml + '<tr><td>{0}<br/><nobr>{1}<nobr></td><td>IncludeLocations</td><td><p>CIDR Addresses: {2} </p><p>IP Ranges: {3}</p><p>Countries and Regions: {4} </p><p>Include Unknown Countries and Regions: {5}</p></td></tr>' -f $out_displayName, $location, $out_cidrAddress, $out_ipRanges, $out_countriesAndRegions, $out_includeUnknownCountriesAndRegions
}

foreach ($location in $excludeLocations) {
    $thisLocation = $namedLocations | Where-Object {$_.Id -eq $location }

    $out_displayName = "(un-named)"
    if ($thisLocation.DisplayName) {
        $out_displayName = $thisLocation.DisplayName
    }

    $out_cidrAddress = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('cidrAddress')) {
        $out_cidrAddress = ($thisLocation.AdditionalProperties['cidrAddress'] | Sort-Object | Out-String)
    }

    $out_ipRanges = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('ipRanges')) {
        $out_ipRanges = ($thisLocation.AdditionalProperties['ipRanges'] | Sort-Object | Out-String)
    }

    $out_countriesAndRegions = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('countriesAndRegions')) {
        $out_countriesAndRegions = ($thisLocation.AdditionalProperties['countriesAndRegions'] | Sort-Object | Out-String)
    }

    $out_includeUnknownCountriesAndRegions = ""
    if ($thisLocation.AdditionalProperties -and $thisLocation.AdditionalProperties.ContainsKey('includeUnknownCountriesAndRegions')) {
        $out_cidrAddress = ($thisLocation.AdditionalProperties['includeUnknownCountriesAndRegions'] | Sort-Object | Out-String)
    }

    $outHtml = $outHtml + '<tr><td>{0}<br/><nobr>{1}<nobr></td><td>ExcludeLocations</td><td><p>CIDR Addresses: {2} </p><p>IP Ranges: {3}</p><p>Countries and Regions: {4} </p><p>Include Unknown Countries and Regions: {5}</p></td></tr>' -f $out_displayName, $location, $out_cidrAddress, $out_ipRanges, $out_countriesAndRegions, $out_includeUnknownCountriesAndRegions
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
            <td>$($includeGroups -join $separator)</td>
            <td>$($excludeGroups -join $separator)</td>
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
    $out_signInFrequency = $policy.SessionControls.SignInFrequency.Value + " " + $policy.SessionControls.SignInFrequency.Type
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

</body>
</html>
'@ -f $out_applicationEnforcedRestrictions, $out_cloudSecurityIsEnabled, $out_disableResilienceDefaults, $out_persistentBrowserMode, $out_signInFrequency
} catch {
    #Throw $_
    $_ | Format-List * -Force | Out-String ; $_.InvocationInfo | Format-List * -Force | Out-String
}

return $outHtml
}



#
# Stage 4: Export Microsoft Azure AD conditional access policies
#
Write-Host "`nSTAGE 4: Export Microsoft Azure AD conditional access policies ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-Graph -Scopes `"Application.Read.All`", `"Group.Read.All`", `"Policy.Read.All`", `"RoleManagement.Read.Directory`", `"User.Read.All`""
}

function Export-M365-CAPs {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)]$OutputFolder
	)
	
	Write-Progress -PercentComplete -1 -Activity "Fetching conditional access policies and related data from Graph API"

	# Get Conditional Access Policies
	$conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -ExpandProperty "*" -All -ErrorAction Stop
	Write-Host "Found $($conditionalAccessPolicies.Count) conditional access policies"

	#Get Conditional Access Named / Trusted Locations
	$namedLocationsData = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
    
    <#$namedLocations = $namedLocationsData | Group-Object -Property Id -AsHashTable
	if (-not $namedLocations) { $namedLocations = @{} }#>

    Write-Host "Found $($namedLocationsData.Count) named locations"

	# Get Azure AD Directory Role Templates
	$directoryRoleTemplates = Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Group-Object -Property Id -AsHashTable
	Write-Host "Found $($directoryRoleTemplates.Count) directory role templates"

	# Service Principals
	$servicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop | Group-Object -Property AppId -AsHashTable
	Write-Host "Found $($servicePrincipals.Count) service principals"

	# Init report 
	$documentation = [System.Collections.Generic.List[Object]]::new()	

    Write-Host "Processing policy..."

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
		
		try {
			# Resolve object IDs of included users
			$includeUsers = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.IncludeUsers | ForEach-Object {
                if (Test-Guid -InputObject $PSItem) {
                    $user = Resolve-MgObject -InputObject $PSItem
				    $includeUsers.Add(($user.AdditionalProperties.userPrincipalName))
                } else {
                    Write-Verbose " - $($PSItem)"
                	$includeUsers.Add($PSItem)
                }
			}
			# Resolve object IDs of excluded users
			$excludeUsers = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.ExcludeUsers | ForEach-Object {
                if (Test-Guid -InputObject $PSItem) {
                    $user = Resolve-MgObject -InputObject $PSItem
				    $excludeUsers.Add(($user.AdditionalProperties.userPrincipalName))
                } else {
                    Write-Verbose " - $($PSItem)"
                	$excludeUsers.Add($PSItem)
                }
			}
			# Resolve object IDs of included groups
			$includeGroups = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.IncludeGroups | ForEach-Object {
                if (Test-Guid -InputObject $PSItem) {
                    $group = Resolve-MgObject -InputObject $PSItem
				    $includeGroups.Add(($group.AdditionalProperties.displayName))
                } else {
                    Write-Verbose " - $($PSItem)"
                	$includeGroups.Add($PSItem)
                }
			}
			# Resolve object IDs of excluded groups
			$excludeGroups = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.ExcludeGroups | ForEach-Object {
                if (Test-Guid -InputObject $PSItem) {
                    $group = Resolve-MgObject -InputObject $PSItem
				    $excludeGroups.Add(($group.AdditionalProperties.displayName))
                } else {
                    Write-Verbose " - $($PSItem)"
                	$excludeGroups.Add($PSItem)
                }
			}
			# Resolve object IDs of included roles
			$includeRoles = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.IncludeRoles | ForEach-Object {
				if ($directoryRoleTemplates.ContainsKey($PSItem)) {
					$includeRoles.Add(($directoryRoleTemplates[$PSItem].DisplayName))
				}
				else {
					$includeRoles.Add($PSItem)
				}
			}

			# Resolve object IDs of excluded roles
			$excludeRoles = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Users.ExcludeRoles | ForEach-Object {
				if ($directoryRoleTemplates.ContainsKey($PSItem)) {
					$excludeRoles.Add(($directoryRoleTemplates[$PSItem].DisplayName))
				}
				else {
					$excludeRoles.Add($PSItem)
				}
			}
			# Resolve object IDs of included apps
			$includeApps = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Applications.IncludeApplications | ForEach-Object {
				if ($servicePrincipals.ContainsKey($PSItem)) {
					$includeApps.Add(($servicePrincipals[$PSItem].DisplayName))
				}
				else {
					$includeApps.Add($PSItem)
				}
			}
			# Resolve object IDs of excluded apps
			$excludeApps = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Applications.ExcludeApplications | ForEach-Object {
				if ($servicePrincipals.ContainsKey($PSItem)) {
					$excludeApps.Add(($servicePrincipals[$PSItem].DisplayName))
				}
				else {
					$excludeApps.Add($PSItem)
				}
			}

			$includeServicePrincipals = [System.Collections.Generic.List[Object]]::new()
			$excludeServicePrincipals = [System.Collections.Generic.List[Object]]::new()

			$policy.Conditions.ClientApplications.IncludeServicePrincipals | ForEach-Object {
				if ((-not [string]::IsNullOrEmpty($PSItem)) -and $servicePrincipals.ContainsKey($PSItem)) {
					$includeServicePrincipals.Add(($servicePrincipals[$PSItem].DisplayName))
				}
				else {
					$includeServicePrincipals.Add($PSItem)
				}
			}
			$policy.Conditions.ClientApplications.ExcludeServicePrincipals | ForEach-Object {
				if ((-not [string]::IsNullOrEmpty($PSItem)) -and $servicePrincipals.ContainsKey($PSItem)) {
					$excludeServicePrincipals.Add(($servicePrincipals[$PSItem].DisplayName))
				}
				else {
					$excludeServicePrincipals.Add($PSItem)
				}
			}
			
			$includeAuthenticationContext = [System.Collections.Generic.List[Object]]::new()
			$policy.Conditions.Applications.IncludeAuthenticationContextClassReferences | ForEach-Object {
				$context = Get-MgIdentityConditionalAccessAuthenticationContextClassReference -Filter "Id eq '$PSItem'"
				$includeAuthenticationContext.Add($context.DisplayName)
			}

			# delimiter for arrays in csv report
			$separator = "`r`n"
			# when terms of use are present just add a generic hint.
			if ($policy.GrantControls.TermsOfUse) { $policy.GrantControls.BuiltInControls += "termsOfUse" }
			if ($policy.GrantControls.AuthenticationStrength) { $policy.GrantControls.BuiltInControls += "authenticationStrength" }

            try {
                $includeLocations = $policy.conditions.Locations.IncludeLocations
                $excludeLocations = $policy.conditions.Locations.ExcludeLocations
            } catch {
                $includeLocations = ""
                $excludeLocations = ""
            }

            $capExportFileName = ($policy.DisplayName) -replace '[^a-zA-Z0-9-_ ]', ' '
            $outHtml = BuildCapHtml -policy $policy `
                         -includeServicePrincipals $includeServicePrincipals -excludeServicePrincipals $excludeServicePrincipals `
                         -includeGroups $includeGroups -excludeGroups $excludeGroups `
                         -includeRoles $includeRoles -excludeRoles $excludeRoles `
                         -includeUsers $includeUsers -excludeUsers $excludeUsers `
                         -includeLocations $includeLocations -excludeLocations $excludeLocations -namedLocations $namedLocationsData `
                         -includeApps $includeApps -excludeApps $excludeApps `
                         -includeAuthenticationContext $includeAuthenticationContext
            $outHtml | Out-File (Join-Path -Path $OutputFolder -ChildPath "CAP_$($capExportFileName).html")
            Write-Host "    Conditional access policy written to 'CAP_$($capExportFileName).html'" -ForegroundColor Green
		} catch {
			#Throw $_
			#Write-Error $PSItem
            #$_ | Format-List * -Force | Out-String
            $_ | Format-List * -Force | Out-String ; $_.InvocationInfo | Format-List * -Force | Out-String
		}
	}
}

Export-M365-CAPs -OutputFolder $OutputFolder



function Export-M365-SignInLogs {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)][String]$csvSignInLogs,
    	[Parameter(Mandatory = $true)][String]$LogFilter
	)

    [array]$report = @()
    if ($LogFilter -eq "Top10k") {
        $logs = Get-MgAuditLogSignIn -Filter "status/errorCode eq 0" -Sort "createdDateTime DESC" -Top 10000
    } elseif ($LogFilter -eq "all") {
        $logs = Get-MgAuditLogSignIn -Filter "status/errorCode eq 0" -All $true
    } else {
        Write-Host "Skipping log export (LogFilter: $($LogFilter))"
        return $null
    }

    Write-Host "Found $($logs.Count) sign in logs"
	Write-Progress -PercentComplete -1 -Activity "Fetching audit log sign-in history from Graph API"

    foreach ($log in $logs) {
		# Display some progress (based on policy count)
		$currentIndex = $logs.indexOf($log) + 1

		$progress = @{
			Activity         = "Exporting Sign-in Logs..."
			PercentComplete  = [Decimal]::Divide($currentIndex, $logs.Count) * 100
			CurrentOperation = "Processing `"$(($log.CreatedDateTime).ToString("d MMMM yyyy"))`""
		}
		if ($currentIndex -eq $logs.Count) { $progress.Add("Completed", $true) }

		Write-Progress @progress

        if ($log.DeviceDetail.DisplayName -eq "" -and $log.DeviceDetail.Browser -ne "") {
            $device = "$($log.DeviceDetail.Browser)"
        } else {
            $device = "$($log.DeviceDetail.DisplayName) ($($log.DeviceDetail.TrustType))"
        }

        if ($LegacyAuthClients -contains $log.ClientAppUsed) {
            Write-Host " - Legacy auth client $($log.ClientAppUsed) ($($log.CreatedDateTime))" -ForegroundColor DarkYellow
            Write-Host "     + $($log.AppDisplayName)  ($($log.Id))" -ForegroundColor DarkYellow
        }

        $line = [pscustomobject][Ordered]@{ 
            Id = $log.Id
            SignInDateTime = $log.CreatedDateTime
            AppDisplayName = $log.AppDisplayName
            ClientAppUsed = $log.ClientAppUsed
            ClientCredentialType = $log.ClientCredentialType
            SignInEventTypes = $log.SignInEventTypes -join ", "
            StatusError = $log.Status.ErrorCode
            StatusDetail = $log.Status.AdditionalDetails
            MfaAuthMethod = $log.MfaDetail.AuthMethod
            UserDisplayName = $log.UserDisplayName
            UserPrincipalName = $log.UserPrincipalName
            UserId = $log.UserId
            UserType = $log.UserType
            RiskState = $log.RiskState
            RiskEventTypes = $log.RiskEventTypesV2 -join ", "
            RiskDuringSignIn = $log.RiskLevelDuringSignIn
            RiskLevelAggregated = $log.RiskLevelAggregated
            RiskDetail = $log.RiskDetail
            LocationCity = $log.Location.City
            LocationRegion = "$($log.Location.State), $($log.Location.CountryOrRegion)"
            Device = $device
            DeviceIsManaged = $log.DeviceDetail.IsManaged
            DeviceIsCompliant = $log.DeviceDetail.IsCompliant
            IPAddress = $log.IPAddress
            IsInteractive = $log.IsInterActive
            ConditionalAccessStatus = $log.ConditionalAccessStatus 
            ConditionalAccessPolicies = ($log.AppliedConditionalAccessPolicies | %{ "$($_.DisplayName) ($($_.Result))" }) -join ",`r`n "
        }
        
        $report += $line
    }

	# Export report as csv
	$report | Export-CSV $csvSignInLogs -NoTypeInformation
	Write-Host "Exported sign in logs to $csvSignInLogs`n" -ForegroundColor Green
}

#
# Stage 5: Export Microsoft Azure AD sign-in log history
#
Write-Host "`nSTAGE 5: Export Microsoft Azure AD Sign-in log history ($((Get-Date).ToString("yyyyMMdd_HHmmss")))" -ForegroundColor Cyan

Export-M365-SignInLogs -LogFilter $LogFilter -csvSignInLogs $csvSignInLogs
}

end {
    Write-Host "Completed $(Get-Date). Disconnecting from Microsoft Graph..."
    Disconnect-MgGraph

    Stop-Transcript
}