function capPolicySummaryGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Build object summarising conditional access policy properties

    LOGIC
    Straight-line transformation with no branching.

    PARAMETERS
    -policy (required)

    RUNNING CONTEXT
    Called by  : capPolicyExport
    Calls      : no other functions in this module

    CMLETS/PERMISSIONS/SCOPES
    None. Operates on data already retrieved.

    --------------------------------------------------------------------------------#>


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
