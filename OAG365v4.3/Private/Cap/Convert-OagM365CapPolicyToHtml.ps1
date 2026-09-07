function Convert-OagM365CapPolicyToHtml {
    <#
    .SYNOPSIS
    Build an HTML report from a conditional access policy

    .DESCRIPTION
    Build an HTML report from a conditional access policy

    .PARAMETER policy
    Conditional access policy object obtained from Microsoft Graph function Get-MgIdentityConditionalAccessPolicy

    .EXAMPLE
    Convert-OagM365CapPolicyToHtml -policy $policy

    .NOTES
        NAME: Convert-OagM365CapPolicyToHtml
        VERSION: 1.2

        FUNCTIONS & PERMISSIONS:
           Resolve-OagM365DirectoryObjectList
           Resolve-OagM365ServicePrincipalList
           Convert-OagM365CapLocationToHtml
           Resolve-OagM365DirectoryObjectList
           Resolve-OagM365DirectoryRoleList
           Test-OagM365Guid
           Resolve-OagM365DirectoryObject
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
$includeApps = Resolve-OagM365DirectoryObjectList -srcObject $policy.Conditions.Applications.IncludeApplications -getAttrib "displayName"
$excludeApps = Resolve-OagM365DirectoryObjectList -srcObject $policy.Conditions.Applications.ExcludeApplications -getAttrib "displayName"

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

$includeServicePrincipals = Resolve-OagM365ServicePrincipalList -srcObject $policy.Conditions.ClientApplications.IncludeServicePrincipals
$excludeServicePrincipals = Resolve-OagM365ServicePrincipalList -srcObject $policy.Conditions.ClientApplications.ExcludeServicePrincipals

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
    $outHtml += Convert-OagM365CapLocationToHtml -type "IncludeLocation" -locations $policy.conditions.Locations.IncludeLocations 
    $outHtml += Convert-OagM365CapLocationToHtml -type "ExcludeLocation" -locations $policy.conditions.Locations.ExcludeLocations 
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


$includeUsers = Resolve-OagM365DirectoryObjectList -srcObject $policy.Conditions.Users.IncludeUsers -getAttrib "userPrincipalName"
$excludeUsers = Resolve-OagM365DirectoryObjectList -srcObject $policy.Conditions.Users.ExcludeUsers -getAttrib "userPrincipalName"

# Resolve object IDs of included groups
$includeGroups = [System.Collections.ArrayList]::new()
$includeGroupsMembers = @{}
$policy.Conditions.Users.IncludeGroups | ForEach-Object {
    if (Test-OagM365Guid -InputObject $PSItem) {
        $group = Resolve-OagM365DirectoryObject -InputObject $PSItem
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
    if (Test-OagM365Guid -InputObject $PSItem) {
        $group = Resolve-OagM365DirectoryObject -InputObject $PSItem
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

$includeRoles = Resolve-OagM365DirectoryRoleList -srcObject $policy.Conditions.Users.IncludeRoles
$excludeRoles = Resolve-OagM365DirectoryRoleList -srcObject $policy.Conditions.Users.ExcludeRoles

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
