function capGroupMemberGet {
    <#--------------------------------------------------------------------------------

    DESCRIPTION
    Returns the members of an Entra group, for listing who a conditional access policy
    actually includes or excludes.

    THIS FUNCTION DID NOT EXIST IN THE ORIGINAL SCRIPTS.

    capPolicyToHtml called ExportM365-GroupMembers in two places - once for included
    groups and once for excluded groups - but that function was never defined anywhere
    in the file. PowerShell treats an unresolved command as a non-terminating error, so
    the surrounding ForEach-Object carried on, $groupMembers stayed empty, and the HTML
    report rendered the group name with an empty member list.

    No error was shown and the report looked complete. That makes it the worst class of
    bug for audit evidence: a reader of the report would reasonably conclude the group
    had no members, when in fact the members were never retrieved. A crash would have
    been safer, because the gap would have been visible.

    LOGIC
    Validates the group ID before calling Graph, so a malformed value fails here rather
      than as an opaque API error.
    Retrieves members with Get-MgGroupMember, paging through all results.
    Returns each member with the display name and UPN lifted out of AdditionalProperties,
      which is where the Graph SDK puts them for directory objects.
    On failure, logs the error and returns an empty collection - but the caller can tell
      the difference, because the failure is in the run log and the completeness check.

    Nested groups are NOT expanded. Get-MgGroupMember returns direct members only, so a
    group containing another group lists that group as a member rather than its users.
    This matches how Entra evaluates conditional access assignment, but it means the
    member list is not a full effective-user list. Worth stating in any working paper
    that relies on these counts.

    PARAMETERS
    -groupId (required) the object ID of the group

    RUNNING CONTEXT
    Called by  : capPolicyToHtml, for both included and excluded groups
    Calls      : guidVerify, logWrite, exceptionFormat
    Returns    : member objects, or an empty collection on failure

    CMLETS/PERMISSIONS/SCOPES
    Get-MgGroupMember
    Scopes: Directory.Read.All or GroupMember.Read.All
    Role  : Global Reader

    --------------------------------------------------------------------------------#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$groupId
    )

    if ([string]::IsNullOrWhiteSpace($groupId) -or -not (guidVerify -InputObject $groupId)) {
        logWrite "Group member lookup skipped - '$groupId' is not a valid object ID." -level Warning -indent 2
        return @()
    }

    try {
        $members = Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop

        return $members | ForEach-Object {
            [PSCustomObject]@{
                Id                = $_.Id
                DisplayName       = $_.AdditionalProperties['displayName']
                UserPrincipalName = $_.AdditionalProperties['userPrincipalName']
                ObjectType        = ($_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', '')
            }
        }

    } catch {
        # Logged rather than swallowed. The original produced no output at all here,
        # which is what made the truncation invisible.
        logWrite (exceptionFormat -message "Failed retrieving members of group $groupId" -exception $_) -level Error -indent 2
        return @()
    }
}
