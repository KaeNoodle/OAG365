# Permissions

Least-privilege mapping for OAG-M365-AuditingScript.

Version 4.0.0 | OFFICIAL

---

## Summary

| Report | Directory roles | Runs read-only? |
|---|---|---|
| `Cap` | Global Reader | Yes |
| `Iam` | Global Reader + Reports Reader | Yes |
| `Org` | Global Reader | Partial — one export needs Global Administrator |
| `Dfo` | Global Reader, or EXO View-Only Organization Management | Yes |
| `ThreatHunt` | Security Reader | Yes, but needs admin consent first |

---

## The two gates

Graph permissions are checked twice, and both must pass. Confusing them is the most common cause
of an export that returns nothing while appearing to succeed.

**Gate 1 — application consent.** The OAuth scope the client application is permitted to request,
for example `Policy.Read.All`. Interactive runs use the pre-registered Microsoft Graph Command
Line Tools application; scopes not already consented for the tenant will prompt, and a
non-administrator cannot grant them.

**Gate 2 — the signed-in user's directory role.** What that user is actually permitted to see. A
granted scope does not override a missing role.

A scope can be granted and the call still fail, or return a filtered subset, because the user's
role does not cover the object. The run log records granted scopes for exactly this reason, but
`scopes-granted.csv` showing all green does not by itself prove the export was complete.

---

## Global Reader and Security Reader are not interchangeable

Neither role contains the other. This is worth stating plainly because the names suggest
otherwise and it has caused incorrect assumptions.

- **Global Reader** reads most of the directory and tenant configuration.
- **Security Reader** covers Microsoft Secure Score and parts of the Defender portal that Global
  Reader cannot reach, including advanced hunting.

The `ThreatHunt` report needs Security Reader. Assigning only Global Reader will not work.

A separate trap: "Security Reader" also exists as an **Azure RBAC role** for Azure resources.
That is a different role with a different scope and does not grant anything in Entra ID. When
requesting access, be explicit that you mean the Entra ID directory role.

---

## Recommended approach: a role-assignable group

Rather than assigning Global Reader directly to individual accounts, create a role-assignable
security group and assign the roles to it.

Suggested composition:

| Role | Why |
|---|---|
| Directory Readers | Baseline directory object read |
| Reports Reader | Sign-in and audit activity used by the `Iam` report |
| Security Reader | Secure Score and Defender advanced hunting for `ThreatHunt` |

This is narrower than Global Reader, and membership is auditable and time-bound in a way that a
direct role assignment is not. Where the audit requires the broader coverage Global Reader
provides, add it to the group rather than to the person.

Role-assignable groups must be created with `isAssignableToRole` set at creation. It cannot be
enabled afterwards, so the group must be created correctly the first time.

---

## Per-report detail

### Cap — Conditional access

**Role:** Global Reader

**Scopes:** `Policy.Read.All`, `Policy.Read.ConditionalAccess`,
`Policy.Read.AuthenticationMethod`, `AuthenticationContext.Read.All`,
`RoleManagement.Read.Directory`, `Application.Read.All`, `Directory.Read.All`

`RoleManagement.Read.Directory` and `Application.Read.All` are needed not for the policies
themselves but to resolve the GUIDs inside them into readable role and application names.
Without them the HTML reports render raw GUIDs, which is technically complete but not usable as
evidence.

### Iam — Identity and access

**Roles:** Global Reader + Reports Reader

**Scopes:** `Directory.Read.All`, `DeviceManagementApps.Read.All`, `Device.Read.All`,
`Application.Read.All`, `AuditLog.Read.All`, `RoleManagement.Read.Directory`,
`PrivilegedEligibilitySchedule.Read.AzureADGroup`,
`PrivilegedAssignmentSchedule.Read.AzureADGroup`, `RoleManagementPolicy.Read.AzureADGroup`

`AuditLog.Read.All` requires Reports Reader in addition to Global Reader. Without it the sign-in
activity columns are blank rather than the call failing, so this is a silent gap — check
`completeness.csv` if user activity data is missing.

The two `Privileged*.Read.AzureADGroup` scopes cover PIM group memberships. PIM for directory
roles is covered by `RoleManagement.Read.Directory`. These are separate and both are needed for
complete PIM evidence.

`RoleManagementPolicy.Read.AzureADGroup` is separate again, and covers the activation *rules* on
a PIM-managed group: approval, MFA on activation, maximum duration, justification. Without it
every group policy lookup returns 403 PermissionScopeNotGranted and the run records group
memberships with no evidence of the controls protecting them.

### Org — Tenant configuration

**Role:** Global Reader, except as noted

**Scopes:** `Organization.Read.All`, `Directory.Read.All`, `Domain.Read.All`,
`OnPremDirectorySynchronization.Read.All`

**Known gap.** `OnPremDirectorySynchronization.Read.All` requires **Global Administrator**. No
read-only role satisfies it — not Global Reader, not Security Reader, not Directory Readers. The
export returns 403 under least-privilege credentials.

Two options, both defensible:

1. Run the report with `-SkipOnPremSync` and record the limitation in the working paper.
2. Obtain a time-boxed Global Administrator assignment through PIM for a single run, and
   document the elevation.

Option 1 is usually appropriate. Elevating to Global Administrator to collect one configuration
export is difficult to justify to the client, and the same information can often be corroborated
from the on-premises side.

The MX, SPF, DKIM and DMARC lookups in this report are public DNS queries and need no tenant
permission at all.

### Dfo — Threat protection

**Role:** Exchange Online View-Only Organization Management, or Global Reader

**Scopes:** none — this report uses Exchange Online PowerShell, not Graph.

The only report not using Graph. It authenticates separately, which is why running `-Report Dfo`
alongside Graph reports produces two login prompts.

### ThreatHunt — Defender advanced hunting

**Role:** Security Reader

**Scopes:** `ThreatHunting.Read.All`

**Known gap.** This scope requires tenant administrator consent and is not consented by default.
Until an administrator grants it, all three hunting queries fail with 403 and the report produces
nothing. Version 4.0.0 detects this specifically and reports it as a permission problem rather
than a generic error.

Consent is granted once per tenant, not per user, so this only needs doing once — but it does
need doing before the audit, not during it.

---

## Verifying what you actually have

Before relying on any export, check the granted scopes:

```powershell
Import-Csv .\results\<RunId>\00_RunLog\scopes-granted.csv | Where-Object Granted -eq 'False'
```

Any row returned means an export ran without a permission it needed. Cross-reference against
`completeness.csv` to see which files were affected.

To confirm the roles on the account being used:

```powershell
Connect-MgGraph -Scopes 'RoleManagement.Read.Directory','Directory.Read.All'
$me = Get-MgUser -UserId (Get-MgContext).Account
Get-MgUserMemberOf -UserId $me.Id |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' } |
    ForEach-Object { $_.AdditionalProperties['displayName'] }
```

---

## References

Australian Signals Directorate. (2023). *Essential Eight maturity model*. Australian Cyber
Security Centre. https://www.cyber.gov.au/resources-business-and-government/essential-cyber-security/essential-eight/essential-eight-maturity-model

Microsoft. (2025). *Least privileged roles by task in Microsoft Entra ID*. Microsoft Learn.
https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/delegate-by-task

Microsoft. (2025). *Microsoft Entra built-in roles*. Microsoft Learn.
https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference

Microsoft. (2025). *Microsoft Graph permissions reference*. Microsoft Learn.
https://learn.microsoft.com/en-us/graph/permissions-reference

Microsoft. (2025). *Use Microsoft Entra groups to manage role assignments*. Microsoft Learn.
https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/groups-concept

> Verify these against current Microsoft documentation before citing in a working paper.
> Role definitions and scope requirements change, and the exports here have been observed to
> shift between releases.
