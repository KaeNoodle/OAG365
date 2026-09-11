# Changelog

## Unreleased — failures observed on a live test tenant

A full run against a test tenant produced three stack traces and one misleading empty
file, none of which were faults in the tenant. This round makes each of them either fix
itself or explain itself.

### Permissions

- Added `RoleManagementPolicy.Read.AzureADGroup` to the IAM scope set. Without it every
  PIM group policy lookup returned `403 PermissionScopeNotGranted`, so the run captured
  who was in each PIM-managed group but no evidence of the activation controls over
  them - approval, MFA on activation, maximum duration, justification. The group
  membership calls kept working throughout, so the gap was silent in the CSV output and
  visible only as a stack trace in the transcript. This scope prompts for consent on the
  next interactive sign-in.

### Failure reporting

- PIM policy lookups now distinguish a missing scope from a missing licence from a real
  fault, and report a missing scope once per run rather than once per role and once per
  group. A four-group tenant previously produced four identical multi-screen 403 dumps.
- Defender advanced hunting now separates `403` (consent not granted, operator can fix)
  from `401` with the scope already granted, which is what the endpoint returns when the
  tenant has no Microsoft Defender XDR workload onboarded. The second is not a permission
  problem and recording it as one puts a remediation in the working paper that would
  never have worked.
- When the hunting endpoint refuses the first query it refuses all of them, so the
  remaining queries are now skipped with the reason recorded instead of repeating the
  same failure three times.
- `completeness.csv` gained a `reason` column and a `Not applicable` status for evidence
  that is absent because the capability does not exist in the tenant. An absent file for
  a control that cannot exist is not a gap, and it no longer reads as one. `Missing`
  still means what it always meant.

### Connection resilience

- The DFO report now runs in its own PowerShell process when the run has already signed
  in to Graph. `ExchangeOnlineManagement` and the `Microsoft.Graph` modules each carry
  their own copy of MSAL and cannot both authenticate in one process: whichever goes
  second fails, for Exchange Online as a `NullReferenceException` thrown while its
  sign-in broker is built. That is why a full run always lost the DFO report while
  `-report DFO` on its own worked. No sign-in method avoids it, so the fix is a clean
  process rather than a different credential.

  The child process joins the run already under way (`-joinRunId`), writes its CSVs into
  the same folder, and hands back what it produced in `child-DFO.json`. The parent folds
  that into the completeness check, so one run still produces one folder, one summary and
  one manifest. The operator signs in twice, once per service.
- `Connect-ExchangeOnline` also falls back to a broker-free sign-in and then to device
  code before giving up, and the failure message now names the two causes that look
  generic: the MSAL conflict above, and sign-in error 530035, which means security
  defaults are enabled and block device code.
- `runInitialize` accepts `-runId` and `-transcriptName` so a second process can write
  into an existing run folder without fighting the first for the transcript file.
- `Unblock-File` is now guarded by a platform check rather than by whether the cmdlet
  exists. Off Windows it exists and throws, which stopped the script before it started.

### Output clarity

- A tenant with no conditional access policies now says so, and says that the empty
  summary file is the evidence of it. Previously the run printed "Processing policy..."
  and wrote an empty CSV, which reads like a broken export rather than a finding.
- MX lookups print one line per domain with a record count, not one line per MX record.
  A single domain routed to Exchange Online was producing nine identical lines.

## 4.1.0 — Renamed entry point, interactive menu, broader Debugger

### Repository layout

- Top-level module folder renamed from `OAG365v1` to `scripts`.
- `-output` now defaults to the folder next to `scripts\` (a `results\` folder at the
  repository root) instead of inside it, and the run-output subfolder itself is renamed
  from `M365\<runId>\` to `results\<runId>\`. Same reasoning for the solo-report path in
  `runEnsure.ps1`, which now passes a separate `moduleRoot`/`outputRoot`.
  This is deliberate: `Tools\New-OagCatalog.ps1` hashes everything under its target folder
  when it builds the code-integrity catalog, so evidence written inside the module folder
  risks being baked into the signed catalog. Once baked in, moving or deleting that
  evidence later shows up as a spurious `CATALOG MISMATCH` on every future run, even though
  no code changed. Passing `-output` explicitly still works exactly as before.

### Entry point

- `OAG-MainRunFile.ps1` renamed to `runMe.ps1`. Every reference across the manifest,
  `codeVerify.ps1`'s fallback signature check, `Tools\Build-Module.ps1`, doc-comments and
  `Documentation\README.md` updated to match.
- Added a looping interactive menu (Reports / Tools / Help / Exit) for when the operator
  runs `runMe.ps1` with no `-report` parameter. Scripted and unattended invocations
  (`-report ...`, `-nonInteractive`, the app-auth parameter sets) are unaffected — the menu
  only appears when `-report` was not explicitly passed and `-nonInteractive` is not set.
- Within one interactive session, picking further reports after the first reuses the
  existing Microsoft Graph / Exchange Online connection instead of reconnecting, and all
  reports picked in the session land in the same run folder and manifest. The first Graph
  connection of an interactive session requests the scopes for every Graph report up
  front, so a later pick is never left under-scoped.
- The Reports menu asks "Get full report" (all 5) or "Get partial report" (choose which)
  rather than listing every report and an "All" option side by side.
- The Tools menu exposes the renamed Debugger and `Verify-Export.ps1` (the latter
  defaulting to the most recent completed run) — not the release-engineering scripts
  (`Build-Module.ps1`, `New-OagCatalog.ps1`), which stay out of the operator-facing menu.
  Help now points to the Debugger when the module won't run at all.

### Sign-in and verification usability

- Tools > Verify-Export now lists the runs it finds under the output folder and takes a
  number, instead of asking the operator to type a full path including the run's timestamp.
  Runs with no `manifest.sha256` are listed but marked as unfinished, since Verify-Export
  cannot check those and would otherwise just error after being chosen. A path can still be
  entered by hand for a run held somewhere else.
- Added `-deviceCode`, which signs in to Graph with a device code instead of the interactive
  window. Web Account Manager sign-in runs on a short timer and opens a window that, on an
  embedded terminal, frequently appears behind everything else — so the common failure is the
  operator never seeing the prompt rather than anything being wrong with the account. The
  interactive path now also warns that the window may open behind, and retries once with a
  device code by itself if the attempt fails.
- Corrected the README's ExchangeOnlineManagement troubleshooting entry, which claimed
  installing the latest version would not satisfy the manifest. `RequiredModules` states a
  minimum, so anything newer does satisfy it.

### Cleanup

- `Tools\Build-Module.ps1` worked against a source layout that no longer exists. It read
  `Private\` and `Public\`, which threw immediately under `$ErrorActionPreference = 'Stop'`,
  so the build was dead. It now reads `Functions\` and `ReportWriter\`, takes its
  `Export-ModuleMember` list from the manifest's `FunctionsToExport` rather than from
  filenames (a filename is not a function name here — `ConditionalAccessPolicy.ps1` defines
  `capReportWrite`), and validates by counting function definitions rather than files, which
  previously guaranteed a count mismatch.
- `Tools\Debugger.ps1` defaulted `-ModulePath` to `$PSScriptRoot`, which resolves to `Tools\`
  rather than the module root, so a standalone run checked the wrong folder. Now defaults to
  the parent.
- Removed two committed editor backup files (`OAG-ModuleManifest.psd1~`,
  `graphObjectResolve.ps1~`) that shadowed real files with stale copies, and added a
  `.gitignore` covering `results/`, `dist/` and `*~`.
- Corrected stale references throughout: `Docs\` → `Documentation\`, `Test-OagM365Integrity`
  → `codeVerify`, `Get-OagM365IamPimRole` → `iamPimRoleGet`, and the README's unblock command,
  which still named the old module folder.
- `orgEmailExport` built its DMARC result with `+=` against an uninitialised variable where
  the surrounding MX and SPF lines used `=`.

### Diagnostics

- `Tools\Test-LanguageMode.ps1` renamed to `Tools\Debugger.ps1`. All existing Constrained
  Language Mode checks are unchanged.
- Added an execution-policy check that flags `AllSigned`/`Restricted` specifically, since
  this module is signed via a file catalog rather than per-file Authenticode and those
  policies refuse to run it outright.
- Added a required-PowerShell-version and required-modules check covering PowerShell 7
  and every module the five reports depend on, printing the install command for anything
  missing. Checks presence only, not a specific pinned version.

## 4.0.0 — Restructure from five scripts into one module

### Structure

- Five standalone scripts consolidated into a module with one entry point,
  `Export-M365.ps1 -Report Cap,Iam,Org,Dfo,ThreatHunt,All`.
- Reports remain independently runnable. This was deliberate in the original design and is
  preserved, because the reports have genuinely different permission requirements.
- Approximately 1,200 lines of duplicated code removed. `Connect-Oag365-ToGraph` was
  byte-for-byte identical across four files; `Format-Oag365-Exception` across all five.
- 43 report-specific functions moved to one function per file, filenames matching function
  names.
- Functions renamed to `Verb-Noun` with a single hyphen and singular noun, satisfying
  `Get-Verb` and PSScriptAnalyzer. The eight `Resolve-Oag365Cap-*` helpers lost the `Cap`
  prefix and moved to `Common`, since the IAM report needs the same GUID resolution.

### Output

- One run ID per execution. Previously each script called `Get-Date` independently, so running
  all five produced five sibling folders with different timestamps and nothing linking them.
- Transcript moved inside the run folder rather than beside it.
- Added `00_RunLog` capturing environment, module versions, granted scopes, events and
  completeness.

### Integrity

- Added `manifest.sha256` covering every file produced.
- Added code integrity self-check at startup, recording catalog and signature status.
- Added `Tools\Verify-Export.ps1` for reviewer-side verification, distinguishing altered,
  missing and unlisted files.
- Added `Tools\New-OagCatalog.ps1` for catalog generation.

### Bugs fixed

**Exchange Online temporary cmdlet module scope.** `Connect-ExchangeOnline` generates a
`tmpEXO_*` module at connection time and imports it into the calling scope. Because the connect
call sat inside a function, the module was discarded when that function returned and every
subsequent EXO cmdlet failed. The DFO report produced no output as a result. The generated
module is now re-imported with `-Global`, and cmdlet availability is verified before the exports
run rather than assumed.

**Connection test parsed as arguments.** The DFO report used
`if (Connect-Oag365Dfo-ToExchangeOnline -eq $true)`. PowerShell parses that as passing two
positional arguments to the function, not as comparing its return value, so the branch never
tested what it appeared to. The result is now assigned and tested properly.

**Silent error suppression in threat hunting.** All three query blocks assigned the formatted
exception to a variable in their `catch` and never wrote it out. A failed query produced no
console output, no transcript entry and no file, so the run appeared to succeed while producing
nothing. Errors are now reported, and permission failures are identified specifically.

**Non-terminating Graph errors bypassing try/catch.** `-ErrorAction Stop` added to Graph calls
so API errors are actually catchable.

**Event log silently empty.** `$Context.Events -ne $null` placed the collection on the left of
the comparison, which PowerShell treats as a filter rather than a boolean test. An empty list
filtered against `$null` returns empty, evaluating as false, so the first event was never added
and the log stayed empty for the entire run — discarding every warning and error. `$null` is now
on the left.

**Manifest hashing an open transcript.** The transcript was still recording when it was hashed,
so its hash changed immediately afterwards and it failed verification on every run. The
transcript is now closed before the manifest is generated.

**Module version lookup throwing when not imported.** `(Get-Module ...).Version.ToString()`
threw a null reference when the functions were dot-sourced rather than imported. Now null-safe.

### Behaviour changes

- Graph scopes are declared per report and passed as parameters, rather than read from a
  script-scope variable that had drifted from the documented list.
- `RequiredModules` in the manifest enforces module versions at import. The original printed a
  warning and continued. ExchangeOnlineManagement is pinned to 3.9.x because 3.10.1 has a known
  regression.
- Constrained Language Mode now aborts. The original printed "Abort now" and carried on.
- Zero-row exports are logged as warnings so an empty result is visible rather than
  indistinguishable from a permission failure.
- Each DFO policy export is isolated, so one failing policy type does not prevent the rest being
  collected.
- The Org report's on-premises sync export is isolated and can be skipped with
  `-SkipOnPremSync`, since it requires Global Administrator and returns 403 under read-only
  credentials.


### Naming and layout (4.0.0 restructure, second pass)

- Functions renamed to camelCase `subjectAction` throughout: `msGraphConnect`,
  `exceptionFormat`, `runInitialize`, `guidVerify`, `exportWrite`, `exportVerify`,
  `capPolicyToHtml` and so on.
- Comment-based help replaced with banner headers carrying five fixed sections:
  DESCRIPTION, LOGIC, PARAMETERS, RUNNING CONTEXT, CMLETS/PERMISSIONS/SCOPES. LOGIC follows
  the code top to bottom; RUNNING CONTEXT names callers and callees so the call graph is
  readable without searching.
- Context parameters removed. Run state lives in `$script:run` and export paths in
  `$script:exportTarget`, both at module scope. Functions read them directly.
- Each report writer calls `runEnsure`, which starts a run if none is active. That is what
  lets a single report be called on its own with no arguments.
- Generic lists replaced with plain arrays, since generic list constructors are blocked in
  Constrained Language Mode.
- 43 ported functions regrouped from one-per-file into 19 files grouped by purpose.

### Additional bug fixed

**Missing `ExportM365-GroupMembers`.** `capPolicyToHtml` called this function in two places,
for included and excluded groups, but it was never defined anywhere in the original file.
PowerShell treats an unresolved command as a non-terminating error, so the surrounding
`ForEach-Object` carried on, the member collection stayed empty, and the HTML report rendered
each group name with an empty member list. No error appeared and the report looked complete.

A reader would reasonably conclude the group had no members when in fact they were never
retrieved. Implemented as `capGroupMemberGet`, which validates the group ID, pages through
`Get-MgGroupMember`, and logs failures rather than swallowing them.

Note that `Get-MgGroupMember` returns direct members only. Nested groups are not expanded, so
a group containing another group lists that group rather than its users. This matches how Entra
evaluates conditional access assignment, but it means the member list is not a full effective
user list, and that should be stated in any working paper relying on these counts.

### Environment constraint identified

The module cannot run unsigned on an OAG managed workstation. WDAC is enforced
(`CiTool --list-policies` for the active policy set), and unsigned PowerShell is dropped into
Constrained Language Mode. The original scripts ran because each carried an Authenticode
signature from a certificate the WDAC policy trusts.

Execution policy is irrelevant: `Set-ExecutionPolicy Unrestricted` does not lift Constrained
Language Mode.

This makes per-file Authenticode, rather than catalog signing alone, the likely requirement.
WDAC can honour catalogs, but only where the `.cat` is registered into each machine's catalog
store, which is a per-machine deployment step. See `Documentation\Verification.md`.

### Known gaps carried forward

- `OnPremDirectorySynchronization.Read.All` requires Global Administrator. Not resolvable by any
  read-only role. See `Documentation\Permissions.md`.
- `ThreatHunting.Read.All` requires tenant administrator consent, not yet granted.
- The evidence manifest is not itself signed. See `Documentation\Verification.md` for what this means.
- The module is unsigned and therefore cannot currently run on a managed workstation. Signing must be arranged before testing against a live tenant.
- The 43 ported functions have been verified to parse and resolve, but have not been run against a live tenant.
