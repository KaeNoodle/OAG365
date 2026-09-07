# Changelog

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

### Known gaps carried forward

- `OnPremDirectorySynchronization.Read.All` requires Global Administrator. Not resolvable by any
  read-only role. See `Docs\Permissions.md`.
- `ThreatHunting.Read.All` requires tenant administrator consent, not yet granted.
- The evidence manifest is not itself signed. See `Docs\Verification.md` for what this means.
