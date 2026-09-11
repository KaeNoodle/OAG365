# OAG-M365-AuditingScript

Exports Microsoft 365 and Entra ID configuration as audit evidence.

Version 4.0.0 | OFFICIAL

---

## What it collects

| Report | Contents | Connects via |
|---|---|---|
| `CAP` | Conditional access, authentication method and strength policies | Graph |
| `IAM` | Users, devices, service principals, auth methods, admin units, PIM roles | Graph |
| `ORG` | Tenant config, domains, MX / SPF / DKIM / DMARC | Graph |
| `DFO` | Exchange Online Protection and Defender for Office 365 policies | Exchange Online |
| `TH` | Defender vulnerabilities, agent health, software inventory | Graph |

---

## Before you start

### PowerShell 7.2 or later

```powershell
$PSVersionTable.PSVersion
```

Windows PowerShell 5.1 will not work. Run `pwsh`, not `powershell`.

### Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Scope CurrentUser -Force
```

The pinned version is deliberate. ExchangeOnlineManagement 3.10.1 has a regression affecting the
policy cmdlets the DFO report needs. The module manifest enforces 3.9.x, so the import fails on a
newer version rather than producing confusing errors mid-run.

### Signing — read this before anything else

This module will not run unsigned on an OAG managed workstation. WDAC is enforced, and unsigned
PowerShell is dropped into Constrained Language Mode, where the module cannot work.

The old scripts ran because each carried an Authenticode signature from a certificate the WDAC
policy trusts. That is the only difference. Nothing was bypassing anything.

Execution policy is irrelevant here. `Set-ExecutionPolicy Unrestricted` does not lift Constrained
Language Mode, because WDAC enforces language mode independently.

See `Documentation\Verification.md` for the signing model and what has to be agreed before
release.

### Unblock the files

If copied from a network location, Mark-of-the-Web silently blocks module import. `runMe.ps1`
unblocks the module folder itself on startup, but it can only do that once PowerShell agrees to
run it — so if `runMe.ps1` is itself blocked, clear the folder by hand first:

```powershell
Get-ChildItem -Path .\scripts -Recurse | Unblock-File
```

---

## Running it

Run it with no `-report` parameter to get an interactive menu (Reports / Tools / Help /
Exit) instead of guessing what to type. Pick reports from there, run more than one in the
same session without logging in again, and use Tools to verify a completed run or diagnose
why the module won't start. Passing `-report` skips the menu entirely, for scripted use:

```powershell
# interactive menu
.\runMe.ps1

# everything, non-interactively
.\runMe.ps1 -report All -nonInteractive

# selected reports - tab-completes
.\runMe.ps1 -report CAP,IAM

# choose the evidence folder
.\runMe.ps1 -report ORG -output C:\Audit\ClientName\2026

# skip the export that needs Global Administrator
.\runMe.ps1 -report ORG -skipOnPremSync
```

### Running one report on its own

Every report writer can be called directly. No context or connection has to be passed in — each
one calls `runEnsure`, which starts a run if none exists:

```powershell
Import-Module .\OAG-ModuleManifest.psd1
capReportWrite
```

Useful when re-running a single failed report without repeating the whole set.

### Authenticating as a registered application

```powershell
# certificate - preferred, no secret to store or rotate
.\runMe.ps1 -appClientId "d359..." -appTenantId "a1b2..." -appCertThumbprint "A1B2..."

# secret
.\runMe.ps1 -appClientId "d359..." -appTenantId "a1b2..." -appSecret "a1bC2d~..."
```

---

## What you get

One run produces one folder:

```
<output>\results\20260824_142530\
├── 00_RunLog\
│   ├── transcript.txt          full console output
│   ├── environment.json        PowerShell version, module version, language mode
│   ├── modules-available.csv   every Graph / EXO module version installed
│   ├── code-integrity.json     catalog and signature check result
│   ├── scopes-granted.csv      requested vs granted - read this one
│   ├── events.csv              every warning and error raised
│   ├── completeness.csv        per-file row counts and status
│   ├── run-summary.json        overall result
│   └── manifest.csv            hashes with row counts
├── ConditionalAccess\   IAM\   Organisation\   ThreatProtection\   Defender\
└── manifest.sha256             SHA-256 of every file above
```

Previously each of the five scripts generated its own timestamp, so running all five produced five
sibling folders that could not be shown to belong to one engagement.

### Reading the output

Check three things before relying on the evidence:

1. `run-summary.json` — the `completenessStatus` field
2. `scopes-granted.csv` — any row with `granted = False` means an export ran without a permission
   it needed
3. `completeness.csv` — any row marked `Empty` or `Missing`

An empty export is not necessarily wrong. A tenant may genuinely have no named locations. The
point is that empty is recorded and visible, rather than indistinguishable from a permission
failure.

---

## Verifying evidence

```powershell
.\Tools\Verify-Export.ps1 -runFolder C:\Audit\results\20260824_142530
```

See `Documentation\Verification.md` for what this does and does not establish — the distinction
matters for how the evidence is described in a working paper.

---

## Troubleshooting

**"Cannot run in Constrained Language Mode"**
The module is not signed with a certificate the WDAC policy trusts. This is the expected failure
on a managed workstation until signing is arranged. See `Documentation\Verification.md`.

**"Module was not loaded because no valid module file was found"**
Wrong folder, or Mark-of-the-Web. Run the `Unblock-File` command above.

**"The required module ExchangeOnlineManagement is not installed"**
The manifest pins 3.9.x, so installing the latest will not satisfy it:
`Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Force -AllowClobber`

**DFO produces nothing / "term is not recognised"**
This was the temporary cmdlet module scope bug, fixed in 4.0.0. Check the transcript for
"EXO cmdlet availability verified" — its absence means `tmpEXO_*` did not import globally.

**TH returns permission denied**
`ThreatHunting.Read.All` needs tenant administrator consent. Record the gap rather than working
around it.

**ORG returns 403 on on-premises sync**
Expected under read-only credentials. Use `-skipOnPremSync` and document the limitation.

**A CSV is empty with no error**
Check `scopes-granted.csv`, then `events.csv`. Version 4.0.0 logs zero-row exports as warnings.

---

## Layout

```
OAG365\                          repository root
├── scripts\                     the module - entry point, functions, tools
│   ├── runMe.ps1                entry point
│   ├── OAG-ModuleManifest.psd1  versions, dependencies, catalog FileList
│   ├── OAG-FileInfo.psm1        loader
│   ├── OAG-FileCatalog.cat      generated by Tools\New-OagCatalog.ps1, then signed
│   ├── ReportWriter\            one file per report
│   ├── Functions\               shared functions, plus CAP\ IAM\ ORG\ DFO\ TH\
│   ├── Tools\                   catalog generation, evidence verification
│   └── Documentation\
└── results\                     evidence output - see "What you get" below
```

`results\` is deliberately not inside `scripts\`. A code-catalog rebuild
(`Tools\New-OagCatalog.ps1`) hashes everything under its target folder, so evidence living
inside the module folder risks getting baked into the signed code catalog. `-output`
defaults to the folder next to `scripts\` for this reason; pass `-output` to send evidence
somewhere else entirely (e.g. `-output C:\Audit\ClientName\2026`).

### Conventions

Functions are named `subjectAction` in camelCase — `msGraphConnect`, `exportWrite`, `guidVerify`,
`capPolicyToHtml`. Files group related functions by purpose.

Every function carries a banner header with the same five sections:

```
DESCRIPTION                what it does and why
LOGIC                      how it works, following the code top to bottom
PARAMETERS                 what it takes
RUNNING CONTEXT            what calls it, what it calls, what it returns
CMLETS/PERMISSIONS/SCOPES  the Graph or EXO cmdlets and the access they need
```

### No context parameters

Run state lives in `$script:run` at module scope, and export paths in `$script:exportTarget`.
Functions read them directly rather than having a context object threaded through every call.

That is why a single report can be run on its own with no arguments. It also means the ported
export functions kept working unchanged, since `$script:` inside a module is module scope shared
by every function in it.

Collections are plain arrays rather than generic lists. Generic list constructors are blocked in
Constrained Language Mode, and at this scale the speed difference is irrelevant.

### After changing any file

```powershell
.\Tools\New-OagCatalog.ps1
```

Then have the resulting `.cat` signed. See `Documentation\Verification.md`.
