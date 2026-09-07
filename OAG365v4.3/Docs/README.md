# OagExportM365

Exports Microsoft 365 and Entra ID configuration as audit evidence.

Version 4.0.0 | OFFICIAL

---

## What this does

Collects configuration evidence from an M365 tenant across five areas, writes it to a single
timestamped folder, and produces a SHA-256 manifest so the evidence can be shown to be unaltered
afterwards.

| Report | Collects | Connects via |
|---|---|---|
| `Cap` | Conditional access, authentication method and strength policies | Graph |
| `Iam` | Users, devices, service principals, auth methods, admin units, PIM roles | Graph |
| `Org` | Tenant config, domains, MX / SPF / DKIM / DMARC | Graph |
| `Dfo` | Exchange Online Protection and Defender for Office 365 policies | Exchange Online |
| `ThreatHunt` | Defender vulnerabilities, agent health, software inventory | Graph |

---

## Before you start

### 1. PowerShell 7.2 or later

```powershell
$PSVersionTable.PSVersion
```

Windows PowerShell 5.1 will not work. If `PSVersion` shows 5.x, install PowerShell 7 from
<https://aka.ms/powershell> and run `pwsh`, not `powershell`.

### 2. Required modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Scope CurrentUser -Force
```

Note the pinned version. ExchangeOnlineManagement 3.10.1 has a regression affecting the policy
cmdlets the `Dfo` report depends on; 3.9.0 is the last build confirmed working. The module
manifest enforces this and the import will fail on a newer version rather than producing
confusing errors partway through a run.

### 3. Permissions

See `Docs\Permissions.md` for the full mapping. In short, Global Reader plus Security Reader
covers four of the five reports. Two exports cannot be obtained with read-only credentials at
all and are documented there as known gaps.

### 4. Unblock the files

If the module was downloaded or copied from a network location, Windows marks the files with
Mark-of-the-Web, which silently blocks module import. This produces a confusing failure because
nothing obviously errors — the module just refuses to load.

```powershell
Get-ChildItem -Path .\OagExportM365 -Recurse | Unblock-File
```

Run this before anything else if the module will not import.

---

## Running it

The simplest case, which prompts for interactive login and runs everything:

```powershell
.\Export-M365.ps1
```

Selected reports only. This is the normal case — you rarely need all five at once, and the
reports have different permission requirements:

```powershell
.\Export-M365.ps1 -Report Cap,Iam
```

`-Report` tab-completes, so you do not need to remember the names.

Choose where evidence goes:

```powershell
.\Export-M365.ps1 -Report Org -Output C:\Audit\ClientName\2026
```

Skip the on-premises sync export, which requires Global Administrator and returns 403 under
read-only credentials:

```powershell
.\Export-M365.ps1 -Report Org -SkipOnPremSync
```

Unattended, with no confirmation pauses:

```powershell
.\Export-M365.ps1 -NonInteractive
```

### Authenticating as a registered application

Where interactive login is not appropriate, authenticate as an Entra enterprise application
using either a certificate or a secret. The application must exist and have the required
scopes consented before running.

```powershell
# Certificate (preferred - no secret to store or rotate)
.\Export-M365.ps1 -AppClientId "d3590ed6-..." -AppTenantId "a1b2c3d4-..." -AppCertThumbprint "A1B2C3..."

# Secret
.\Export-M365.ps1 -AppClientId "d3590ed6-..." -AppTenantId "a1b2c3d4-..." -AppSecret "a1bC2d~..."
```

Create a signing-capable certificate in an elevated session:

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=OagExportM365" `
          -CertStoreLocation "Cert:\LocalMachine\My" -KeyExportPolicy Exportable `
          -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
          -NotAfter (Get-Date).AddMonths(3)

Export-Certificate -Cert $cert -FilePath ".\OagExportScript.cer"
```

Upload the `.cer` to the registered application in Entra.

---

## What you get

One run produces one folder:

```
<Output>\M365\20260823_142530\
├── 00_RunLog\
│   ├── transcript.txt          full console output
│   ├── environment.json        PowerShell version, module version, host, language mode
│   ├── modules-available.csv   every Graph / EXO module version installed
│   ├── code-integrity.json     catalog and signature check result
│   ├── scopes-granted.csv      requested vs granted - read this one
│   ├── events.csv              every warning and error raised
│   ├── completeness.csv        per-file row counts and status
│   ├── run-summary.json        overall result
│   └── manifest.csv            hashes with row counts and descriptions
├── ConditionalAccess\
├── IAM\
├── Organisation\
├── ThreatProtection\
├── Defender\
└── manifest.sha256             SHA-256 of every file above
```

Previously each of the five scripts generated its own timestamp, so running all five produced
five sibling folders that could not be shown to belong to one engagement. One run now means one
folder.

### Reading the output

Check three things before relying on the evidence:

1. **`run-summary.json`** — the `CompletenessStatus` field. `Complete`, `Complete with warnings`
   or `Incomplete`.
2. **`scopes-granted.csv`** — any row with `Granted = False` means an export ran without the
   permission it needed. The resulting file may be empty or partial.
3. **`completeness.csv`** — any row with status `Empty` or `Missing`.

An export that returns no records is not necessarily wrong. A tenant may genuinely have no named
locations configured. The point is that an empty result is recorded and visible rather than
silently indistinguishable from a permission failure.

---

## Verifying evidence

```powershell
.\Tools\Verify-Export.ps1 -RunFolder C:\Audit\M365\20260823_142530
```

Recomputes every hash and reports files that were altered, deleted, or added after the run. See
`Docs\Verification.md` for what this does and does not establish — the distinction matters for
how the evidence is described in a working paper.

---

## Troubleshooting

**"Module was not loaded because no valid module file was found"**
Wrong folder, or Mark-of-the-Web. Run the `Unblock-File` command above.

**"The required module ExchangeOnlineManagement ... is not installed"**
The manifest pins 3.9.x. Installing the latest version will not satisfy it:
`Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Force -AllowClobber`

**"PowerShell is running in ConstrainedLanguage, not FullLanguage"**
Application control is restricting the session. This cannot be worked around in code — see
`Docs\ConstrainedLanguageMode.md`. Diagnose the cause first:

```powershell
.\Tools\Test-LanguageMode.ps1
```

**Dfo report produces nothing / "term is not recognised"**
This was the temporary cmdlet module scope bug and is fixed in 4.0.0. If it recurs, check the
transcript for "EXO cmdlet availability verified" — its absence means the `tmpEXO_*` module did
not import at global scope.

**ThreatHunt returns permission denied**
`ThreatHunting.Read.All` requires tenant administrator consent. Until it is granted this report
cannot run. Record the gap rather than working around it.

**Org report 403 on on-premises sync**
Expected under read-only credentials. That single export requires Global Administrator. Use
`-SkipOnPremSync` and document the limitation.

**A CSV is empty but no error appeared**
Check `scopes-granted.csv` first, then `events.csv`. Version 4.0.0 logs zero-row exports as
warnings specifically so this is visible.

---

## For developers

```
OagExportM365\
├── Export-M365.ps1          entry point and dispatcher
├── OagExportM365.psd1       manifest: versions, dependencies, catalog FileList
├── OagExportM365.psm1       loader
├── Public\                  5 report orchestrators, one file per function
├── Private\
│   ├── Common\              connect, disconnect, logging, run context, integrity, manifest
│   ├── Cap\ Iam\ Org\ Dfo\  report-specific functions ported from the original scripts
│   └── ThreatHunt\          query definitions and executor
├── Tools\                   catalog generation, evidence verification
└── Docs\
```

Filenames match function names exactly. Functions follow `Verb-Noun` with a single hyphen and a
singular noun, so `Get-Verb` and PSScriptAnalyzer are satisfied.

After changing any file, regenerate and re-sign the catalog:

```powershell
.\Tools\New-OagCatalog.ps1
```

Then have the resulting `.cat` signed. See `Docs\Verification.md`.
