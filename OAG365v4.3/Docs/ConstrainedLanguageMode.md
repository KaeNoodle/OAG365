# Constrained Language Mode

What to do when the export will not run because PowerShell is constrained.

Version 4.0.0 | OFFICIAL

---

## Symptom

```
PowerShell is running in ConstrainedLanguage, not FullLanguage.
```

Or, if running the functions directly rather than through the entry point, errors such as:

```
Cannot convert value to type "System.Management.Automation.LanguagePrimitives+InternalPSCustomObject".
Only core types are supported in this language mode.
```

Confirm with:

```powershell
$ExecutionContext.SessionState.LanguageMode
```

---

## Why this cannot be coded around

Constrained Language Mode (CLM) restricts PowerShell to a small allow-list of types. Testing
against a genuinely constrained session confirms the following are blocked:

| Capability | Available in CLM? |
|---|---|
| `[PSCustomObject]@{}` | **No** |
| `[System.Collections.Generic.List[object]]` | **No** |
| `[System.Collections.Generic.HashSet[string]]` | **No** |
| `[math]::Round()` and other .NET static calls | **No** |
| Hashtable literals | Yes |
| Method calls on core types such as `DateTime` | Yes |

The first item is decisive. `[PSCustomObject]` is used for the run context, the export target
paths, every completeness record and effectively every row of exported data. It is not an
incidental convenience that could be swapped out.

Even if this module were rewritten to avoid all of the above — a substantial rewrite producing
markedly worse code — the Microsoft Graph SDK has the same dependencies. The restriction reaches
past our own code into the modules we depend on, so there is no version of this tool that runs
under CLM.

This is the correct behaviour on the part of application control. The module stops rather than
continuing, because a partial evidence set that does not announce itself as partial is worse
than no evidence at all.

---

## Finding the cause

```powershell
.\Tools\Test-LanguageMode.ps1
```

This is written to run inside CLM. It checks the four mechanisms that impose it and reports
which is active:

1. **`__PSLockdownPolicy` environment variable** — a legacy and testing mechanism. Occasionally
   set by accident or left behind after a test. The simplest cause to resolve, but do not remove
   it without asking; it may be deliberate.
2. **WDAC / Device Guard** — the most common cause on a managed SOE. A script that does not
   satisfy the code integrity policy runs in CLM rather than being blocked outright, which is why
   the failure looks like a language problem rather than an application control one.
3. **AppLocker script rules** — in `Enabled` enforcement, scripts that match no Allow rule run in
   CLM. In `AuditOnly` they do not.
4. **Third-party application control** — the ISAG environment uses Airlock Digital, which is the
   most likely cause here.

Run it in the same shell, on the same machine, where the export fails. Language mode can differ
between `powershell.exe` and `pwsh.exe`, between 32-bit and 64-bit hosts, and between local and
network paths.

---

## Resolving it

All four options need the application control administrator. Pick based on what the diagnostic
found.

### Option 1 — allowlist by publisher (preferred)

Have the module signed by OAG, then have the signing certificate trusted by the application
control policy. Signed code satisfies WDAC and Airlock publisher rules and runs in FullLanguage.

This is the option worth pushing for, because it does double duty: the signing needed to make
the module run is the same signing needed for the evidence integrity requirement in
`Docs\Verification.md`. One request, two problems solved. It also survives file moves and
version updates, so it does not need redoing each time the module changes.

Sequence:

```powershell
.\Tools\New-OagCatalog.ps1
# then have the .cat signed, and the certificate added to the application control policy
```

### Option 2 — allowlist by path

Request that a specific tooling directory be trusted, and run the module from there. For example
a controlled path such as `C:\OAG\Tools\` that ordinary users cannot write to.

Weaker than publisher trust, because anything placed in that directory inherits the trust. It is
usually faster to obtain, so it is a reasonable interim measure while publisher trust is arranged.

### Option 3 — temporary exemption

Airlock supports one-time-password exemptions that lift enforcement for a defined window. Suitable
for a single audit run where the tooling is not going to be used repeatedly.

Record the exemption reference in the working paper. An auditor who disabled a control to collect
evidence needs that visible, not buried.

### Option 4 — a different machine

If the audit permits it, run from a management workstation or jump host that is not subject to the
same policy. Confirm first that this is acceptable for the engagement: the machine still needs to
be trusted enough to hold tenant credentials and evidence.

---

## What not to do

Do not attempt to bypass the restriction. Downgrading to PowerShell 2, using a language-mode
escape, or copying the code into a different host to evade policy are all techniques used to
defeat application control, and using them would be difficult to defend regardless of intent —
particularly for an auditor whose own work includes assessing that control.

If the constraint cannot be resolved within the fieldwork window, record it as a scope limitation.
That is an accurate and defensible outcome.

---

## Suggested wording for a working paper

> Automated collection of M365 configuration evidence was initially prevented by application
> control on the audit workstation, which restricted PowerShell to Constrained Language Mode.
> This was resolved by [signing the tooling and adding the certificate to the application control
> policy / obtaining a path exemption for the tooling directory / a time-limited exemption
> reference NNNN] on [date], approved by [name]. Evidence collection proceeded on [date].

If it was not resolved:

> Automated collection could not be performed because application control on the audit workstation
> restricts PowerShell to Constrained Language Mode, and an exemption was not obtained within the
> fieldwork period. Configuration evidence for [areas] was instead obtained by [alternative
> method], and the reduced coverage is noted at [reference].

---

## References

Australian Signals Directorate. (2023). *Essential Eight maturity model*. Australian Cyber
Security Centre. https://www.cyber.gov.au/resources-business-and-government/essential-cyber-security/essential-eight/essential-eight-maturity-model

Australian Signals Directorate. (2024). *Implementing application control*. Australian Cyber
Security Centre. https://www.cyber.gov.au/resources-business-and-government/maintaining-devices-and-systems/system-hardening-and-administration/application-control/implementing-application-control

Microsoft. (2025). *About language modes*. Microsoft Learn.
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_language_modes

Microsoft. (2025). *Application Control for Windows*. Microsoft Learn.
https://learn.microsoft.com/en-us/windows/security/application-security/application-control/app-control-for-business/appcontrol
