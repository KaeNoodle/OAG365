# Verification

How to establish that this module and the evidence it produces have not been tampered with.

Version 4.0.0 | OFFICIAL

---

## Two separate questions

The requirement is usually phrased as "prove the script was not tampered with to produce false
results". That is really two questions with two different answers, and conflating them produces
an overstated claim.

| Question | Mechanism | Where |
|---|---|---|
| Is the **code** what OAG signed? | Signed file catalog, checked at startup | `code-integrity.json` |
| Is the **evidence** what the run produced? | SHA-256 manifest, checked by reviewer | `manifest.sha256` |

Neither answers the third question — whether the evidence is a true and complete reflection of
the tenant. Nothing cryptographic can answer that. It is addressed by the scope log, the
completeness check, and reproducibility.

---

## Code integrity

### How it is signed

The module is covered by a single file catalog, `OAG-FileCatalog.cat`, rather than by an
Authenticode signature on each of roughly 60 `.ps1` files.

`New-FileCatalog` records a SHA-256 hash of every file listed in the manifest's `FileList` into
the catalog. Only the catalog is signed. `Test-FileCatalog` then verifies the whole module in one
call, which is what `Test-OagM365Integrity` does at the start of every run.

The reason is practical. Per-file Authenticode would mean 60 signing operations for every change,
however small. That cost creates a real incentive to skip re-signing, and an unsigned module in
practice is worse than a catalogued one.

**The trade-off, which needs to be understood before release:** under an `AllSigned` execution
policy, per-file Authenticode causes PowerShell itself to refuse to run an altered script. A
catalog does not do this. PowerShell will happily run an altered file; it is this module's own
startup check that detects the change and records it.

So the catalog gives **detection and logging**. Per-file signing gives **prevention**. If
prevention at execution time is required, the module must be signed per file and the layout
adjusted accordingly. This should be agreed with the code-signing owner before release rather
than discovered afterwards.

### Regenerating after a change

```powershell
.\Tools\New-OagCatalog.ps1
```

Then sign the catalog:

```powershell
Set-AuthenticodeSignature -FilePath .\OAG-FileCatalog.cat `
    -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0] `
    -TimestampServer 'http://timestamp.sectigo.com'
```

**Always timestamp.** Without a timestamp the signature stops validating the moment the signing
certificate expires. With one, code signed during the certificate's validity period continues to
verify afterwards. The current OAG certificate expires in February 2027; untimestamped signatures
would all fail from that date.

### Verifying manually

```powershell
Test-FileCatalog -CatalogFilePath .\OAG-FileCatalog.cat -Path . -Detailed
Get-AuthenticodeSignature -FilePath .\OAG-FileCatalog.cat | Format-List
```

`Test-FileCatalog` returns `Valid` or `ValidationFailed`. When it fails, the `CatalogItems` and
`PathItems` dictionaries can be compared to identify which specific files differ —
`Test-OagM365Integrity` does this automatically and names them in the run log, because "catalog
invalid" alone is not actionable.

### What the run records

Every run writes `00_RunLog\code-integrity.json`:

```json
{
  "CatalogStatus": "Valid",
  "SignatureStatus": "Valid",
  "SignerSubject": "CN=Office of the Auditor General, ...",
  "SignerThumbprint": "...",
  "CertNotAfter": "2027-02-24",
  "FilesChecked": 62,
  "Verdict": "Verified"
}
```

This ties the evidence to a specific signed code version. A verdict other than `Verified` does
not automatically invalidate the evidence, but it needs explaining before the evidence is relied
on.

Note that `Test-FileCatalog` and `Get-AuthenticodeSignature` are Windows-only. On other platforms
the check records `NotSupportedOnThisPlatform` and the verdict is `Unverified` — the run
continues, but the code integrity claim cannot be made.

---

## Evidence integrity

### The manifest

At the end of every run, `manifest.sha256` is written to the run folder containing a SHA-256 hash
of every file produced, in the standard `sha256sum` format:

```
A3F5...9B2C  ConditionalAccess/capSummary.csv
7E1D...4A88  IAM/users.csv
```

A richer `00_RunLog\manifest.csv` carries the same hashes alongside file size, timestamp, row
count and description.

The manifest is generated **after** the transcript is closed. This ordering is deliberate: a
transcript that is still recording grows as the manifest is written, so its hash would change
immediately after being taken and it would fail verification on every single run. Routine false
failures train reviewers to ignore the check, which defeats its purpose entirely.

### Verifying as a reviewer

```powershell
.\Tools\Verify-Export.ps1 -runFolder C:\Audit\M365\20260823_142530
```

Four outcomes are distinguished, because they mean different things:

| Outcome | Meaning |
|---|---|
| `Matched` | File present, hash agrees |
| `ALTERED` | File present, hash differs — changed after the run |
| `MISSING` | Listed in the manifest, not on disk — deleted |
| `UNLISTED` | On disk, not in the manifest — added after the run |

Exit code is 0 when clean and 1 when verification fails, so it can be used in a pipeline.

Cross-platform alternative:

```bash
cd /path/to/run/folder && sha256sum -c manifest.sha256
```

---

## What this does not establish

This section belongs in the working paper. Overstating the claim is the failure mode to avoid.

**The manifest does not prove the evidence is accurate or complete.** An export run with
insufficient permissions produces an incomplete evidence set whose hashes verify perfectly. The
hash confirms the bytes are unchanged since the run; it says nothing about whether the run saw
everything it should have.

**The manifest does not defend against the operator.** Whoever can alter the exported files can
also regenerate the manifest. This protects against later accidental corruption, casual editing,
and alteration by a third party — not against the person who produced the export. Closing that
gap requires one of:

- signing the manifest with a key the operator does not hold, or
- committing the evidence to a records system promptly, where it cannot subsequently be amended,
  and relying on that system's controls.

The second is usually the practical answer, and it means the manifest's real value is proving
that what is in the records system matches what came out of the tenant on the day.

**Signature validity does not mean the code is correct.** It means the code matches what was
signed. Whether the signed code collects the right things is a review question, not a
cryptographic one.

### What actually supports completeness

Three things, all recorded in `00_RunLog`:

1. **`scopes-granted.csv`** — what the run was permitted to read. This is what distinguishes
   "the tenant has no such policy" from "we could not see it".
2. **`completeness.csv`** — expected outputs against actual, with row counts, flagging empty and
   missing files.
3. **`environment.json`** and **`modules-available.csv`** — which module versions produced the
   evidence, so a second run can be made under matching conditions.

Reproducibility carries more weight than any of these individually. An export repeated with the
same permissions and producing materially the same result is strong evidence; a single
unrepeated run with a valid hash is not.

---

## Suggested working paper wording

> Evidence was collected using OAG-M365-AuditingScript v4.0.0. The module's file catalog and Authenticode
> signature were verified at execution and recorded as [verdict]. A SHA-256 manifest of all
> exported files was generated at completion and re-verified on [date] with no discrepancies,
> confirming the files are unaltered since export.
>
> Completeness was assessed against the run's scope log and completeness record. [n] requested
> permission(s) were not granted, affecting [exports], and this limitation is recorded at
> [reference]. The manifest establishes that the evidence has not been modified since collection;
> it does not independently establish that the collection was complete.

Adjust to the engagement, but keep the final sentence. It is the difference between an accurate
claim and an overstated one.

---

## References

Microsoft. (2025). *About signing*. Microsoft Learn.
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing

Microsoft. (2025). *New-FileCatalog*. Microsoft Learn.
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/new-filecatalog

Microsoft. (2025). *Test-FileCatalog*. Microsoft Learn.
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/test-filecatalog
