<#
.SYNOPSIS
Diagnoses why the OAG365 module won't run: Constrained Language Mode, execution policy,
or a missing PowerShell version or module.

.DESCRIPTION
Three unrelated things can each stop this module dead, and the fix for each is different, so
the first job is identifying which one (or ones) is active:

  - Constrained Language Mode (CLM), imposed by application control, not by PowerShell itself.
    Several different mechanisms can cause it - see the numbered CAUSE sections below.
  - Execution policy. This module is signed via a file catalog rather than per-file
    Authenticode, which an AllSigned or Restricted policy will refuse to run outright.
  - PowerShell 7 or a required module not installed. Checks presence only, not a specific
    pinned version - the manifest's own RequiredModules only enforces a minimum.

This script is deliberately written to run WITHIN Constrained Language Mode. It uses only
cmdlets, hashtables, arrays and [PSCustomObject], and avoids generic collections, New-Object,
::new(), Add-Type and .NET static method calls, all of which CLM blocks. If it fails to run,
that failure is itself diagnostic and worth reporting.

.EXAMPLE
.\Debugger.ps1

.EXAMPLE
.\Debugger.ps1 -ModulePath C:\Tools\OAG-M365-AuditingScript

.NOTES
VERSION: 2.0
Run on the machine and in the shell where the export fails.
#>
[CmdletBinding()]
param (
    [string]$ModulePath = $PSScriptRoot
)

Write-Host ""
Write-Host "PowerShell Language Mode diagnostic" -ForegroundColor Cyan
Write-Host "===================================================================="
Write-Host ""

# ---------------------------------------------------------------------------------------
# 1. Current state
# ---------------------------------------------------------------------------------------
$mode = $ExecutionContext.SessionState.LanguageMode

Write-Host "CURRENT SESSION" -ForegroundColor Cyan
Write-Host "  Language mode     : $mode" -ForegroundColor $(if ($mode -eq 'FullLanguage') { 'Green' } else { 'Red' })
Write-Host "  PowerShell        : $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
Write-Host "  Host              : $($Host.Name)"
Write-Host "  Executable        : $([Environment]::ProcessPath)" -ErrorAction SilentlyContinue
Write-Host "  Process is 64-bit : $([Environment]::Is64BitProcess)" -ErrorAction SilentlyContinue
Write-Host "  User              : $env:USERDOMAIN\$env:USERNAME"
Write-Host "  Computer          : $env:COMPUTERNAME"
Write-Host ""

if ($mode -eq 'FullLanguage') {
    Write-Host "This session is NOT constrained. If the export still fails, the problem is" -ForegroundColor Green
    Write-Host "elsewhere - check Mark-of-the-Web (Unblock-File) and module versions." -ForegroundColor Green
    Write-Host ""
}

# ---------------------------------------------------------------------------------------
# 2. Cause 1 - the __PSLockdownPolicy environment variable
#    A legacy and testing mechanism. Easy to check, easy to fix, and occasionally set by
#    accident or left behind by a previous test.
# ---------------------------------------------------------------------------------------
Write-Host "CAUSE 1: __PSLockdownPolicy environment variable" -ForegroundColor Cyan

$lockdownFound = $false
foreach ($scope in @('Process', 'User', 'Machine')) {
    $value = $null
    try { $value = [Environment]::GetEnvironmentVariable('__PSLockdownPolicy', $scope) } catch { }
    if ($value) {
        $lockdownFound = $true
        Write-Host "  $scope : $value" -ForegroundColor Red
    } else {
        Write-Host "  $scope : (not set)" -ForegroundColor Gray
    }
}

if ($lockdownFound) {
    Write-Host ""
    Write-Host "  FOUND. This is forcing CLM and is the simplest cause to resolve." -ForegroundColor DarkYellow
    Write-Host "  A value of 4 means enforce; 8 means audit. Removing it requires admin" -ForegroundColor DarkYellow
    Write-Host "  rights for Machine scope, and a new shell afterwards." -ForegroundColor DarkYellow
    Write-Host "  Do not remove it without asking - it may be deliberate policy." -ForegroundColor DarkYellow
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 3. Cause 2 - WDAC / Device Guard code integrity policy
#    The most common cause on a managed SOE. A script that does not pass the policy runs
#    in CLM rather than being blocked outright.
# ---------------------------------------------------------------------------------------
Write-Host "CAUSE 2: WDAC / Device Guard code integrity" -ForegroundColor Cyan

$dg = $null
if (Get-Command -Name Get-CimInstance -ErrorAction SilentlyContinue) {
    try {
        $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace 'root\Microsoft\Windows\DeviceGuard' -ErrorAction Stop
    } catch {
        Write-Host "  Could not query Device Guard WMI (this is normal on non-Windows or where WDAC is absent)" -ForegroundColor Gray
    }
} else {
    Write-Host "  Get-CimInstance unavailable - not Windows, or blocked" -ForegroundColor Gray
}

if ($dg) {
    $ciStatus = switch ($dg.CodeIntegrityPolicyEnforcementStatus) {
        0 { 'Off' } 1 { 'Audit mode' } 2 { 'Enforced' } default { "Unknown ($($dg.CodeIntegrityPolicyEnforcementStatus))" }
    }
    $umciStatus = switch ($dg.UsermodeCodeIntegrityPolicyEnforcementStatus) {
        0 { 'Off' } 1 { 'Audit mode' } 2 { 'Enforced' } default { "Unknown ($($dg.UsermodeCodeIntegrityPolicyEnforcementStatus))" }
    }

    Write-Host "  Kernel code integrity   : $ciStatus"
    Write-Host "  User-mode code integrity: $umciStatus" -ForegroundColor $(if ($umciStatus -eq 'Enforced') { 'Red' } else { 'Gray' })

    if ($dg.SecurityServicesConfigured) {
        Write-Host "  Services configured     : $($dg.SecurityServicesConfigured -join ', ')"
    }
    if ($dg.SecurityServicesRunning) {
        Write-Host "  Services running        : $($dg.SecurityServicesRunning -join ', ')"
    }

    if ($umciStatus -eq 'Enforced') {
        Write-Host ""
        Write-Host "  LIKELY CAUSE. User-mode code integrity is enforced, so unsigned or" -ForegroundColor DarkYellow
        Write-Host "  untrusted scripts run in CLM. A script signed by a certificate the" -ForegroundColor DarkYellow
        Write-Host "  policy trusts will run in FullLanguage." -ForegroundColor DarkYellow
    }
}

# Active WDAC policies on disk
$ciPaths = @(
    "$env:SystemRoot\System32\CodeIntegrity\CiPolicies\Active"
    "$env:SystemRoot\System32\CodeIntegrity"
)
foreach ($p in $ciPaths) {
    if (Test-Path $p) {
        $policies = Get-ChildItem -Path $p -Filter '*.cip' -ErrorAction SilentlyContinue
        if ($policies) {
            Write-Host "  Policies in $p :" -ForegroundColor Gray
            foreach ($pol in $policies) { Write-Host "    $($pol.Name)" -ForegroundColor Gray }
        }
    }
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 4. Cause 3 - AppLocker script rules
#    AppLocker in Allow mode puts non-allowed scripts into CLM.
# ---------------------------------------------------------------------------------------
Write-Host "CAUSE 3: AppLocker script rules" -ForegroundColor Cyan

$appLockerSvc = $null
if (Get-Command -Name Get-Service -ErrorAction SilentlyContinue) {
    $appLockerSvc = Get-Service -Name AppIDSvc -ErrorAction SilentlyContinue
}
if ($appLockerSvc) {
    Write-Host "  AppIDSvc status : $($appLockerSvc.Status)"
} else {
    Write-Host "  AppIDSvc        : not present" -ForegroundColor Gray
}

if (Get-Command -Name Get-AppLockerPolicy -ErrorAction SilentlyContinue) {
    $policyXml = $null
    try { $policyXml = Get-AppLockerPolicy -Effective -Xml -ErrorAction Stop } catch { }

    if ($policyXml) {
        if ($policyXml -match 'Type="Script"') {
            Write-Host "  Script rules    : PRESENT" -ForegroundColor DarkYellow
            if ($policyXml -match 'Type="Script"[^>]*EnforcementMode="Enabled"') {
                Write-Host "  Enforcement     : Enabled" -ForegroundColor Red
                Write-Host ""
                Write-Host "  LIKELY CAUSE. AppLocker script enforcement puts scripts that do not" -ForegroundColor DarkYellow
                Write-Host "  match an Allow rule into CLM." -ForegroundColor DarkYellow
            } elseif ($policyXml -match 'Type="Script"[^>]*EnforcementMode="AuditOnly"') {
                Write-Host "  Enforcement     : Audit only (should not cause CLM)" -ForegroundColor Gray
            }
        } else {
            Write-Host "  Script rules    : none found" -ForegroundColor Gray
        }
    } else {
        Write-Host "  No effective AppLocker policy returned" -ForegroundColor Gray
    }
} else {
    Write-Host "  Get-AppLockerPolicy unavailable" -ForegroundColor Gray
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 5. Cause 4 - third-party application control (Airlock Digital and similar)
# ---------------------------------------------------------------------------------------
Write-Host "CAUSE 4: Third-party application control" -ForegroundColor Cyan

$acProducts = @(
    @{ Name = 'Airlock Digital'; Service = 'AirlockClient' }
    @{ Name = 'Airlock Digital (alt)'; Service = 'Airlock' }
    @{ Name = 'Carbon Black';    Service = 'CbDefense' }
    @{ Name = 'Ivanti / AppSense'; Service = 'AppSenseApplicationManager' }
    @{ Name = 'ThreatLocker';    Service = 'ThreatLockerService' }
)

$found = $false
$hasGetService = [bool](Get-Command -Name Get-Service -ErrorAction SilentlyContinue)
foreach ($product in $acProducts) {
    $svc = $null
    if ($hasGetService) { $svc = Get-Service -Name $product.Service -ErrorAction SilentlyContinue }
    if ($svc) {
        $found = $true
        Write-Host "  $($product.Name) : $($svc.Status)" -ForegroundColor DarkYellow
    }
}
if (-not $found) {
    Write-Host "  None of the checked products detected by service name" -ForegroundColor Gray
    Write-Host "  (this list is not exhaustive - check what the SOE actually deploys)" -ForegroundColor Gray
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 6. Execution policy and script trust
# ---------------------------------------------------------------------------------------
Write-Host "EXECUTION POLICY" -ForegroundColor Cyan
$executionPolicyBlocking = $false
if (Get-Command -Name Get-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Get-ExecutionPolicy -List | ForEach-Object {
        Write-Host ("  {0,-16}: {1}" -f $_.Scope, $_.ExecutionPolicy)
    }

    $effective = Get-ExecutionPolicy
    $executionPolicyBlocking = $effective -in @('AllSigned', 'Restricted')
    Write-Host ("  {0,-16}: {1}" -f 'Effective', $effective) -ForegroundColor $(if ($executionPolicyBlocking) { 'Red' } else { 'Green' })

    if ($executionPolicyBlocking) {
        Write-Host ""
        Write-Host "  BLOCKING. This module is signed via a file catalog, not per-file" -ForegroundColor DarkYellow
        Write-Host "  Authenticode (see Documentation\Verification.md), so $effective refuses" -ForegroundColor DarkYellow
        Write-Host "  to run it at all." -ForegroundColor DarkYellow
        Write-Host "    Fix: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned" -ForegroundColor Gray
    }
} else {
    Write-Host "  Get-ExecutionPolicy unavailable (non-Windows)" -ForegroundColor Gray
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 7. Module location and signing state
# ---------------------------------------------------------------------------------------
Write-Host "MODULE LOCATION AND SIGNING" -ForegroundColor Cyan
Write-Host "  Path : $ModulePath"

if (Test-Path $ModulePath) {
    # Mark-of-the-Web silently blocks module import and is easy to miss.
    $blocked = @()
    Get-ChildItem -Path $ModulePath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $zone = Get-Item -Path $_.FullName -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
        if ($zone) { $blocked += $_.Name }
    }
    if ($blocked.Count -gt 0) {
        Write-Host "  Mark-of-the-Web : $($blocked.Count) file(s) BLOCKED" -ForegroundColor Red
        Write-Host "    Fix: Get-ChildItem -Path '$ModulePath' -Recurse | Unblock-File" -ForegroundColor DarkYellow
    } else {
        Write-Host "  Mark-of-the-Web : none" -ForegroundColor Green
    }

    # Signature state of the entry point / catalog
    foreach ($target in @('runMe.ps1', 'OAG-FileCatalog.cat')) {
        $full = Join-Path -Path $ModulePath -ChildPath $target
        if (Test-Path $full) {
            $sig = $null
            if (Get-Command -Name Get-AuthenticodeSignature -ErrorAction SilentlyContinue) {
                $sig = Get-AuthenticodeSignature -FilePath $full -ErrorAction SilentlyContinue
            }
            if ($sig) {
                $col = if ($sig.Status -eq 'Valid') { 'Green' } else { 'DarkYellow' }
                Write-Host "  $target : $($sig.Status)" -ForegroundColor $col
                if ($sig.SignerCertificate) {
                    Write-Host "      Signer: $($sig.SignerCertificate.Subject)" -ForegroundColor Gray
                }
            }
        } else {
            Write-Host "  $target : not present" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "  Path does not exist" -ForegroundColor Red
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 8. Required PowerShell version and modules
#    Checks presence only - not a specific pinned version. Every module here is the union
#    of what the five reports' -Modules lists in runMe.ps1 need; keep this list in sync if
#    that ever changes.
# ---------------------------------------------------------------------------------------
Write-Host "REQUIRED MODULES" -ForegroundColor Cyan

$requiredModulesMissing = $false

$psMajor = $PSVersionTable.PSVersion.Major
Write-Host "  PowerShell : $($PSVersionTable.PSVersion)" -ForegroundColor $(if ($psMajor -ge 7) { 'Green' } else { 'Red' })
if ($psMajor -lt 7) {
    $requiredModulesMissing = $true
    Write-Host "    MISSING. This module requires PowerShell 7." -ForegroundColor DarkYellow
    Write-Host "    Fix: winget install --id Microsoft.PowerShell --source winget --installer-type wix" -ForegroundColor Gray
}
Write-Host ""

$requiredModuleNames = @(
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.DirectoryObjects'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.SignIns'
    'Microsoft.Graph.Identity.Governance'
    'Microsoft.Graph.Identity.DirectoryManagement'
    'Microsoft.Graph.Reports'
    'ExchangeOnlineManagement'
)

$missingModules = @()
foreach ($moduleName in $requiredModuleNames) {
    $installed = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($installed) {
        Write-Host "  $moduleName : $($installed.Version)" -ForegroundColor Green
    } else {
        Write-Host "  $moduleName : NOT INSTALLED" -ForegroundColor Red
        $missingModules += $moduleName
    }
}

if ($missingModules.Count -gt 0) {
    $requiredModulesMissing = $true
    Write-Host ""
    Write-Host "  MISSING $($missingModules.Count) module(s)." -ForegroundColor DarkYellow
    Write-Host "    Fix: Install-Module -Name $($missingModules -join ', ')" -ForegroundColor Gray
}
Write-Host ""

# ---------------------------------------------------------------------------------------
# 9. What actually breaks
# ---------------------------------------------------------------------------------------
Write-Host "CAPABILITY TEST" -ForegroundColor Cyan
Write-Host "  Checking the specific operations the export module depends on."
Write-Host ""

# NOTE: this section uses plain hashtables, not [PSCustomObject]. Testing confirmed that
# [PSCustomObject]@{} is itself blocked in Constrained Language Mode ("Only core types are
# supported in this language mode"), which is worth knowing in its own right: the export
# module builds its run context and nearly every exported record as PSCustomObject, so that
# alone makes it unable to run under CLM.
$tests = @()

$ok = $false
try { $null = & { [PSCustomObject]@{ a = 1 } }; $ok = $true } catch { }
$tests += @{ Feature = '[PSCustomObject] literal'; Works = $ok }

$ok = $false
try { $null = & { [System.Collections.Generic.List[object]]::new() }; $ok = $true } catch { }
$tests += @{ Feature = 'Generic List (run context)'; Works = $ok }

$ok = $false
try { $null = & { [System.Collections.Generic.HashSet[string]]::new() }; $ok = $true } catch { }
$tests += @{ Feature = 'HashSet'; Works = $ok }

$ok = $false
try { $null = & { New-Object -TypeName PSObject }; $ok = $true } catch { }
$tests += @{ Feature = 'New-Object'; Works = $ok }

$ok = $false
try { $null = & { [math]::Round(1.234, 2) }; $ok = $true } catch { }
$tests += @{ Feature = '.NET static methods'; Works = $ok }

$ok = $false
try { $null = & { (Get-Date).ToString('yyyyMMdd') }; $ok = $true } catch { }
$tests += @{ Feature = 'Method calls on DateTime'; Works = $ok }

$ok = $false
try { $null = & { @{ x = 1 } }; $ok = $true } catch { }
$tests += @{ Feature = 'Hashtable literal'; Works = $ok }

$blockedCount = 0
foreach ($t in $tests) {
    if ($t.Works) {
        Write-Host ("  {0,-28}: available" -f $t.Feature) -ForegroundColor Green
    } else {
        Write-Host ("  {0,-28}: BLOCKED" -f $t.Feature) -ForegroundColor Red
        $blockedCount = $blockedCount + 1
    }
}

Write-Host ""
Write-Host "===================================================================="
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host ""

if ($mode -ne 'FullLanguage') {
    Write-Host "  Session is in $mode with $blockedCount blocked capability/ies." -ForegroundColor Red
    Write-Host ""
    Write-Host "  The module cannot be made to work in this state. It is not a matter of" -ForegroundColor DarkYellow
    Write-Host "  rewriting it to avoid restricted features - the Microsoft Graph SDK has" -ForegroundColor DarkYellow
    Write-Host "  the same dependencies, so the restriction reaches past our own code." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  Take the cause identified above to whoever administers application" -ForegroundColor DarkYellow
    Write-Host "  control and request an approved execution path. See Docs\README.md." -ForegroundColor DarkYellow
} elseif ($executionPolicyBlocking) {
    Write-Host "  Session is in FullLanguage, but the execution policy above will still" -ForegroundColor Red
    Write-Host "  refuse to run this module. See the fix under EXECUTION POLICY." -ForegroundColor Red
} elseif ($requiredModulesMissing) {
    Write-Host "  Session is in FullLanguage and execution policy is fine, but PowerShell" -ForegroundColor DarkYellow
    Write-Host "  itself or a required module is missing. See the fix(es) under REQUIRED MODULES." -ForegroundColor DarkYellow
} else {
    Write-Host "  Session is in FullLanguage, execution policy allows it, and every" -ForegroundColor Green
    Write-Host "  required module is installed. The module should run." -ForegroundColor Green
}
Write-Host ""
