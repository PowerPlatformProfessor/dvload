<#
.SYNOPSIS
  Smoke-tests the unattended (app-only) path against a real Dataverse
  environment. Windows-only, like the unattended path itself.

.DESCRIPTION
  Covers TEST-PROTOCOL.md sections 4.5-4.8 (app-only auth, precedence,
  app-logout), 5 (validation) and 6 (dry-run / insert), plus the identity
  split introduced when the UI moved behind `dvload serve`.

  Safe by default: everything runs as a dry run and nothing reaches
  Dataverse's write endpoints unless you pass -Write.

  The secret is only ever read from an environment variable. It is never a
  parameter, never echoed, and never written to disk by this script — the
  CLI puts it in the DPAPI-backed secure store, which is the point.

.PARAMETER EnvUrl
  Dataverse environment URL. Defaults to the one in the sample mapping.

.PARAMETER ClientId
  App registration (client) id of the confidential client. Omit to reuse
  credentials already stored by a previous `dvload app-login`.

.PARAMETER TenantId
  Tenant id for that app registration.

.PARAMETER SecretEnv
  Name of the environment variable holding the client secret.

.PARAMETER Write
  Actually create records. Prompts first unless -Force. Without this, the
  run stops at --dry-run.

.PARAMETER Force
  Skip the confirmation prompt for -Write. For CI.

.PARAMETER RequireAppOnly
  Fail instead of falling back to a delegated partial pass. Use in CI, where
  "it passed as whoever was signed in" is not the thing being tested.

.PARAMETER KeepCredentials
  Leave the app-only credentials in the secure store afterwards. By default
  the script runs `app-logout` at the end if it was the thing that stored
  them, so a test run doesn't silently change how later `dvload run`
  invocations authenticate.

.NOTES
  Run it, don't dot-source it: `.\tests\unattended-smoke.ps1`, not
  `. .\tests\unattended-smoke.ps1`. Dot-sourcing works, but it executes in
  your session's scope, so the exit code is returned rather than set.

.EXAMPLE
  $env:DVLOAD_SECRET = "<secret>"
  .\tests\unattended-smoke.ps1 -ClientId <app-id> -TenantId <tenant-id>

.EXAMPLE
  # Full pass including writes, against an org you don't mind touching
  $env:DVLOAD_SECRET = "<secret>"
  .\tests\unattended-smoke.ps1 -ClientId <id> -TenantId <id> -Write
#>

[CmdletBinding()]
param(
  [string] $EnvUrl   = "https://org1f722cc4.crm.dynamics.com",
  [string] $ClientId,
  [string] $TenantId,
  [string] $SecretEnv = "DVLOAD_SECRET",
  [string] $Mapping   = "$PSScriptRoot\dummy-data\contacts.dvmap.json",
  [string] $Workbook  = "$PSScriptRoot\dummy-data\contacts.xlsx",
  [string] $Dvload    = "dvload",
  [switch] $Write,
  [switch] $Force,
  [switch] $KeepCredentials,
  [switch] $RequireAppOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
  throw "Windows-only: the unattended path uses the DPAPI secure store and Task Scheduler."
}

# ---------------------------------------------------------------------------
# Result tracking
# ---------------------------------------------------------------------------

$script:Results = [System.Collections.Generic.List[object]]::new()
$script:StoredCredentials = $false

function Add-Result {
  param([string] $Id, [string] $Name, [string] $Status, [string] $Detail = "")
  $script:Results.Add([pscustomobject]@{
    Test = $Id; Name = $Name; Status = $Status; Detail = $Detail
  })
  $colour = switch ($Status) { "PASS" { "Green" } "FAIL" { "Red" } default { "Yellow" } }
  Write-Host ("  [{0,-4}] {1} {2}" -f $Status, $Name, $(if ($Detail) { "- $Detail" } else { "" })) -ForegroundColor $colour
}

<#
  Runs the CLI and captures stdout, stderr and the exit code together.

  Deliberately does not use $ErrorActionPreference/try-catch for CLI failures:
  a non-zero exit is a test result, not a script error, and several checks
  below assert that a command fails.
#>
function Invoke-Dvload {
  param([string[]] $Arguments)

  # The call operator, NOT Start-Process.
  #
  # On Windows `dvload` is an npm shim — `dvload.ps1`, `dvload.cmd` and an
  # extensionless shell script sit side by side, and `Get-Command` usually
  # resolves the .ps1. Start-Process passes whatever it's given to the Win32
  # loader, which can only execute real binaries, so a shim fails with
  # "%1 is not a valid Win32 application". `&` goes through PowerShell's own
  # command resolution, which understands all three.
  $errFile = (New-TemporaryFile).FullName
  try {
    # Redirect stderr to a file so the two streams stay separable; stdout
    # comes back as the pipeline value. Merging them would corrupt the JSON
    # that --json emits on stdout.
    $stdoutLines = & $Dvload @Arguments 2> $errFile
    $code = $LASTEXITCODE

    $out = if ($null -eq $stdoutLines) { "" } else { ($stdoutLines | Out-String) }
    $err = Get-Content $errFile -Raw -ErrorAction SilentlyContinue
    if ($null -eq $err) { $err = "" }
    # A shim that fails before exec leaves $LASTEXITCODE unset.
    if ($null -eq $code) { $code = 0 }

    [pscustomobject]@{
      ExitCode = $code
      Stdout   = $out
      Stderr   = $err
    }
  } finally {
    Remove-Item $errFile -Force -ErrorAction SilentlyContinue
  }
}

# ---------------------------------------------------------------------------
# 0. Preflight
# ---------------------------------------------------------------------------

Write-Host "`n=== 0. Preflight ===" -ForegroundColor Cyan

$cmd = Get-Command $Dvload -ErrorAction SilentlyContinue
if (-not $cmd) {
  throw "'$Dvload' not found on PATH. Pass -Dvload with a full path, or point it at the built entry point."
}
Add-Result "0.1" "dvload on PATH" "PASS" $cmd.Source

foreach ($f in @($Mapping, $Workbook)) {
  if (-not (Test-Path $f)) {
    throw "Missing test fixture: $f. Generate with: node tests\dummy-data\make-contacts.js"
  }
}
Add-Result "0.2" "Fixtures present" "PASS"

# The mapping carries its own environmentUrl and the CLI uses that, not -EnvUrl.
# A mismatch would mean the auth checks below and the actual run target
# different environments, which is a confusing way to fail.
$mappingJson = Get-Content $Mapping -Raw | ConvertFrom-Json
# Guarded because Set-StrictMode turns a missing property into a terminating
# error, which would report as a script bug rather than a bad mapping.
if (-not $mappingJson.PSObject.Properties.Name.Contains("environmentUrl")) {
  throw "$Mapping has no environmentUrl. It is required by the schema."
}
if ($mappingJson.environmentUrl -ne $EnvUrl) {
  Add-Result "0.3" "Mapping targets -EnvUrl" "FAIL" `
    "mapping says $($mappingJson.environmentUrl), -EnvUrl says $EnvUrl"
  throw "Aborting: fix the mapping's environmentUrl or pass a matching -EnvUrl."
}
Add-Result "0.3" "Mapping targets -EnvUrl" "PASS" $EnvUrl

# ---------------------------------------------------------------------------
# 1. App-only credentials (TEST-PROTOCOL 4.5)
# ---------------------------------------------------------------------------

Write-Host "`n=== 1. App-only auth ===" -ForegroundColor Cyan

if ($ClientId -and $TenantId) {
  if (-not (Test-Path "env:$SecretEnv")) {
    throw "Environment variable '$SecretEnv' is not set. Set it to the client secret, e.g. `$env:$SecretEnv = '<secret>'"
  }
  $r = Invoke-Dvload @("app-login", "--env", $EnvUrl, "--client-id", $ClientId,
                       "--tenant-id", $TenantId, "--secret-env", $SecretEnv)
  if ($r.ExitCode -ne 0) {
    Add-Result "4.5" "app-login" "FAIL" $r.Stderr.Trim()
    throw "app-login failed; nothing else will be meaningful."
  }
  $script:StoredCredentials = $true
  Add-Result "4.5" "app-login" "PASS" "credentials in secure store"
} else {
  Add-Result "4.5" "app-login" "SKIP" "no -ClientId/-TenantId; reusing stored credentials"
}

# whoami is the most informative check here: it proves the credentials, the
# tenant and the security roles are all correct before any data moves.
#
# Three outcomes, deliberately distinguished — an earlier version collapsed
# them and reported "cannot acquire an app-only token" when the token probe
# had actually succeeded, just as a delegated user. Being unable to tell "no
# credentials" from "different credentials" is exactly the confusion this
# script exists to remove.
$r = Invoke-Dvload @("whoami", "--env", $EnvUrl)
$who = "$($r.Stdout)$($r.Stderr)"
# Greedy `.*` so it binds to the LAST colon on the line: the environment URL
# contains "https:", and a lazy match would stop there and capture nothing.
$mode = if ($who -match "Auth mode for .*:\s*(\w+)") { $Matches[1] } else { "unknown" }
$tokenOk = $who -match "Token acquisition: OK"

$script:AppOnly = $false

if ($r.ExitCode -ne 0 -or $mode -eq "none" -or -not $tokenOk) {
  Add-Result "4.2" "whoami + token probe" "FAIL" $who.Trim()
  throw @"
Aborting: no usable credentials for $EnvUrl.

  Unattended (app-only):  dvload app-login --env $EnvUrl --client-id <id> --tenant-id <id> --secret-env $SecretEnv
                          (needs an app registration + Application User — docs/SCHEDULED-RUNS.md steps 1-2)
  Interactive:            dvload login --env $EnvUrl
"@
}

if ($mode -eq "appOnly") {
  $script:AppOnly = $true
  Add-Result "4.2" "whoami: app-only + token OK" "PASS"
} else {
  # Delegated works unattended — the refresh token lasts ~90 days idle — it
  # is just fragile, which is why `dvload schedule` recommends app-login.
  # So this is a real partial pass: the whole load pipeline gets exercised,
  # only the unattended *identity* doesn't.
  if ($RequireAppOnly) {
    Add-Result "4.2" "whoami: app-only + token OK" "FAIL" "mode is '$mode', -RequireAppOnly was set"
    throw "Aborting: -RequireAppOnly was specified but no app-only credentials are stored for $EnvUrl."
  }
  Add-Result "4.2" "whoami: token OK (mode: $mode)" "PASS" "no app-only credentials — partial pass"
  Write-Host "  Running as the signed-in user, not an Application User." -ForegroundColor Yellow
  Write-Host "  Pass -ClientId/-TenantId to cover the real unattended identity." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 2. Validation (TEST-PROTOCOL 5)
# ---------------------------------------------------------------------------

Write-Host "`n=== 2. Validation ===" -ForegroundColor Cyan

$r = Invoke-Dvload @("validate", $Mapping)
if ($r.ExitCode -eq 0) {
  Add-Result "5.1" "validate (schema + live metadata)" "PASS"
} else {
  Add-Result "5.1" "validate (schema + live metadata)" "FAIL" "$($r.Stdout)$($r.Stderr)".Trim()
}

# ---------------------------------------------------------------------------
# 3. Dry run (TEST-PROTOCOL 6)
# ---------------------------------------------------------------------------

Write-Host "`n=== 3. Dry run ===" -ForegroundColor Cyan

$r = Invoke-Dvload @("run", $Mapping, "-w", $Workbook, "--dry-run", "--json", "--non-interactive")
if ($r.ExitCode -ne 0) {
  Add-Result "6.1" "run --dry-run" "FAIL" "$($r.Stdout)$($r.Stderr)".Trim()
} else {
  $dry = $r.Stdout | ConvertFrom-Json
  if ($dry.total -gt 0) {
    Add-Result "6.1" "run --dry-run" "PASS" "$($dry.total) row(s) planned, 0 written"
  } else {
    Add-Result "6.1" "run --dry-run" "FAIL" "planned 0 rows — check sourceTable matches the workbook"
  }
}

# ---------------------------------------------------------------------------
# 4. Identity split (the sidecar change)
# ---------------------------------------------------------------------------

Write-Host "`n=== 4. Identity precedence ===" -ForegroundColor Cyan

# 4.7: with app-only configured, the default is app-only and --user forces
# delegated. This matters more since the UI moved behind `dvload serve`: the
# sidecar pins delegated, so the pane writes as you while an unattended run
# writes as the Application User. Same mapping, two identities, by design.
#
# Only meaningful when both are configured — with delegated alone there is no
# precedence to test, because everything is already delegated.
if ($script:AppOnly) {
  $r = Invoke-Dvload @("run", $Mapping, "-w", $Workbook, "--dry-run", "--json", "--user", "--non-interactive")
  if ($r.ExitCode -eq 0) {
    Add-Result "4.7" "--user forces delegated" "PASS" "both identities available; --user selected delegated"
  } else {
    # Expected when nobody has run `dvload login` on this machine.
    Add-Result "4.7" "--user forces delegated" "SKIP" "no delegated session (run 'dvload login' to cover this)"
  }
  Write-Host "  Records created by the task pane are owned by the signed-in user;" -ForegroundColor DarkGray
  Write-Host "  records created by a scheduled run are owned by the Application User." -ForegroundColor DarkGray
} else {
  Add-Result "4.7" "--user forces delegated" "SKIP" "delegated only — no app-only credentials to take precedence over"
  Write-Host "  Only one identity configured, so the pane and unattended runs write as the same user." -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# 5. Write pass (opt-in)
# ---------------------------------------------------------------------------

$doWrite = [bool] $Write

$identity = if ($script:AppOnly) { "the Application User" } else { "the signed-in user" }

if ($doWrite -and -not $Force) {
  Write-Host "`n=== 5. Write pass ===" -ForegroundColor Cyan
  $expectedHost = ([uri] $EnvUrl).Host
  Write-Host "About to CREATE records in $EnvUrl as $identity." -ForegroundColor Yellow
  # Typing the host, rather than y/N, so this can't be waved through by
  # reflex on the wrong environment.
  $answer = Read-Host "Type '$expectedHost' to confirm"
  if ($answer -ne $expectedHost) {
    Add-Result "6.2" "run (insert)" "SKIP" "not confirmed"
    $doWrite = $false
  }
}

if ($doWrite) {
  Write-Host "`n=== 5. Write pass ===" -ForegroundColor Cyan
  $r = Invoke-Dvload @("run", $Mapping, "-w", $Workbook, "--json", "--non-interactive")
  if ($r.ExitCode -ne 0) {
    Add-Result "6.2" "run (insert)" "FAIL" "$($r.Stdout)$($r.Stderr)".Trim()
  } else {
    $run = $r.Stdout | ConvertFrom-Json
    if ($run.failed -eq 0 -and $run.succeeded -gt 0) {
      Add-Result "6.2" "run (insert)" "PASS" "$($run.succeeded) succeeded, $($run.created) created"
      Write-Host "  Verify ownership: the new contacts' 'Created By' should be $identity." -ForegroundColor DarkGray
    } else {
      Add-Result "6.2" "run (insert)" "FAIL" "$($run.succeeded) succeeded, $($run.failed) failed"
    }
  }
} elseif (-not $Write) {
  Add-Result "6.2" "run (insert)" "SKIP" "dry-run only; pass -Write to create records"
}

# ---------------------------------------------------------------------------
# 6. Teardown (TEST-PROTOCOL 4.8)
# ---------------------------------------------------------------------------

Write-Host "`n=== 6. Teardown ===" -ForegroundColor Cyan

if ($script:StoredCredentials -and -not $KeepCredentials) {
  $r = Invoke-Dvload @("app-logout", "--env", $EnvUrl)
  if ($r.ExitCode -eq 0) {
    Add-Result "4.8" "app-logout" "PASS" "secure store cleared"
  } else {
    Add-Result "4.8" "app-logout" "FAIL" $r.Stderr.Trim()
  }
} else {
  Add-Result "4.8" "app-logout" "SKIP" $(if ($KeepCredentials) { "-KeepCredentials" } else { "credentials were pre-existing" })
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$script:Results | Format-Table -AutoSize

$failed = @($script:Results | Where-Object Status -eq "FAIL").Count
$passed = @($script:Results | Where-Object Status -eq "PASS").Count
$skipped = @($script:Results | Where-Object Status -eq "SKIP").Count

Write-Host "$passed passed, $failed failed, $skipped skipped." -ForegroundColor $(if ($failed) { "Red" } else { "Green" })

if (-not $script:AppOnly) {
  Write-Host "`nPARTIAL PASS: ran as the signed-in user, not an Application User." -ForegroundColor Yellow
  Write-Host "The load pipeline is covered; the unattended identity is not. To cover it:" -ForegroundColor Yellow
  Write-Host "  `$env:$SecretEnv = '<secret>'" -ForegroundColor Yellow
  Write-Host "  .\tests\unattended-smoke.ps1 -ClientId <app-id> -TenantId <tenant-id>" -ForegroundColor Yellow
  Write-Host "  (app registration + Application User setup: docs/SCHEDULED-RUNS.md steps 1-2)" -ForegroundColor Yellow
}

if (-not $doWrite) {
  Write-Host "`nNothing was written to $EnvUrl." -ForegroundColor DarkGray
}
# Single quotes: a backtick inside a double-quoted PowerShell string is an
# escape character, so "`dvload schedule`" would silently lose both marks.
Write-Host 'Not covered here (needs Excel + Task Scheduler): --refresh, dvload schedule, and the task pane.' -ForegroundColor DarkGray
Write-Host "See TEST-PROTOCOL.md sections 18 and 19.`n" -ForegroundColor DarkGray

# Dot-sourcing (`. .\unattended-smoke.ps1`) runs this in the caller's scope,
# where `exit` would terminate their PowerShell session. Detect it and return
# instead — losing the exit code is a far smaller cost than closing the window
# someone is working in.
if ($MyInvocation.InvocationName -eq ".") {
  Write-Host "(dot-sourced: returning instead of exiting. Run as '.\tests\unattended-smoke.ps1' for a real exit code.)" -ForegroundColor DarkGray
  $global:LASTEXITCODE = $(if ($failed) { 1 } else { 0 })
  return
}

if ($failed) { exit 1 } else { exit 0 }
