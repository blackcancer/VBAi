[CmdletBinding()]
param(
 [Parameter(Mandatory)][string]$PendingReport,
 [Parameter(Mandatory)][Guid]$ExpectedMvid,
 [Parameter(Mandatory)][ValidatePattern('^[a-fA-F0-9]{64}$')][string]$ExpectedAssemblySha256,
 [string]$CdbPath='D:\Windows Kits\10\Debuggers\x64\cdb.exe',
 [string]$DebuggerPreflightReport,
 [ValidateRange(10,30)][int]$MaxSeconds=30,
 [switch]$Execute
)
$ErrorActionPreference='Stop'
if ($PendingReport -notmatch '^[A-Za-z]:[\\/]' -or [IO.Path]::GetFileName($PendingReport) -ne 'teardown.pending.json') { throw 'Exact absolute local pending report required.' }
if ((Get-Item -LiteralPath $PendingReport).Length -gt 16384) { throw 'Pending report exceeds its bound.' }
$pending=Get-Content -LiteralPath $PendingReport -Raw -Encoding UTF8 | ConvertFrom-Json
$directory=[IO.Path]::GetDirectoryName($PendingReport)
if ($pending.Nonce -notmatch '^[a-f0-9]{32}$' -or [IO.Path]::GetFileName($directory) -notmatch '^[a-f0-9]{32}$' -or
    $pending.Root -ne $directory -or [int]$pending.ProcessId -le 0 -or
    $pending.Scenario -notin @('NativeVariantArraysRoundTripWithBoundsAndOneInvocation','NativeParamArrayCallsPreserveArityValuesAndSingleInvocation') -or
    [IO.Path]::GetFileName($pending.Executable) -ine 'EXCEL.EXE' -or $pending.Executable -notmatch '^[A-Za-z]:[\\/]' -or
    [IO.Path]::GetFileName($pending.AssemblyPath) -ine 'VBAi.dll' -or $pending.AssemblyPath -notmatch '^[A-Za-z]:[\\/]') { throw 'Pending owned procedure-value identity/path mismatch.' }
$started=[DateTime]::Parse($pending.ProcessStartedUtc).ToUniversalTime()
if ([Guid]$pending.AssemblyMvid -ne $ExpectedMvid -or
    (Get-FileHash -LiteralPath $pending.AssemblyPath -Algorithm SHA256).Hash -ne $ExpectedAssemblySha256) { throw 'Loaded candidate evidence mismatch.' }
if (-not [IO.File]::Exists($CdbPath)) { throw 'Existing debugger unavailable; no installation.' }
$plan=[ordered]@{Mode='PREPARE_ONLY';ProcessId=[int]$pending.ProcessId;ProcessStartedUtc=$pending.ProcessStartedUtc;
 Nonce=$pending.Nonce;AssemblyMvid=$ExpectedMvid.ToString();ExpectedAssemblySha256=$ExpectedAssemblySha256;
 Scenario=$pending.Scenario;CleanupCallsIssuedByController=0;MemoryDumps=0;HostTerminationCalls=0;
 Scope='Exact owned disposable Excel; failfast exception/context/stack only. No COM calls, cleanup replay, full dump or global WER changes.'}
if (-not $Execute) { $plan | ConvertTo-Json -Depth 5; return }
. (Join-Path $PSScriptRoot 'OwnedTeardownTrace.Common.ps1')
if (-not [IO.Path]::IsPathRooted($DebuggerPreflightReport)) { throw 'Measured exception-handler preflight required before attach.' }
$preflight=Get-Content -LiteralPath $DebuggerPreflightReport -Raw -Encoding UTF8|ConvertFrom-Json
$common=Join-Path $PSScriptRoot 'OwnedTeardownTrace.Common.ps1'
if ($preflight.State -ne 'PASS' -or $preflight.Mode -ne 'TEARDOWN_EXCEPTION_PREFLIGHT' -or $preflight.ExceptionCaptureVerified -ne $true -or
    $preflight.NormalHelperExitCode -ne 0 -or $preflight.NormalHelperDetachedVerified -ne $true -or
    $preflight.FailfastHelperExitCodeHex -ne '0xC0000409' -or $preflight.SyntheticFailfastRequests -ne 1 -or
    $preflight.CdbSha256 -ne (Get-FileHash -LiteralPath $CdbPath -Algorithm SHA256).Hash -or
    $preflight.CommonScriptSha256 -ne (Get-FileHash -LiteralPath $common -Algorithm SHA256).Hash) { throw 'Different/incomplete helper handler preflight; no attachment.' }
$target=[Diagnostics.Process]::GetProcessById([int]$pending.ProcessId)
if ($target.StartTime.ToUniversalTime() -ne $started -or (Get-OwnedTeardownImage $target) -ine $pending.Executable) { throw 'PID/start/image mismatch; no attachment.' }
$startupPath=Join-Path $directory 'startup.json'
if((Get-Item -LiteralPath $startupPath).Length -gt 65536){throw 'Owned startup evidence exceeds its bound.'}
$startup=Get-Content -LiteralPath $startupPath -Raw -Encoding UTF8|ConvertFrom-Json
$loaded=$startup.BridgeStatus.Data
if($startup.Phase -ne 'Ready' -or $startup.Owned -ne $true -or [int]$startup.ProcessId -ne $target.Id -or
 [DateTime]::Parse($startup.HostStartedUtc).ToUniversalTime() -ne $started -or $startup.FixtureRoot -ne $directory -or
 $startup.HostExecutable -ine $pending.Executable -or $startup.BridgeStatus.Ok -ne $true -or
 [int]$loaded.HostProcessId -ne $target.Id -or [Guid]$loaded.AssemblyModuleVersionId -ne $ExpectedMvid -or
 $loaded.AssemblyPath -ine $pending.AssemblyPath){throw 'Actual owned startup bridge candidate evidence mismatch.'}
$plan.LoadedCandidateEvidence='OwnedFixture.ReadyStartupBridgeStatus'
$plan.AdditionalBridgeCalls=0
$present=$false
if (-not [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -or $present) { throw 'Existing or unavailable debugger; no attachment.' }
foreach($name in @('teardown.armed.json','teardown.detached.json','teardown-controller.json','teardown.cdb.log','teardown.commands.txt')) {
 if ([IO.File]::Exists((Join-Path $directory $name))) { throw 'Diagnostic output already exists; no re-attachment.' }
}
$session=$null;$primary=$null;$cleanup=$null;$plan.Mode='EXECUTE';$plan.State='ATTACHING'
try {
 $session=Start-OwnedTeardownDebugger $CdbPath $target $directory $pending.Nonce
 $plan.DebuggerProcessId=$session.Process.Id
 Wait-OwnedTeardownArmed $session $target $pending.Nonce
 Write-TeardownMarker (Join-Path $directory 'teardown.armed.json') @{Phase='Armed';ProcessId=$pending.ProcessId;
   ProcessStartedUtc=$pending.ProcessStartedUtc;Nonce=$pending.Nonce;AssemblyMvid=$pending.AssemblyMvid}
 $plan.AttachmentObserved=$true
 $watch=[Diagnostics.Stopwatch]::StartNew()
 while ($watch.Elapsed.TotalSeconds -lt $MaxSeconds -and -not $session.Process.HasExited) { Start-Sleep -Milliseconds 100 }
} catch { $primary=$_.Exception.ToString();$plan.Failure=$primary }
finally {
 if ($session -ne $null) {
  try {
   Stop-OwnedTeardownDebugger $session $target
   $plan.DebuggerDetachedVerified=(-not $target.HasExited)
   $plan.HostExitObserved=$target.HasExited
   if($target.HasExited){$plan.HostExitCodeHex='0x'+([uint32]([int64]$target.ExitCode -band 4294967295)).ToString('X8')}
   Write-TeardownMarker (Join-Path $directory 'teardown.detached.json') @{Phase='Detached';ProcessId=$pending.ProcessId;
     ProcessStartedUtc=$pending.ProcessStartedUtc;Nonce=$pending.Nonce;AssemblyMvid=$pending.AssemblyMvid}
  } catch { $cleanup=$_.Exception.ToString();$plan.CleanupFailure=$cleanup }
  if($session.Process.HasExited){$plan.DebuggerExitCode=$session.Process.ExitCode}
  if([IO.File]::Exists($session.Log)){
   $text=Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode
   $plan.ExceptionCaptureVerified=Test-TeardownExceptionCapture $text
  }
 }
 $plan.State=if($primary -eq $null -and $cleanup -eq $null){'TERMINAL'}else{'FAILED_PRESERVED'}
 $plan.CompletedUtc=[DateTime]::UtcNow.ToString('o')
 Write-TeardownJson (Join-Path $directory 'teardown-controller.json') $plan
}
if($primary -ne $null -or $cleanup -ne $null){throw 'Diagnostic failed; primary and cleanup evidence retained; no native action replay.'}
