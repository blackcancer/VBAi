[CmdletBinding()]
param([Parameter(Mandatory)][string]$HelperPath,[Parameter(Mandatory)][string]$OutputRoot,
 [string]$CdbPath='D:\Windows Kits\10\Debuggers\x64\cdb.exe',[switch]$Execute)
$ErrorActionPreference='Stop'
if(-not [IO.Path]::IsPathRooted($HelperPath) -or -not [IO.Path]::IsPathRooted($OutputRoot) -or
 [IO.Path]::GetFileName($HelperPath) -ne 'VBAi.OwnedTeardown.Helper.exe' -or -not [IO.File]::Exists($HelperPath) -or -not [IO.File]::Exists($CdbPath)){
 throw 'Exact existing disposable helper/debugger and absolute output required; no installation.'
}
$common=Join-Path $PSScriptRoot 'OwnedTeardownTrace.Common.ps1'
$report=[ordered]@{State='PREPARE_ONLY';Mode='TEARDOWN_EXCEPTION_PREFLIGHT';HelperPath=$HelperPath;HelperSha256=(Get-FileHash -LiteralPath $HelperPath).Hash;
 CdbPath=$CdbPath;CdbSha256=(Get-FileHash -LiteralPath $CdbPath).Hash;CommonScriptSha256=(Get-FileHash -LiteralPath $common).Hash;RegisterCaptureMode='LiveEventThread';
 CollectorProfile='AV_FIRSTCHANCE_FORWARD_AV_SECONDCHANCE_AND_FAILFAST';
 Scope='Four newly owned non-Office helpers: normal detach/STOP, one failfast, one software second-chance AV, one locally handled software first-chance AV. No Office, COM, dump or global policy.'}
if(-not $Execute){$report|ConvertTo-Json -Depth 4;return}
. $common
$trial=Join-Path $OutputRoot ([Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($trial)
$primary=$null;$cleanup=New-Object Collections.Generic.List[string];$report.State='STARTED'
foreach($mode in @('normal','failfast','secondchanceav','firstchanceav')){
 $target=$null;$session=$null
 try{
  $dir=Join-Path $trial $mode;[void][IO.Directory]::CreateDirectory($dir)
  $info=New-Object Diagnostics.ProcessStartInfo
  $info.FileName=$HelperPath;$info.UseShellExecute=$false;$info.CreateNoWindow=$true
  $info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
  $target=New-Object Diagnostics.Process;$target.StartInfo=$info;[void]$target.Start()
  $ready=$target.StandardOutput.ReadLineAsync()
  if(-not $ready.Wait(10000) -or $ready.Result -notmatch '^READY (\d+) (.+)$' -or [int]$Matches[1] -ne $target.Id -or
    [DateTime]::Parse($Matches[2]).ToUniversalTime() -ne $target.StartTime.ToUniversalTime() -or
    (Get-OwnedTeardownImage $target) -ine $HelperPath){throw 'Owned helper identity handshake failed; preserve without termination.'}
  $report[$mode+'HelperPid']=$target.Id;$report[$mode+'HelperStartedUtc']=$target.StartTime.ToUniversalTime().ToString('o')
  $present=$false
  if(-not [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -or $present){throw 'Helper already debugged or state unavailable; preserve without attachment.'}
  $nonce=[Guid]::NewGuid().ToString('N');$session=Start-OwnedTeardownDebugger $CdbPath $target $dir $nonce
  Wait-OwnedTeardownArmed $session $target $nonce
  if($mode -eq 'normal'){
   Stop-OwnedTeardownDebugger $session $target
   $report.NormalHelperDetachedVerified=$true
   $report.NormalStopBreakpointVerified=$session.StopBreakpointVerified
   $report.NormalDebuggerExitCode=$session.Process.ExitCode
   $target.StandardInput.WriteLine('STOP');$target.StandardInput.Flush()
   $stopped=$target.StandardOutput.ReadLineAsync()
   if(-not $stopped.Wait(5000) -or $stopped.Result -ne 'STOPPED'){throw 'Normal helper STOP handshake was not observed.'}
   if(-not $target.WaitForExit(5000)){throw 'Normal helper did not exit; preserve without termination.'}
   $report.NormalHelperExitCode=$target.ExitCode
   if($target.ExitCode -ne 0){throw 'Normal helper exit is not zero.'}
  }elseif($mode -eq 'firstchanceav'){
   $report.SyntheticFirstChanceAvRequests=1
   $target.StandardInput.WriteLine('FIRSTCHANCEAV');$target.StandardInput.Flush()
   foreach($expected in @('FIRSTCHANCE_AV_REQUESTED','FIRSTCHANCE_AV_HANDLED')){
    $line=$target.StandardOutput.ReadLineAsync()
    if(-not $line.Wait(5000) -or $line.Result -ne $expected){throw 'Synthetic first-chance AV request/local-handler evidence missing.'}
   }
   if(-not $target.WaitForExit(5000) -or $target.ExitCode -ne 0 -or -not $session.Process.WaitForExit(5000)){throw 'Handled first-chance trial did not terminate normally.'}
   Stop-OwnedTeardownDebugger $session $target
   $text=Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode
   $report.FirstChanceAvForwardedVerified=Test-OwnedFirstChanceAv $text $target.Id $nonce
   if(-not $report.FirstChanceAvForwardedVerified){throw 'Executed exact first-chance AV forwarding was not observed.'}
  }else{
   $command=if($mode -eq 'failfast'){'FAILFAST'}else{'SECONDCHANCEAV'}
   $code=if($mode -eq 'failfast'){'c0000409'}else{'c0000005'}
   $report[$mode+'SyntheticRequests']=1
   if($mode -eq 'failfast'){$report.SyntheticFailfastRequests=1}else{$report.SyntheticSecondChanceAvRequests=1}
   $target.StandardInput.WriteLine($command);$target.StandardInput.Flush()
   if(-not $session.Process.WaitForExit(10000)){throw 'Failfast debugger handler did not complete.'}
   Stop-OwnedTeardownDebugger $session $target
   if(-not $target.WaitForExit(5000)){throw 'Synthetic failfast target did not exit; preserve without termination.'}
   $text=Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode
   $verified=Test-TeardownExceptionCapture $text $target.Id $code
   if($mode -eq 'failfast'){$report.ExceptionCaptureVerified=$verified}else{$report.SecondChanceAvCaptureVerified=$verified}
   if(-not $verified){throw 'Synthetic exact PID/code/live-register/stack handler was not verified.'}
  }
 }catch{$primary=$_.Exception.ToString();$report.Failure=$primary}
 finally{
  if($session -ne $null -and -not $session.Process.HasExited -and -not $session.StopAttempted){try{Stop-OwnedTeardownDebugger $session $target}catch{$cleanup.Add($_.Exception.ToString())}}
  try{
   if($session -ne $null -and $session.Process.HasExited){$report[$mode+'DebuggerExitCodeHex']='0x'+([uint32]([int64]$session.Process.ExitCode -band 4294967295)).ToString('X8')}
   if($target -ne $null){$report[$mode+'HelperExitObserved']=$target.HasExited;if($target.HasExited){$report[$mode+'HelperExitCodeHex']='0x'+([uint32]([int64]$target.ExitCode -band 4294967295)).ToString('X8')}}
  }catch{$cleanup.Add($_.Exception.ToString())}
  $report.CleanupFailures=@($cleanup.ToArray())
  Write-TeardownJson (Join-Path $trial 'preflight.json') $report
 }
 if($primary -ne $null -or $cleanup.Count -ne 0){break}
}
$report.State=if($primary -eq $null -and $cleanup.Count -eq 0 -and $report.ExceptionCaptureVerified -eq $true -and
 $report.NormalHelperExitCode -eq 0 -and $report.NormalDebuggerExitCode -eq 0 -and $report.NormalHelperDetachedVerified -eq $true -and $report.NormalStopBreakpointVerified -eq $true -and
 $report.FailfastHelperExitCodeHex -eq '0xC0000409' -and $report.failfastHelperExitObserved -eq $true -and
 $report.failfastDebuggerExitCodeHex -in @('0x00000000','0xC0000409') -and
 $report.SecondChanceAvCaptureVerified -eq $true -and $report.secondchanceavHelperExitObserved -eq $true -and
 $report.secondchanceavHelperExitCodeHex -eq '0xC0000005' -and $report.secondchanceavDebuggerExitCodeHex -in @('0x00000000','0xC0000005') -and
 $report.FirstChanceAvForwardedVerified -eq $true -and $report.firstchanceavHelperExitObserved -eq $true -and
 $report.firstchanceavHelperExitCodeHex -eq '0x00000000' -and $report.firstchanceavDebuggerExitCodeHex -eq '0x00000000' -and
 $report.SyntheticFailfastRequests -eq 1 -and $report.SyntheticSecondChanceAvRequests -eq 1 -and $report.SyntheticFirstChanceAvRequests -eq 1){'PASS'}else{'FAIL'}
$report.CompletedUtc=[DateTime]::UtcNow.ToString('o');Write-TeardownJson (Join-Path $trial 'preflight.json') $report
Write-Output (Join-Path $trial 'preflight.json')
if($report.State -ne 'PASS'){throw 'Owned helper preflight failed; do not attach to Office.'}
