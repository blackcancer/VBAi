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
 CdbPath=$CdbPath;CdbSha256=(Get-FileHash -LiteralPath $CdbPath).Hash;CommonScriptSha256=(Get-FileHash -LiteralPath $common).Hash;
 Scope='Two newly owned non-Office helpers only: normal attach/detach/STOP, then one synthetic failfast. No Office, COM, dump or global policy.'}
if(-not $Execute){$report|ConvertTo-Json -Depth 4;return}
. $common
$trial=Join-Path $OutputRoot ([Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($trial)
$primary=$null;$cleanup=New-Object Collections.Generic.List[string];$report.State='STARTED'
foreach($mode in @('normal','failfast')){
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
   $target.StandardInput.WriteLine('STOP');$target.StandardInput.Flush()
   $stopped=$target.StandardOutput.ReadLineAsync()
   if(-not $stopped.Wait(5000) -or $stopped.Result -ne 'STOPPED'){throw 'Normal helper STOP handshake was not observed.'}
   if(-not $target.WaitForExit(5000)){throw 'Normal helper did not exit; preserve without termination.'}
   $report.NormalHelperExitCode=$target.ExitCode
   if($target.ExitCode -ne 0){throw 'Normal helper exit is not zero.'}
  }else{
   $target.StandardInput.WriteLine('FAILFAST');$target.StandardInput.Flush();$report.SyntheticFailfastRequests=1
   if(-not $session.Process.WaitForExit(10000)){throw 'Failfast debugger handler did not complete.'}
   Stop-OwnedTeardownDebugger $session $target
   if(-not $target.WaitForExit(5000)){throw 'Synthetic failfast target did not exit; preserve without termination.'}
   $report.FailfastHelperExitCodeHex='0x'+([uint32]([int64]$target.ExitCode -band 4294967295)).ToString('X8')
   $text=Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode
   $report.ExceptionCaptureVerified=Test-TeardownExceptionCapture $text
   if(-not $report.ExceptionCaptureVerified){throw 'Synthetic exception/context/stack handler was not verified.'}
  }
 }catch{$primary=$_.Exception.ToString();$report.Failure=$primary}
 finally{
  if($session -ne $null -and -not $session.Process.HasExited){try{Stop-OwnedTeardownDebugger $session $target}catch{$cleanup.Add($_.Exception.ToString())}}
  $report.CleanupFailures=@($cleanup.ToArray())
  Write-TeardownJson (Join-Path $trial 'preflight.json') $report
 }
 if($primary -ne $null -or $cleanup.Count -ne 0){break}
}
$report.State=if($primary -eq $null -and $cleanup.Count -eq 0 -and $report.ExceptionCaptureVerified -eq $true -and
 $report.NormalHelperExitCode -eq 0 -and $report.NormalHelperDetachedVerified -eq $true -and
 $report.FailfastHelperExitCodeHex -eq '0xC0000409'){'PASS'}else{'FAIL'}
$report.CompletedUtc=[DateTime]::UtcNow.ToString('o');Write-TeardownJson (Join-Path $trial 'preflight.json') $report
Write-Output (Join-Path $trial 'preflight.json')
if($report.State -ne 'PASS'){throw 'Owned helper preflight failed; do not attach to Office.'}
