[CmdletBinding()]
param(
 [Parameter(Mandatory)][string]$HelperPath,
 [Parameter(Mandatory)][string]$OutputRoot,
 [string]$CdbPath='D:\Windows Kits\10\Debuggers\x64\cdb.exe'
)
$ErrorActionPreference='Stop'
if(-not [IO.Path]::IsPathRooted($HelperPath) -or -not [IO.Path]::IsPathRooted($OutputRoot) -or
 [IO.Path]::GetFileName($HelperPath) -ne 'VBAi.Tests.exe' -or (Test-Path -LiteralPath $OutputRoot)) {throw 'A frozen helper and fresh absolute output directory are required.'}
$helper=Resolve-Path -LiteralPath $HelperPath
$debugger=Resolve-Path -LiteralPath $CdbPath
[IO.Directory]::CreateDirectory($OutputRoot) | Out-Null
$report=[ordered]@{State='PREPARING';Scope='Synthetic CLR guard only; no Office, COM activation, native preference or historical-cause proof';
 HelperPath=$helper.Path;HelperSha256=(Get-FileHash -LiteralPath $helper.Path).Hash;
 CdbPath=$debugger.Path;CdbSha256=(Get-FileHash -LiteralPath $debugger.Path).Hash;
 NativePreferenceWrites=0;ForcedTerminations=0}
function Checkpoint{$report|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $OutputRoot 'preflight.json') -Encoding UTF8}
function Start-Owned($path,$arguments){
 $info=[Diagnostics.ProcessStartInfo]::new($path,$arguments);$info.UseShellExecute=$false;$info.CreateNoWindow=$true
 $info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
 $process=[Diagnostics.Process]::new();$process.StartInfo=$info;[void]$process.Start();return $process
}
function Read-OwnedLine($process){$pending=$process.StandardOutput.ReadLineAsync();if(-not $pending.Wait(10000)){throw 'Owned helper response timeout; process retained.'};return $pending.Result}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Q026ClrTraceNative {
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool DebugBreakProcess(IntPtr process);
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CheckRemoteDebuggerPresent(IntPtr process,out bool present);
}
'@
$target=$null;$cdb=$null;$primary=$null;$cleanup=$null
try{
 $target=Start-Owned $helper.Path ''
 $ready=Read-OwnedLine $target
 if($ready -notmatch '^READY (\d+) (.+)$' -or [int]$Matches[1] -ne $target.Id -or
  [DateTime]::Parse($Matches[2]).ToUniversalTime() -ne $target.StartTime.ToUniversalTime()){throw 'Owned helper handshake differs.'}
 $report.TargetProcessId=$target.Id;$report.TargetStartUtc=$target.StartTime.ToUniversalTime().ToString('o');Checkpoint
 $expected=Read-OwnedLine $target
 if($expected -notmatch '^EXPECTED (.+)$'){throw 'Synthetic expected snapshot is missing.'}
 $expectedJson=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Matches[1]))
 [IO.File]::WriteAllText((Join-Path $OutputRoot 'expected-tabs.json'),$expectedJson,[Text.Encoding]::UTF8)
 $commands=Join-Path $OutputRoot 'commands.txt';$log=Join-Path $OutputRoot 'trace.cdb.log'
 $script=Join-Path $PSScriptRoot 'Q026OptionsGuardTrace.js'
 $provider=Join-Path ([IO.Path]::GetDirectoryName($debugger.Path)) 'winext/JsProvider.dll'
 if(-not (Test-Path -LiteralPath $script) -or -not (Test-Path -LiteralPath $provider)){throw 'Installed CLR trace script/provider is unavailable.'}
 $report.TraceScriptSha256=(Get-FileHash -LiteralPath $script).Hash
 $text='.logopen /u "'+$log.Replace('\','/')+'"'+"`n.loadby sos clr`n.load `""+$provider.Replace('\','/')+"`"`n.scriptload `""+$script.Replace('\','/')+"`"`nsxe -c `".echo Q026_CLR_EXCEPTION;!pe;!clrstack -a;dx @`$scriptContents.capture();gn`" clr`n.echo Q026_CLR_TRACE_ARMED`ng`n"
 [IO.File]::WriteAllText($commands,$text,[Text.Encoding]::GetEncoding(1252))
 $cdb=Start-Owned $debugger.Path ('-pd -p '+$target.Id+' -netsyms:no -cf "'+$commands+'"')
 $report.DebuggerProcessId=$cdb.Id;$out=$cdb.StandardOutput.ReadToEndAsync();$err=$cdb.StandardError.ReadToEndAsync()
 $deadline=[Diagnostics.Stopwatch]::StartNew()
 while($deadline.Elapsed.TotalSeconds -lt 10 -and -not $cdb.HasExited){
  if((Test-Path -LiteralPath $log) -and (Get-Content -LiteralPath $log -Raw -Encoding Unicode) -match 'Q026_CLR_TRACE_ARMED'){break}
  Start-Sleep -Milliseconds 100
 }
 $present=$false
 if(-not (Test-Path -LiteralPath $log) -or (Get-Content -LiteralPath $log -Raw -Encoding Unicode) -notmatch 'Q026_CLR_TRACE_ARMED' -or
  -not [Q026ClrTraceNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -or -not $present){throw 'CLR trace did not arm; no synthetic request sent.'}
 $report.AttachmentObserved=$true;$report.SyntheticRequests=1;Checkpoint
 $target.StandardInput.WriteLine('OBSERVE');$target.StandardInput.Flush()
 if((Read-OwnedLine $target) -ne 'GUARD_REJECTED'){throw 'Synthetic guard did not refuse.'}
 $captured=Get-Content -LiteralPath $log -Raw -Encoding Unicode
 $report.GuardMessageCaptured=$captured -match 'VBE options changed since inspection'
 $report.GuardStackCaptured=$captured -match 'SetVbeOption'
 $report.GuardLocalsCaptured=$captured -match 'LOCALS:'
 $report.GuardSnapshotCaptured=$captured -match 'Q026_GUARD_END '
 if(-not $report.GuardMessageCaptured -or -not $report.GuardStackCaptured -or -not $report.GuardLocalsCaptured -or -not $report.GuardSnapshotCaptured){throw 'CLR guard message/stack/locals/snapshot capture is incomplete.'}
 $lines=$captured -split "`r?`n"
 $begins=@($lines|Where-Object {$_ -match '^Q026_GUARD_BEGIN '}|ForEach-Object {$_.Substring('Q026_GUARD_BEGIN '.Length)|ConvertFrom-Json})
 $ends=@($lines|Where-Object {$_ -match '^Q026_GUARD_END '}|ForEach-Object {$_.Substring('Q026_GUARD_END '.Length)|ConvertFrom-Json})
 $chunks=@($lines|Where-Object {$_ -match '^Q026_GUARD_CHUNK '}|ForEach-Object {$_.Substring('Q026_GUARD_CHUNK '.Length)|ConvertFrom-Json})
 if($begins.Count -ne 1 -or $ends.Count -ne 1 -or $begins[0].Chunks -ne $chunks.Count -or
  ($begins[0]|ConvertTo-Json -Compress) -cne ($ends[0]|ConvertTo-Json -Compress)){throw 'The synthetic guard frames are missing or inconsistent.'}
 $text=[Text.StringBuilder]::new()
 for($index=0;$index -lt $chunks.Count;$index++){
  if($chunks[$index].Capture -ne $begins[0].Capture -or $chunks[$index].Index -ne $index -or $chunks[$index].Text.Length -gt 2000){throw 'Guard chunk identity/order/bound differs.'}
  [void]$text.Append($chunks[$index].Text)
 }
 if($text.Length -ne $begins[0].Length){throw 'The guard capture is incomplete.'}
 $state=$text.ToString()|ConvertFrom-Json
 $report.FullSyntheticTabsEqual=($state.Tabs|ConvertTo-Json -Depth 30 -Compress) -ceq ($expectedJson|ConvertFrom-Json|ConvertTo-Json -Depth 30 -Compress)
 $report.ExpectedVersionVerified=$state.Request.ExpectedOptionsVersion -ceq ('0'*64)
 $report.GuardValueVerified=$state.Request.Value -ceq 'Synthetic alternate'
 $state|ConvertTo-Json -Depth 30|Set-Content -LiteralPath (Join-Path $OutputRoot 'captured-guard.json') -Encoding UTF8
 if(-not $report.FullSyntheticTabsEqual -or -not $report.ExpectedVersionVerified -or -not $report.GuardValueVerified){throw 'Captured CLR preferences/request differ from the exact synthetic baseline.'}
 $report.State='CAPTURED';Checkpoint
}catch{$primary=$_;$report.State='FAILED';$report.Failure=$_.Exception.ToString()}
finally{
 if($cdb -and -not $cdb.HasExited){
  try{
   $cdb.StandardInput.WriteLine('sxd clr');$cdb.StandardInput.WriteLine('bc *');$cdb.StandardInput.WriteLine('.logclose');$cdb.StandardInput.WriteLine('qd');$cdb.StandardInput.Flush()
   if(-not [Q026ClrTraceNative]::DebugBreakProcess($target.Handle)){throw 'Owned debugger break failed; target retained.'}
   if(-not $cdb.WaitForExit(10000)){throw 'Debugger did not detach; no target shutdown attempted.'}
  }catch{$cleanup=$_.Exception.ToString();$report.DetachError=$cleanup}
 }
 if($cdb -and $cdb.HasExited){
  $report.DebuggerExitCode=$cdb.ExitCode
  [IO.File]::WriteAllText((Join-Path $OutputRoot 'cdb.stdout.log'),$out.GetAwaiter().GetResult())
  [IO.File]::WriteAllText((Join-Path $OutputRoot 'cdb.stderr.log'),$err.GetAwaiter().GetResult())
 }
 if($target -and -not $target.HasExited){
  $present=$false
  $report.DebuggerDetachedVerified=[Q026ClrTraceNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -and -not $present
  if($report.DebuggerDetachedVerified){
   $target.StandardInput.WriteLine('QUIT');$target.StandardInput.Flush()
   $report.OriginalHandleExitObserved=$target.WaitForExit(10000)
   if($report.OriginalHandleExitObserved){$report.TargetExitCode=$target.ExitCode}
  }
 }
 if(-not $primary -and -not $cleanup -and $report.DebuggerDetachedVerified -and $report.OriginalHandleExitObserved -and $report.TargetExitCode -eq 0 -and $report.DebuggerExitCode -eq 0){$report.State='SYNTHETIC_PREFLIGHT_VERIFIED'}
 Checkpoint
}
if($report.State -ne 'SYNTHETIC_PREFLIGHT_VERIFIED'){throw 'CLR preflight failed. Read the retained report; no native attachment is authorized by this result.'}
Write-Output $report.State
