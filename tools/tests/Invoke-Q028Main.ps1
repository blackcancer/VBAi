#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot,[Parameter(Mandatory=$true)][string]$ExpectedPlanSha256)
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module ScheduledTasks
$planPath=Join-Path $EvidenceRoot 'q028-plan.json'
if((Get-FileHash $planPath).Hash -cne $ExpectedPlanSha256){throw 'Frozen Q028 plan differs.'}
$plan=Get-Content $planPath -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $plan.MainDesktopAuthorized -or $plan.NativeDesktop -cne 'Default'){throw 'Explicit frozen Main Q028 scope required.'}
foreach($file in $plan.FrozenFiles){if((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.Sha256){throw ('Frozen Q028 input differs: '+$file.Path)}}
$output=Join-Path $EvidenceRoot 'launcher'
if(Test-Path $output){throw 'Fresh one-shot Q028 launcher required.'}
[void][IO.Directory]::CreateDirectory($output)
$taskName='VBAi-Q028Main-'+[Guid]::NewGuid().ToString('N')
$sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$intent=@{TaskName=$taskName;OwnerSid=$sid;Script=(Join-Path $EvidenceRoot 'Invoke-FrozenQ028.ps1');ScriptSha256=(Get-FileHash (Join-Path $EvidenceRoot 'Invoke-FrozenQ028.ps1')).Hash;Plan=$planPath;PlanSha256=$ExpectedPlanSha256;Desktop='WinSta0\Default';NoDesktopSwitch=$true;NoForceNativeHost=$true;Utc=[DateTime]::UtcNow.ToString('o')}
$intent|ConvertTo-Json|Set-Content (Join-Path $output 'intent.json') -Encoding UTF8
$body=@'
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
function Durable($name,$value){$s=[IO.File]::Open((Join-Path $PSScriptRoot $name),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read);try{$b=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $value -Depth 20));$s.Write($b,0,$b.Length);$s.Flush($true)}finally{$s.Dispose()}}
$p=Get-Content (Join-Path $PSScriptRoot 'intent.json') -Raw -Encoding UTF8|ConvertFrom-Json
if([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -cne $p.OwnerSid -or (Get-FileHash $p.Plan).Hash -cne $p.PlanSha256 -or (Get-FileHash $p.Script).Hash -cne $p.ScriptSha256){throw 'Original actor or frozen Q028 inputs changed.'}
if(-not [Environment]::Is64BitProcess){throw 'Original x64 launcher required.'}
$manifest=Get-Content $p.Plan -Raw -Encoding UTF8|ConvertFrom-Json
foreach($f in $manifest.FrozenFiles){if((Get-FileHash $f.Path).Hash -cne $f.Sha256){throw 'Frozen Q028 source changed.'}}
Add-Type -TypeDefinition @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Security.Principal;
public static class Q028MainProcess {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] struct Startup { internal uint Size; internal string Reserved,Desktop,Title; internal uint X,Y,Width,Height,XChars,YChars,Fill,Flags; internal ushort Show,ReservedSize; internal IntPtr ReservedPointer,Input,Output,Error; }
 [StructLayout(LayoutKind.Sequential)] public struct Child { public IntPtr Process,Thread; public uint Pid,Tid; }
 [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
 [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
 [DllImport("user32.dll")] static extern IntPtr GetThreadDesktop(uint tid);
 [DllImport("user32.dll")] static extern IntPtr GetProcessWindowStation();
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern bool GetUserObjectInformation(IntPtr h,int type,StringBuilder name,uint size,out uint required);
 [DllImport("advapi32.dll")] static extern bool OpenProcessToken(IntPtr h,uint access,out IntPtr token);
 [DllImport("advapi32.dll")] static extern bool GetTokenInformation(IntPtr token,int type,out int value,uint size,out uint required);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool CreateProcessW(string image,StringBuilder command,IntPtr pa,IntPtr ta,bool inherit,uint flags,IntPtr env,string directory,ref Startup startup,out Child child);
 [DllImport("kernel32.dll",SetLastError=true)] public static extern uint WaitForSingleObject(IntPtr handle,uint timeout);
 [DllImport("kernel32.dll",SetLastError=true)] public static extern bool GetExitCodeProcess(IntPtr handle,out uint code);
 [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
 static string Name(IntPtr handle) { uint n;var b=new StringBuilder(256);if(handle==IntPtr.Zero||!GetUserObjectInformation(handle,2,b,512,out n))throw new InvalidOperationException("Native desktop identity unavailable.");return b.ToString(); }
 public static void RequireActor() { if(Name(GetProcessWindowStation())!="WinSta0"||Name(GetThreadDesktop(GetCurrentThreadId()))!="Default")throw new InvalidOperationException("Exact Default launcher required."); IntPtr token;if(!OpenProcessToken(GetCurrentProcess(),8,out token))throw new InvalidOperationException("Token identity unavailable.");try { int elevated;uint n;if(!GetTokenInformation(token,20,out elevated,4,out n)||elevated!=0)throw new InvalidOperationException("Limited actor required."); }finally{CloseHandle(token);} }
 public static Child Start(string image,string file,string directory) { RequireActor();if(file.Contains("\"")||file.IndexOf((char)0)>=0)throw new ArgumentException("Invalid script path.");var startup=new Startup{Size=(uint)Marshal.SizeOf(typeof(Startup)),Desktop="WinSta0\\Default"};Child child;if(!CreateProcessW(image,new StringBuilder("\""+image+"\" -NoProfile -NonInteractive -STA -File \""+file+"\""),IntPtr.Zero,IntPtr.Zero,false,0x08000000,IntPtr.Zero,directory,ref startup,out child))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());return child; }
}
"@
[Q028MainProcess]::RequireActor()
Durable 'launcher-entry.json' @{ActorPid=$PID;OwnerSid=$p.OwnerSid;Desktop='Default';Attempts=1;Utc=[DateTime]::UtcNow.ToString('o')}
$child=$null;$observed=$false;$entered=$false;$code=[uint32]1;$failure=$null
try{
 $entered=$true;$child=[Q028MainProcess]::Start((Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'),$p.Script,(Split-Path $p.Script))
 Durable 'worker-original-create.json' @{ProcessId=$child.Pid;ThreadId=$child.Tid;OriginalHandle=$child.Process.ToInt64();HandleSource='OriginalCreateProcessHandle'}
 $process=Get-Process -Id $child.Pid
 Durable 'worker-original-process.json' @{ProcessId=$child.Pid;ThreadId=$child.Tid;OriginalHandle=$child.Process.ToInt64();HandleSource='OriginalCreateProcessHandle';StartUtc=$process.StartTime.ToUniversalTime().ToString('o');Image=$process.Path}
 $process.Dispose()
 $wait=[Q028MainProcess]::WaitForSingleObject($child.Process,[uint32]::MaxValue)
 if($wait -ne 0 -or -not [Q028MainProcess]::GetExitCodeProcess($child.Process,[ref]$code)){throw 'Original Q028 worker exit unavailable.'}
 $observed=$true
 Durable 'worker-original-exit.json' @{ProcessId=$child.Pid;OriginalHandle=$child.Process.ToInt64();ExitCode=$code;ExitObserved=$true;Forced=$false;Utc=[DateTime]::UtcNow.ToString('o')}
 Durable 'terminal.json' @{State='ORIGINAL_MAIN_WORKER_EXIT_OBSERVED';ExitCode=$code;OwnerSid=$p.OwnerSid;Desktop='Default';NoDesktopSwitch=$true;Forced=$false;Utc=[DateTime]::UtcNow.ToString('o')}
}catch{
 $failure=$_.Exception.ToString();Durable 'launcher-failure.json' @{Error=$failure;StartEntered=$entered;OriginalExitObserved=$observed;NoRetry=$true;NoKill=$true}
 if($child -and -not $observed){$wait=[Q028MainProcess]::WaitForSingleObject($child.Process,[uint32]::MaxValue);if($wait -eq 0 -and [Q028MainProcess]::GetExitCodeProcess($child.Process,[ref]$code)){$observed=$true;Durable 'failed-original-worker-exit.json' @{ProcessId=$child.Pid;OriginalHandle=$child.Process.ToInt64();ExitObserved=$true;ExitCode=$code;Forced=$false}}}
 if($entered -and -not $observed){[Threading.ManualResetEvent]::new($false).WaitOne()|Out-Null}
}finally{if($child -and $observed){[Q028MainProcess]::CloseHandle($child.Thread)|Out-Null;[Q028MainProcess]::CloseHandle($child.Process)|Out-Null}}
if($failure){exit 1};exit $code
'@
$launcher=Join-Path $output 'Launch.ps1'
[IO.File]::WriteAllText($launcher,$body,[Text.UTF8Encoding]::new($false))
$action=New-ScheduledTaskAction -Execute (Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe') -Argument ('-NoProfile -NonInteractive -STA -WindowStyle Hidden -File "'+$launcher+'"') -WorkingDirectory $output
$principal=New-ScheduledTaskPrincipal -UserId $sid -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -DisallowHardTerminate -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Description 'Q028 frozen real-desktop Office batch; owned hosts only, no VBA execution'|Out-Null
Start-ScheduledTask -TaskName $taskName
@{State='DISPATCHED_ONCE';TaskName=$taskName;EvidenceRoot=$EvidenceRoot}|ConvertTo-Json -Compress
