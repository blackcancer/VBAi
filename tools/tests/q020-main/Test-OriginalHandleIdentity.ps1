#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
$gate=Join-Path $PSScriptRoot 'Q020MainDesktopGate.cs'
$runner=Join-Path $PSScriptRoot 'Invoke-Q020MainNativeMacroQualification.ps1'
if(-not [IO.Path]::IsPathRooted($OutputRoot)){throw 'Absolute fresh OutputRoot required.'}
$output=[IO.Path]::GetFullPath($OutputRoot)
if(Test-Path -LiteralPath $output){throw 'Fresh one-shot output required.'}
[IO.Directory]::CreateDirectory($output)|Out-Null
function Durable([string]$path,$value){
 $bytes=[Text.Encoding]::UTF8.GetBytes(($value|ConvertTo-Json -Depth 20))
 $f=[IO.File]::Open($path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
 try{$f.Write($bytes,0,$bytes.Length);$f.Flush($true)}finally{$f.Dispose()}
}
$gt=[IO.File]::ReadAllText($gate)
if($gt.Contains('MainModule') -or $gt.Contains('StartTime')){throw 'Gate must not use module-enumeration/startup-managed identity.'}
$t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($runner,[ref]$t,[ref]$e)
if($e.Count){throw 'Runner AST invalid.'}
$launch=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Launch-Owned'},$true))[0].Extent.Text
if($launch.Contains('.StartTime') -or $launch.Contains('.Path') -or $launch.Contains('.MainModule')){throw 'Launch-Owned still depends on managed startup metadata.'}
if($launch.IndexOf('$script:swChild=') -ge $launch.IndexOf("host-original-create.json") -or $launch.IndexOf("host-original-create.json") -ge $launch.IndexOf("GetMethod('ValidateSolidWorksChild'")){throw 'Caller child/raw receipt must precede identity admission.'}
$nativeLaunch=[regex]::Match($gt,'(?s)internal static NativeChild LaunchSolidWorks.*?(?=internal static void ValidateSolidWorksChild)').Value
if(-not $nativeLaunch -or $nativeLaunch.Contains('ReadIdentity(') -or $nativeLaunch.Contains('Remember(') -or $nativeLaunch.Contains('ValidateIdentity(')){throw 'Native launch performs identity operations before returning ownership.'}
Add-Type -Path $gate
$factory=@'
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
public static class Q020IdentityCanaryFactory {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] struct Startup {
  internal uint Size; internal string Reserved,Desktop,Title;
  internal uint X,Y,Width,Height,XChars,YChars,Fill,Flags;
  internal ushort Show,ReservedSize; internal IntPtr ReservedPointer,Input,Output,Error;
 }
 [StructLayout(LayoutKind.Sequential)] struct Child {internal IntPtr Process,Thread;internal uint Pid,Tid;}
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool CreateProcessW(string exe,StringBuilder command,IntPtr pa,IntPtr ta,bool inherit,uint flags,IntPtr env,string directory,ref Startup startup,out Child child);
 public static int CreationCount;
 public static object Create(string executable,string directory,Type actualChildType) {
  var startup=new Startup{Size=(uint)Marshal.SizeOf(typeof(Startup)),Desktop="WinSta0\\Default"};Child child;
  ++CreationCount;
  string command="\""+executable+"\" -NoLogo -NoProfile -NonInteractive -Command \"Start-Sleep -Seconds 3; exit 23\"";
  if(!CreateProcessW(executable,new StringBuilder(command),IntPtr.Zero,IntPtr.Zero,false,0x08000000,IntPtr.Zero,directory,ref startup,out child))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
  // The fixture constructs the actual unchanged NativeChild around these original kernel handles.
  return actualChildType.GetConstructor(BindingFlags.Instance|BindingFlags.NonPublic,null,new[]{typeof(IntPtr),typeof(IntPtr),typeof(int),typeof(uint)},null).Invoke(new object[]{child.Process,child.Thread,(int)child.Pid,child.Tid});
 }
}
'@
Add-Type -TypeDefinition $factory
$type=[Q020MainDesktopGate].GetNestedType('NativeChild',[Reflection.BindingFlags]'NonPublic')
$flags=[Reflection.BindingFlags]'Instance,NonPublic'
$exe=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
$exeHash=(Get-FileHash -LiteralPath $exe).Hash
$cases=@()
foreach($entry in @(@{Name='Wrong expected identity after durable raw receipt';Wrong=$true},@{Name='Immediate SDK identity before module enumeration';Wrong=$false})){
 $id=($cases.Count+1).ToString('D2')
 $child=[Q020IdentityCanaryFactory]::Create($exe,$output,$type)
 $handle=$type.GetProperty('ProcessHandle',$flags).GetValue($child,$null)
 $pidValue=$type.GetProperty('ProcessId',$flags).GetValue($child,$null)
 $tid=$type.GetProperty('ThreadId',$flags).GetValue($child,$null)
 $threadHandle=$type.GetProperty('ThreadHandle',$flags).GetValue($child,$null)
 $rawPath=Join-Path $output ($id+'-raw-original-create.json')
 Durable $rawPath @{OwnerPid=$PID;ProcessId=$pidValue;ThreadId=$tid;OriginalHandle=$handle.ToInt64();OriginalThreadHandle=$threadHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';IdentityVerified=$false;Utc=[DateTime]::UtcNow.ToString('o')}
 $rawHash=(Get-FileHash -LiteralPath $rawPath).Hash
 $expected=$(if($entry.Wrong){Join-Path $output 'deliberately-different-powershell.exe'}else{$exe})
 $identityError=$null;$before=[DateTime]::UtcNow.ToString('o');$verified=$false
 try{$type.GetMethod('ValidateIdentity',$flags).Invoke($child,@([string]$expected))|Out-Null}
 catch{$cause=$_.Exception;while($cause.InnerException){$cause=$cause.InnerException};$identityError=$cause.Message}
 $verified=$type.GetProperty('IdentityVerified',$flags).GetValue($child,$null)
 $attempted=$type.GetProperty('IdentityAttempted',$flags).GetValue($child,$null)
 $image=$type.GetProperty('ImagePath',$flags).GetValue($child,$null)
 $birth=$type.GetProperty('BirthUtc',$flags).GetValue($child,$null)
 # Observe only this original retained handle; natural child exit, no Quit/force/second launch.
 $exited=[bool]$type.GetMethod('Wait',$flags).Invoke($child,@([int]10000))
 if(-not $exited){Durable (Join-Path $output ($id+'-retained.json')) @{State='UNCERTAIN_RETAINED';OriginalHandle=$handle.ToInt64();Pid=$pidValue};throw 'Canary child not settled; original handle retained.'}
 $code=$type.GetMethod('ExitCode',$flags).Invoke($child,@())
 Durable (Join-Path $output ($id+'-original-exit.json')) @{Pid=$pidValue;Handle=$handle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;Utc=[DateTime]::UtcNow.ToString('o')}
 $sameHandle=($type.GetProperty('ProcessHandle',$flags).GetValue($child,$null) -eq $handle)
 $child.Dispose()
 if(-not $sameHandle -or -not $attempted -or $image -ine $exe -or -not $birth -or $code -ne 23 -or $verified -eq $entry.Wrong -or ($entry.Wrong -and $identityError -cne 'Original native executable identity differs.') -or (-not $entry.Wrong -and $identityError)){throw ('Canary native identity assertion failed: '+$entry.Name+' '+$identityError)}
 $cases+=@{Name=$entry.Name;Pass=$true;RawReceipt=$rawPath;RawReceiptSha256=$rawHash;ValidationEnteredUtc=$before;IdentityAttempted=$attempted;IdentityVerified=$verified;ObservedImage=$image;ObservedBirthUtc=$birth;Error=$identityError;SameOriginalHandleAtExit=$sameHandle;ExitCode=$code;CreationCount=$id}
}
if([Q020IdentityCanaryFactory]::CreationCount -ne 2){throw 'Exactly one creation per canary case required.'}
$receipt=@{State='OWNED_POWERSHELL_ORIGINAL_HANDLE_CANARY_PASS';Cases=$cases;CreationCount=2;AllClaimsSettled=$true;NativeAcceptance=$false;Scope='Two disposable hidden PowerShell children only; actual NativeChild ValidateIdentity via original CreateProcess handles before any managed module read. Deliberate expected-image refusal retains original child and observes natural exit; no SOLIDWORKS, UI, COM, registry, trust, forced exit or launch retry.';HelperSha256=(Get-FileHash -LiteralPath $gate).Hash;RunnerSha256=(Get-FileHash -LiteralPath $runner).Hash;Executable=$exe;ExecutableSha256=$exeHash;RawBeforeValidationOrderVerified=$true}
Durable (Join-Path $output 'original-handle-canary.json') $receipt
'PASS two owned hidden PowerShell original-handle canaries; SDK image/time identity and failure retention; no SOLIDWORKS/native UI/registry.'
