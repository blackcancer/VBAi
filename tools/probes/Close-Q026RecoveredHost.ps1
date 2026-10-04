#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RecoveryRoot,
    [Parameter(Mandatory=$true)][string]$OutputPath,[string]$BootstrapClosureObservation,[string]$IsolationDesktopRoot)
$ErrorActionPreference='Stop'
if(-not [IO.Path]::IsPathRooted($OutputPath) -or (Test-Path -LiteralPath $OutputPath) -or (Test-Path -LiteralPath ($OutputPath+'.progress.json'))){throw 'Fresh one-shot shutdown output required.'}
$root=@(Get-ChildItem -LiteralPath (Join-Path $EvidenceRoot 'native/hosts') -Directory)
if($root.Count -ne 1){throw 'One owned host required.'}
$startup=Get-Content -LiteralPath (Join-Path $root[0].FullName 'startup.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $IsolationDesktopRoot){$IsolationDesktopRoot=Join-Path $EvidenceRoot 'isolation/desktop'}
$desktop=(Get-Content -LiteralPath (Join-Path $IsolationDesktopRoot 'desktop-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json).Desktop
$worker=Get-Content -LiteralPath (Join-Path $IsolationDesktopRoot 'worker-desktop.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if($worker.ActualDesktop -ne $desktop -or $worker.ExpectedDesktop -ne $desktop -or $worker.DesktopSwitches -ne 0){throw 'Private worker desktop evidence mismatch.'}
if(-not $startup.Owned){throw 'Owned launch required.'}
if($BootstrapClosureObservation){
    if($RecoveryRoot){throw 'Bootstrap-only closure cannot stand in for preference recovery.'}
    $boot=Get-Content -LiteralPath (Join-Path $root[0].FullName 'owned-bootstrap.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $campaign=Get-Content -LiteralPath (Join-Path $EvidenceRoot 'campaign.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $worker=Get-Content -LiteralPath (Join-Path $EvidenceRoot 'isolation/desktop/campaign-exit.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $closed=Get-Content -LiteralPath $BootstrapClosureObservation -Raw -Encoding UTF8 | ConvertFrom-Json
    $preferenceReceipts=@(Get-ChildItem -LiteralPath (Join-Path $EvidenceRoot 'native/phases') -Recurse -Filter 'options-*.json')
    if($boot.Phase -ne 'FailedPreserved' -or $boot.FailedAtPhase -ne 'ExactNativeApplicationAndSeedAttached' -or
        $boot.ProcessId -ne $startup.ProcessId -or $boot.ProcessStartUtc -ne $startup.HostStartedUtc -or
        $boot.BootstrapCloseAttempts -ne 0 -or $boot.BootstrapQuitAttempts -ne 0 -or $boot.ForceTerminationAttempts -ne 0 -or
        $boot.LoadedAssemblyMvid -or $preferenceReceipts.Count -ne 0 -or
        $campaign.NativeState -ne 'FAILED_OR_SKIPPED' -or $campaign.NativeExitCode -ne 1 -or $worker.ExitCode -ne 1 -or
        $closed.ProcessId -ne $startup.ProcessId -or $closed.ProcessStartUtc -ne $startup.HostStartedUtc -or
        $closed.Desktop -ne $desktop -or -not $closed.ObservationOnly -or -not $closed.EnumerationSucceeded -or -not $closed.OptionsDialogAbsent){
        throw 'Terminal bootstrap failure before all preference dispatch and independent dialog absence required.'
    }
    $seedStream=[IO.FileStream]::new($boot.Seed,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $sha=[Security.Cryptography.SHA256]::Create()
    try {$seedHash=[BitConverter]::ToString($sha.ComputeHash($seedStream)).Replace('-','')}
    finally {$sha.Dispose();$seedStream.Dispose()}
    if($seedHash -ine $boot.SeedSha256){throw 'Disposable bootstrap seed changed.'}
} else {
    if(-not $RecoveryRoot){throw 'Settled preference recovery required.'}
    $recovered=Get-Content -LiteralPath (Join-Path $RecoveryRoot 'terminal.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if($recovered.State -ne 'PREFERENCES_RECOVERED' -or $recovered.Qualified -ne $false -or $recovered.HostShutdownInvoked){throw 'Settled complete recovery required; this does not qualify the failed native test.'}
    if($recovered.ProcessId -ne $startup.ProcessId -or $recovered.ProcessStartUtc -ne $startup.HostStartedUtc){throw 'Recovered launch identity mismatch.'}
}
Add-Type -ReferencedAssemblies System.dll,System.Core.dll,System.Web.Extensions.dll -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
public static class Q026RecoveryShutdown {
 static readonly List<object> retainedReferences=new List<object>();
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr OpenDesktopW(string n,uint f,bool i,uint a);
 [DllImport("user32.dll",SetLastError=true)] static extern bool SetThreadDesktop(IntPtr d);
 [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr d);
 delegate bool Visitor(IntPtr h,IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumWindows(Visitor v,IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h,Visitor v,IntPtr p);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h,StringBuilder s,int c);
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateWindowExW(uint e,string c,string t,uint s,int x,int y,int w,int h,IntPtr p,IntPtr m,IntPtr i,IntPtr a);
 [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr h);
 [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h,uint id,ref Guid iid,[MarshalAs(UnmanagedType.Interface)]out object value);
 static object Get(object o,string name,params object[] args){return o.GetType().InvokeMember(name,BindingFlags.GetProperty|BindingFlags.IgnoreCase,null,o,args);}
 static object Call(object o,string name,params object[] args){return o.GetType().InvokeMember(name,BindingFlags.InvokeMethod|BindingFlags.IgnoreCase,null,o,args);}
 static void Release(object o){if(o!=null && Marshal.IsComObject(o))Marshal.FinalReleaseComObject(o);}
 public static void Run(string desktop,int pid,string start,string seed,string executable,string output,bool bootstrapOnly) {
  var state=new Dictionary<string,object>{{"ProcessId",pid},{"ProcessStartUtc",start},{"Desktop",desktop},{"Qualified",false},{"ForcedTermination",false},{"CloseEntries",0},{"QuitEntries",0},{"State","IDENTITY_GUARD"}};
  Action persist=()=>File.WriteAllText(output+".progress.json",new JavaScriptSerializer().Serialize(state),new UTF8Encoding(false));
  state["RecoveryKind"]=bootstrapOnly?"BootstrapFailureNoPreferenceDispatch":"CompletePreferenceRecovery";
  persist();Exception failure=null;int stopDispatch=0;
  Action requireDispatch=()=>{if(Interlocked.CompareExchange(ref stopDispatch,0,0)!=0)throw new InvalidOperationException("Shutdown deadline expired; no further native action.");};
  var process=Process.GetProcessById(pid);IntPtr handle=process.Handle;
  if(process.StartTime.ToUniversalTime()!=DateTime.Parse(start).ToUniversalTime() || !string.Equals(process.MainModule.FileName,executable,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Owned process changed before recovery shutdown.");
  state["RecoveryProcessHandle"]=handle.ToInt64();state["OriginalCampaignHandleProof"]=false;persist();
  var thread=new Thread(()=> {
   IntPtr desk=OpenDesktopW(desktop,0,false,0xC7),reader=IntPtr.Zero;
   object window=null,app=null,books=null,book=null;
   try {
    if(desk==IntPtr.Zero || !SetThreadDesktop(desk))throw new Win32Exception(Marshal.GetLastWin32Error());
    reader=CreateWindowExW(0,"STATIC","Q026 recovered host shutdown",0,0,0,1,1,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
    if(reader==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());
    var documents=new List<IntPtr>();bool dialog=false;
    if(!EnumWindows((h,p)=> {
      uint owner;GetWindowThreadProcessId(h,out owner);if(owner!=pid)return true;
      var cls=new StringBuilder(256);GetClassNameW(h,cls,cls.Capacity);if(cls.ToString()=="#32770")dialog=true;
      EnumChildWindows(h,(child,unused)=> {GetWindowThreadProcessId(child,out owner);cls.Clear();GetClassNameW(child,cls,cls.Capacity);if(owner==pid && cls.ToString()=="EXCEL7")documents.Add(child);return true;},IntPtr.Zero);return true;
    },IntPtr.Zero) || dialog || documents.Count!=1)throw new InvalidOperationException("One exact owned document and no modal dialog required; no Close/Quit.");
    Guid dispatch=new Guid("00020400-0000-0000-C000-000000000046");
    if(AccessibleObjectFromWindow(documents[0],0xFFFFFFF0,ref dispatch,out window)!=0 || window==null)throw new InvalidOperationException("Owned native document attachment failed; no COM activation fallback.");
    app=Get(window,"Application");uint applicationPid;GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(Get(app,"Hwnd"))),out applicationPid);
    if(applicationPid!=pid || process.HasExited || process.StartTime.ToUniversalTime()!=DateTime.Parse(start).ToUniversalTime())throw new InvalidOperationException("Attached application owner changed.");
    books=Get(app,"Workbooks");if(Convert.ToInt32(Get(books,"Count"))!=1)throw new InvalidOperationException("Foreign document present; no shutdown.");
    book=Get(books,"Item",1);
    if(!string.Equals(Path.GetFullPath(Convert.ToString(Get(book,"FullName"))),Path.GetFullPath(seed),StringComparison.OrdinalIgnoreCase) || !Convert.ToBoolean(Get(book,"Saved")))throw new InvalidOperationException("Owned saved seed identity changed; no shutdown.");
    requireDispatch();state["State"]="CLOSE_INTENT";state["CloseEntries"]=1;persist();
    Call(book,"Close",false,Type.Missing,Type.Missing);state["CloseReturned"]=true;
    if(Convert.ToInt32(Get(books,"Count"))!=0)throw new InvalidOperationException("Owned seed close unverified; no Quit.");
    requireDispatch();state["State"]="QUIT_INTENT";state["QuitEntries"]=1;persist();
    Call(app,"Quit");state["QuitReturned"]=true;state["State"]="QUIT_RETURNED";persist();
   } catch(Exception e){failure=e;} finally {
    if(failure!=null && Convert.ToInt32(state["CloseEntries"])!=0 && !state.ContainsKey("QuitReturned")){
      retainedReferences.AddRange(new[]{book,books,app,window});state["State"]="RETAINED_NATIVE_OUTCOME";state["Error"]=failure.ToString();persist();
      for(;;)Thread.Sleep(1000);
    }
    try{Release(book);Release(books);Release(app);Release(window);}catch(Exception e){if(failure==null)failure=e;}
    if(reader!=IntPtr.Zero)DestroyWindow(reader);if(desk!=IntPtr.Zero)CloseDesktop(desk);
   }
  });thread.SetApartmentState(ApartmentState.STA);thread.IsBackground=true;thread.Start();
  if(!thread.Join(30000)){Interlocked.Exchange(ref stopDispatch,1);state["State"]="RETAINED_PENDING_COM";persist();for(;;)Thread.Sleep(1000);}
  if(failure!=null){state["State"]="RETAINED_SHUTDOWN_ERROR";state["Error"]=failure.ToString();persist();for(;;)Thread.Sleep(1000);}
  if(!process.WaitForExit(15000)){state["State"]="RETAINED_EXIT_NOT_OBSERVED";persist();for(;;)Thread.Sleep(1000);}
  state["ExitCode"]=process.ExitCode;state["State"]=process.ExitCode==0?(bootstrapOnly?"BOOTSTRAP_FAILURE_HOST_NORMAL_EXIT":"RECOVERED_HOST_NORMAL_EXIT"):"ABNORMAL_EXIT";persist();
  File.WriteAllText(output,new JavaScriptSerializer().Serialize(state),new UTF8Encoding(false));process.Dispose();
  if(Convert.ToInt32(state["ExitCode"])!=0)throw new InvalidOperationException("Owned Excel exited abnormally; failed qualification remains failed.");
 }
}
'@
[Q026RecoveryShutdown]::Run($desktop,$startup.ProcessId,$startup.HostStartedUtc,$startup.InitialWorkbook.FullName,$startup.HostExecutable,$OutputPath,[bool]$BootstrapClosureObservation)
Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8
