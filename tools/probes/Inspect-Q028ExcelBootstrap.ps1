#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$BootstrapPath,[Parameter(Mandatory=$true)][string]$OutputPath,
    [switch]$CloseOwnedSeed, [string]$ObservationPath)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $OutputPath){throw 'Observation one-shot receipt already exists'}
$boot=Get-Content -LiteralPath $BootstrapPath -Raw -Encoding UTF8|ConvertFrom-Json
if($boot.Phase -ne 'FailedPreserved' -or $boot.FailedAtPhase -ne 'ExactNativeApplicationAndSeedAttached' -or $boot.BootstrapCloseAttempts -ne 0 -or $boot.BootstrapQuitAttempts -ne 0 -or $boot.LoadedAssemblyMvid){throw 'Exact bootstrap-before-bridge failure required'}
$seedStream=[IO.FileStream]::new($boot.Seed,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
$seedSha=[Security.Cryptography.SHA256]::Create()
try {$seedHash=[BitConverter]::ToString($seedSha.ComputeHash($seedStream)).Replace('-','')}
finally {$seedSha.Dispose();$seedStream.Dispose()}
if($seedHash -ine $boot.SeedSha256){throw 'Seed changed'}
if($CloseOwnedSeed){
    $observed=Get-Content -LiteralPath $ObservationPath -Raw -Encoding UTF8|ConvertFrom-Json
    if($observed.State -ne 'OBSERVED_NO_MUTATION' -or -not $observed.ObservationOnly -or
       $observed.ProcessId -ne $boot.ProcessId -or $observed.ProcessStartUtc -ne $boot.ProcessStartUtc -or
       $observed.AddInConnect -ne $false -or -not $observed.SeedSaved -or $observed.SeedFullName -ne $boot.Seed){throw 'Exact prior disconnected-add-in saved-seed observation required'}
}
$procs=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
if($procs.Count -ne 1 -or $procs[0].Id -ne $boot.ProcessId){throw 'Exact sole test Excel required'}
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
public static class Q028BootstrapObservation {
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
 public static void Run(string desktop,int pid,string start,string seed,string executable,string output,bool close) {
  var state=new Dictionary<string,object>{{"ProcessId",pid},{"ProcessStartUtc",start},{"Desktop",desktop},{"Qualified",false},{"ForcedTermination",false},{"CloseEntries",0},{"QuitEntries",0},{"State","IDENTITY_GUARD"}};
  Action persist=()=>File.WriteAllText(output+".progress.json",new JavaScriptSerializer().Serialize(state),new UTF8Encoding(false));
  state["RecoveryKind"]=close?"BootstrapFailureSavedSeedShutdown":"ReadOnlyBootstrapObservation";
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
    reader=CreateWindowExW(0,"STATIC","Q028 bootstrap observation and recovery",0,0,0,1,1,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
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
    if(close) {
      requireDispatch();state["State"]="CLOSE_INTENT";state["CloseEntries"]=1;persist();
      Call(book,"Close",false,Type.Missing,Type.Missing);state["CloseReturned"]=true;
      if(Convert.ToInt32(Get(books,"Count"))!=0)throw new InvalidOperationException("Owned seed close unverified; no Quit.");
      requireDispatch();state["State"]="QUIT_INTENT";state["QuitEntries"]=1;persist();
      Call(app,"Quit");state["QuitReturned"]=true;state["State"]="QUIT_RETURNED";persist();
    } else {
    state["ReadStage"]="VBE";persist();
    object vbe=Get(app,"VBE");state["ReadStage"]="AddIns";persist();
    object addins=Get(vbe,"AddIns"), entry=null;
try {
 state["ReadStage"]="Item(VBAi.AddIn)";persist();
 entry=Call(addins,"Item","VBAi.AddIn");
 state["ReadStage"]="Entry identity and connection";persist();
 state["AddInProgId"]=Get(entry,"ProgId");state["AddInGuid"]=Get(entry,"Guid");state["AddInConnect"]=Get(entry,"Connect");
 state["SeedSaved"]=Get(book,"Saved");state["SeedFullName"]=Get(book,"FullName");
 state["ObservationOnly"]=true;state["State"]="OBSERVED_NO_MUTATION";persist();
} finally { Release(entry);Release(addins);Release(vbe); }} } catch(Exception e){failure=e;} finally {
    if(failure!=null && Convert.ToInt32(state["CloseEntries"])!=0 && !state.ContainsKey("QuitReturned")){
      retainedReferences.AddRange(new[]{book,books,app,window});state["State"]="RETAINED_NATIVE_OUTCOME";state["Error"]=failure.ToString();persist();
      for(;;)Thread.Sleep(1000);
    }
    try{Release(book);Release(books);Release(app);Release(window);}catch(Exception e){if(failure==null)failure=e;}
    if(reader!=IntPtr.Zero)DestroyWindow(reader);if(desk!=IntPtr.Zero)CloseDesktop(desk);
   }
  });thread.SetApartmentState(ApartmentState.STA);thread.IsBackground=true;thread.Start();
  if(!thread.Join(30000)){Interlocked.Exchange(ref stopDispatch,1);state["State"]="RETAINED_PENDING_COM";persist();for(;;)Thread.Sleep(1000);}
  if(failure!=null){state["State"]="OBSERVATION_ERROR";state["Error"]=failure.ToString();persist();throw failure;}
  if(close) {
    if(!process.WaitForExit(15000)){state["State"]="RETAINED_EXIT_NOT_OBSERVED";persist();for(;;)Thread.Sleep(1000);}
    state["ExitCode"]=process.ExitCode;state["State"]=process.ExitCode==0?"BOOTSTRAP_FAILURE_HOST_NORMAL_EXIT":"ABNORMAL_EXIT";persist();
  }
  File.WriteAllText(output,new JavaScriptSerializer().Serialize(state),new UTF8Encoding(false));process.Dispose();
 }
}
'@
[Q028BootstrapObservation]::Run($boot.Desktop,$boot.ProcessId,$boot.ProcessStartUtc,$boot.Seed,$boot.Executable,$OutputPath,[bool]$CloseOwnedSeed)
Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8
