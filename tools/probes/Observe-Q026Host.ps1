#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$CancelObservationPath,[switch]$ReadDialogText)
$ErrorActionPreference='Stop'
if (-not [IO.Path]::IsPathRooted($OutputPath) -or (Test-Path -LiteralPath $OutputPath)) { throw 'A fresh absolute receipt is required.' }
$hostRoot=@(Get-ChildItem -LiteralPath (Join-Path $EvidenceRoot 'native/hosts') -Directory)
if($hostRoot.Count -ne 1){throw 'Exactly one owned host receipt is required.'}
$startup=Get-Content -LiteralPath (Join-Path $hostRoot[0].FullName 'startup.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$desktop=(Get-Content -LiteralPath (Join-Path $EvidenceRoot 'isolation/desktop/desktop-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json).Desktop
if($desktop -notmatch '^VBAiTests_[0-9a-f]{32}$' -or -not $startup.Owned){throw 'Reviewed private owned launch is required.'}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
public static class Q026HostObservation {
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr OpenDesktopW(string n,uint f,bool i,uint a);
 [DllImport("user32.dll",SetLastError=true)] static extern bool SetThreadDesktop(IntPtr d);
 [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr d);
 delegate bool Visitor(IntPtr h,IntPtr p);
 [DllImport("user32.dll",SetLastError=true)] static extern bool EnumWindows(Visitor v,IntPtr p);
 [DllImport("user32.dll",SetLastError=true)] static extern bool EnumChildWindows(IntPtr h,Visitor v,IntPtr p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr SendMessageTimeoutW(IntPtr h,uint m,UIntPtr w,StringBuilder l,uint f,uint t,out UIntPtr r);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h,StringBuilder s,int c);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h,StringBuilder s,int c);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
 [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr h,int id);
 [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
 [DllImport("user32.dll",SetLastError=true)] static extern bool PostMessageW(IntPtr h,uint m,IntPtr w,IntPtr l);
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateWindowExW(uint e,string c,string t,uint s,int x,int y,int w,int h,IntPtr p,IntPtr m,IntPtr i,IntPtr a);
 [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr h);
 public static int CancelDispatchEntries;
 public static void Cancel(string desktop,int pid,string start,long captured) {
   CancelDispatchEntries=0;Exception error=null;
   var thread=new Thread(()=> {
    IntPtr d=OpenDesktopW(desktop,0,false,0xC7), reader=IntPtr.Zero;
    try {
     if(d==IntPtr.Zero || !SetThreadDesktop(d)) throw new Win32Exception(Marshal.GetLastWin32Error());
     reader=CreateWindowExW(0,"STATIC","Q026 explicit cancel recovery",0,0,0,1,1,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
     if(reader==IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
     using(var process=Process.GetProcessById(pid)) {
       if(process.ProcessName!="EXCEL" || process.StartTime.ToUniversalTime()!=DateTime.Parse(start).ToUniversalTime()) throw new InvalidOperationException("Owned PID/start mismatch.");
     }
     var dialog=new IntPtr(captured);uint owner;GetWindowThreadProcessId(dialog,out owner);
     var cls=new StringBuilder(256);var title=new StringBuilder(4096);
     GetClassNameW(dialog,cls,cls.Capacity);GetWindowTextW(dialog,title,title.Capacity);
     if(owner!=pid || cls.ToString()!="#32770" || title.ToString()!="Options" || !IsWindowVisible(dialog)) throw new InvalidOperationException("Captured Options dialog identity changed.");
     var cancel=GetDlgItem(dialog,2);GetWindowThreadProcessId(cancel,out owner);cls.Clear();GetClassNameW(cancel,cls,cls.Capacity);
     if(cancel==IntPtr.Zero || owner!=pid || GetParent(cancel)!=dialog || cls.ToString()!="Button" || !IsWindowEnabled(cancel)) throw new InvalidOperationException("Captured Cancel button identity unavailable.");
     CancelDispatchEntries++;
     if(!PostMessageW(cancel,0xF5,IntPtr.Zero,IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
    } catch(Exception e){error=e;} finally {if(reader!=IntPtr.Zero)DestroyWindow(reader);if(d!=IntPtr.Zero)CloseDesktop(d);}
   });thread.SetApartmentState(ApartmentState.STA);thread.Start();
   if(!thread.Join(10000)) throw new InvalidOperationException("Cancellation observation exceeded deadline; retain without replay.");
   if(error!=null)throw error;
 }
 public static bool IncludeDialogText;
 public static object[] Read(string desktop,int pid,string start) {
   using(var process=Process.GetProcessById(pid)) {
     if(process.ProcessName!="EXCEL" || process.StartTime.ToUniversalTime()!=DateTime.Parse(start).ToUniversalTime()) throw new InvalidOperationException("Owned PID/start mismatch.");
   }
   var rows=new List<object>(); Exception error=null;
   var thread=new Thread(()=> {
     IntPtr d=OpenDesktopW(desktop,0,false,0xC7), reader=IntPtr.Zero;
     try {
       if(d==IntPtr.Zero || !SetThreadDesktop(d)) throw new Win32Exception(Marshal.GetLastWin32Error());
       reader=CreateWindowExW(0,"STATIC","Q026 read-only window inventory",0,0,0,1,1,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
       if(reader==IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
       bool ok=EnumWindows((h,p)=> {
         uint owner; if(GetWindowThreadProcessId(h,out owner)==0) throw new InvalidOperationException("Window ownership unreadable.");
         if(owner!=pid) return true;
         var cls=new StringBuilder(256); var title=new StringBuilder(4096);
         if(GetClassNameW(h,cls,cls.Capacity)==0) throw new InvalidOperationException("Owned class unreadable.");
         GetWindowTextW(h,title,title.Capacity);
         var children=new List<object>(); Exception childError=null;
         if(IncludeDialogText && cls.ToString()=="#32770") {
           EnumChildWindows(h,(child,unused)=> {
             try {
               if(children.Count>=32) throw new InvalidOperationException("Dialog control inventory exceeds the bound.");
               uint childOwner;GetWindowThreadProcessId(child,out childOwner);
               if(childOwner!=pid) throw new InvalidOperationException("Dialog child ownership changed.");
               var childClass=new StringBuilder(256);GetClassNameW(child,childClass,childClass.Capacity);
               var text=new StringBuilder(4096);UIntPtr result;
               if(SendMessageTimeoutW(child,0x000D,new UIntPtr(4096),text,3,100,out result)==IntPtr.Zero)
                 throw new InvalidOperationException("Dialog text read did not complete; no input sent.");
               if(result.ToUInt64()>=4095) throw new InvalidOperationException("Dialog text exceeds the bound.");
               children.Add(new {Handle=child.ToInt64(),Class=childClass.ToString(),Text=text.ToString()});return true;
             } catch(Exception failure) {childError=failure;return false;}
           },IntPtr.Zero);
           if(childError!=null) throw childError;
         }
         rows.Add(new {Handle=h.ToInt64(),ProcessId=owner,Class=cls.ToString(),Title=title.ToString(),Visible=IsWindowVisible(h),Children=children.ToArray()}); return true;
       },IntPtr.Zero);
       if(!ok) throw new Win32Exception(Marshal.GetLastWin32Error());
     } catch(Exception e) {error=e;} finally {if(reader!=IntPtr.Zero)DestroyWindow(reader);if(d!=IntPtr.Zero)CloseDesktop(d);}
   }); thread.SetApartmentState(ApartmentState.STA); thread.Start();
   if(!thread.Join(10000)) throw new InvalidOperationException("Read-only desktop enumeration exceeded its deadline.");
   if(error!=null) throw error;
   using(var process=Process.GetProcessById(pid)) {
     if(process.HasExited || process.StartTime.ToUniversalTime()!=DateTime.Parse(start).ToUniversalTime()) throw new InvalidOperationException("Owned identity changed during observation.");
   }
   return rows.ToArray();
 }
}
'@
if($CancelObservationPath){
    $prior=Get-Content -LiteralPath $CancelObservationPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $dialogs=@($prior.Windows | Where-Object {$_.Class -eq '#32770'})
    if($prior.ProcessId -ne $startup.ProcessId -or $prior.ProcessStartUtc -ne $startup.HostStartedUtc -or
        $prior.Desktop -ne $desktop -or -not $prior.EnumerationSucceeded -or $dialogs.Count -ne 1 -or $dialogs[0].Title -ne 'Options'){
        throw 'Exactly one independently captured owned Options dialog is required.'
    }
    $intent=$OutputPath+'.cancel-intent.json'
    if(Test-Path -LiteralPath $intent){throw 'Cancellation already claimed; do not replay.'}
    @{ProcessId=$startup.ProcessId;ProcessStartUtc=$startup.HostStartedUtc;Dialog=$dialogs[0].Handle;
        ExplicitRecovery=$true;MaximumDispatchEntries=1;NativeReplayAllowed=$false;Utc=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath $intent -Encoding UTF8
    try { [Q026HostObservation]::Cancel($desktop,$startup.ProcessId,$startup.HostStartedUtc,$dialogs[0].Handle) }
    catch {
        @{DispatchEntries=[Q026HostObservation]::CancelDispatchEntries;Error=$_.Exception.ToString();NativeReplayAllowed=$false} |
            ConvertTo-Json | Set-Content -LiteralPath ($OutputPath+'.cancel-failure.json') -Encoding UTF8
        throw
    }
    $closed=$false
    for($attempt=0;$attempt -lt 40;$attempt++){
        $windows=@([Q026HostObservation]::Read($desktop,$startup.ProcessId,$startup.HostStartedUtc))
        if(@($windows | Where-Object {$_.Class -eq '#32770'}).Count -eq 0){$closed=$true;break}
        Start-Sleep -Milliseconds 250
    }
    if(-not $closed){throw 'Captured cancellation closure not observed; retain without replay.'}
}
[Q026HostObservation]::IncludeDialogText=[bool]$ReadDialogText
$windows=@([Q026HostObservation]::Read($desktop,$startup.ProcessId,$startup.HostStartedUtc))
@{ObservedUtc=[DateTime]::UtcNow.ToString('o');ProcessId=$startup.ProcessId;ProcessStartUtc=$startup.HostStartedUtc;
    Desktop=$desktop;Windows=$windows;OptionsDialogAbsent=(@($windows | Where-Object {$_.Class -eq '#32770'}).Count -eq 0);
    EnumerationSucceeded=$true;ObservationOnly=(-not $CancelObservationPath);BridgeCommands=0;HostInput=0;
    ExplicitCancelRecovery=[bool]$CancelObservationPath;CleanupInvoked=[bool]$CancelObservationPath} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Get-Content -LiteralPath $OutputPath -Raw
