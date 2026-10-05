using System;using System.Collections.Generic;using System.ComponentModel;using System.Diagnostics;using System.IO;using System.Runtime.InteropServices;using System.Text;
public sealed class Q020OwnedLayout {
 readonly int pid;readonly string birth,image;readonly IntPtr root;readonly uint rootThread;
 public Q020OwnedLayout(int hostPid,string hostBirth,string hostImage,long nativeRoot,uint nativeRootThread){if(hostPid<=0||string.IsNullOrEmpty(hostBirth)||!Path.IsPathRooted(hostImage)||nativeRoot<=0||nativeRootThread==0)throw new ArgumentException("Actual owned host/root identity required.");pid=hostPid;birth=hostBirth;image=Path.GetFullPath(hostImage);root=new IntPtr(nativeRoot);rootThread=nativeRootThread;}
 [StructLayout(LayoutKind.Sequential)] public struct Point {public int X,Y;}
 [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)] public struct Placement {public int Length,Flags,ShowCmd;public Point MinPosition,MaxPosition;public Rect NormalPosition;}
 [StructLayout(LayoutKind.Sequential)] struct MonitorInfo {public int Size;public Rect Monitor,Work;public uint Flags;}
 public sealed class Layout {public Placement Before;public Rect Work,Target;public bool ShowRestoreEntered,SetPositionEntered,RestoreEntered,RestoreVerified;internal Q020OwnedLayout Owner;}
 public sealed class Toolbox {public long Handle,Owner,RootOwner;public uint Pid,Thread;public string Class,Caption;public bool Visible,HideEntered,HideReturned,HidePreviousVisibility,HiddenVerified,RestoreEntered,RestoreReturned,RestorePreviousVisibility,RestoredVerified;public Rect BeforeRect;public Placement BeforePlacement;internal Q020OwnedLayout Context;}
 delegate bool Visitor(IntPtr w,IntPtr unused);
 [DllImport("kernel32.dll")]static extern uint GetCurrentThreadId();
 [DllImport("kernel32.dll")]static extern void SetLastError(uint code);
 [DllImport("user32.dll")]static extern IntPtr GetThreadDesktop(uint tid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)]static extern bool GetUserObjectInformation(IntPtr h,int type,StringBuilder s,int bytes,out int n);
 [DllImport("user32.dll",SetLastError=true)]static extern bool EnumDesktopWindows(IntPtr d,Visitor callback,IntPtr state);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);
 [DllImport("user32.dll")]static extern IntPtr GetAncestor(IntPtr w,uint flag);
 [DllImport("user32.dll")]static extern IntPtr GetWindow(IntPtr w,uint flag);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr w);
 [DllImport("user32.dll")]static extern bool IsIconic(IntPtr w);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetWindowRect(IntPtr w,out Rect rect);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetWindowPlacement(IntPtr w,ref Placement p);
 [DllImport("user32.dll",SetLastError=true)]static extern bool SetWindowPlacement(IntPtr w,ref Placement p);
 [DllImport("user32.dll")]static extern bool ShowWindow(IntPtr w,int command);
 [DllImport("user32.dll",SetLastError=true)]static extern bool SetWindowPos(IntPtr w,IntPtr after,int x,int y,int width,int height,uint flags);
 [DllImport("user32.dll")]static extern IntPtr MonitorFromWindow(IntPtr w,uint flags);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetMonitorInfo(IntPtr monitor,ref MonitorInfo info);
 public void RequireOwner(){
  var desktop=new StringBuilder(256);int n;if(!GetUserObjectInformation(GetThreadDesktop(GetCurrentThreadId()),2,desktop,512,out n)||desktop.ToString()!="Default")throw new InvalidOperationException("Owned layout actor desktop differs.");
  using(var p=Process.GetProcessById(pid)){if(p.HasExited||p.StartTime.ToUniversalTime().ToString("o")!=birth||!string.Equals(p.MainModule.FileName,image,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Owned layout generation differs.");}
  uint actual;uint tid=GetWindowThreadProcessId(root,out actual);var cls=new StringBuilder(256);if(actual!=(uint)pid||tid!=rootThread||GetAncestor(root,2)!=root||!IsWindowVisible(root)||GetClassName(root,cls,256)==0||cls.ToString()!="wndclass_desked_gsk")throw new InvalidOperationException("Exact native root changed.");
 }
 public Toolbox[] ReadToolboxes(){
  RequireOwner();
  var rows=new List<Toolbox>();Exception error=null;int count=0;
  Visitor visitor=(w,s)=>{try{
   if(++count>8192)throw new InvalidOperationException("Complete Default inventory bound exceeded.");
   uint ownerPid;uint tid=GetWindowThreadProcessId(w,out ownerPid);if(ownerPid!=(uint)pid)return true;if(tid==0)throw new InvalidOperationException("Owned native HWND thread unavailable.");
   var cls=new StringBuilder(256);int length=GetClassName(w,cls,256);if(length<=0||length>=255)throw new InvalidOperationException("Owned native HWND class incomplete.");
   if(!cls.ToString().StartsWith("F3 MinFrame ",StringComparison.Ordinal))return true;
   var caption=new StringBuilder(256);int textLength=GetWindowText(w,caption,256);if(textLength>=255)throw new InvalidOperationException("Owned Toolbox caption truncated.");if(caption.ToString()!="Boîte à outils"&&caption.ToString()!="Toolbox")return true;
   if(tid!=rootThread||GetWindow(w,4)!=root||GetAncestor(w,3)!=root||GetAncestor(w,2)!=w)throw new InvalidOperationException("Exact floating Toolbox native owner/root/thread differs.");
   Rect rect;if(!GetWindowRect(w,out rect))throw new Win32Exception(Marshal.GetLastWin32Error());var placement=new Placement{Length=Marshal.SizeOf(typeof(Placement))};if(!GetWindowPlacement(w,ref placement))throw new Win32Exception(Marshal.GetLastWin32Error());if(IsIconic(w))throw new InvalidOperationException("Owned floating Toolbox is iconic.");
   rows.Add(new Toolbox{Context=this,BeforeRect=rect,BeforePlacement=placement,Handle=w.ToInt64(),Pid=ownerPid,Thread=tid,Class=cls.ToString(),Caption=caption.ToString(),Visible=IsWindowVisible(w),Owner=root.ToInt64(),RootOwner=root.ToInt64()});return true;
  }catch(Exception e){error=e;SetLastError(0);return false;}};
  SetLastError(0);bool complete=EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()),visitor,IntPtr.Zero);int last=Marshal.GetLastWin32Error();GC.KeepAlive(visitor);
  if(error!=null)throw error;if(!complete)throw new Win32Exception(last,"Default Toolbox inventory incomplete.");RequireOwner();return rows.ToArray();
 }
 public Toolbox DiscoverNativeToolbox(){var rows=ReadToolboxes();Toolbox answer=null;int count=0;foreach(var row in rows)if(row.Visible){answer=row;count++;}if(count!=1)throw new InvalidOperationException("Exactly one fresh visible native MSForms Toolbox required; no VBIDE Type10 association.");return answer;}
 static bool EqualRect(Rect a,Rect b){return a.Left==b.Left&&a.Top==b.Top&&a.Right==b.Right&&a.Bottom==b.Bottom;}
 void RequireToolbox(Toolbox state,bool visible){RequireOwner();if(state==null||state.Context!=this)throw new InvalidOperationException("Toolbox state belongs to another owner.");IntPtr w=new IntPtr(state.Handle);uint actual;uint tid=GetWindowThreadProcessId(w,out actual);var cls=new StringBuilder(256);var caption=new StringBuilder(256);Rect rect;
  if(actual!=(uint)pid||tid!=rootThread||GetWindow(w,4)!=root||GetAncestor(w,3)!=root||GetAncestor(w,2)!=w||GetClassName(w,cls,256)==0||cls.ToString()!=state.Class||GetWindowText(w,caption,256)>=255||caption.ToString()!=state.Caption||IsWindowVisible(w)!=visible||IsIconic(w))throw new InvalidOperationException("Original Toolbox HWND/owner/thread/class/caption/visibility changed; no replacement action.");
  if(!GetWindowRect(w,out rect))throw new Win32Exception(Marshal.GetLastWin32Error());if(!EqualRect(rect,state.BeforeRect))throw new InvalidOperationException("Original native Toolbox rectangle changed.");
 }
 public void HideToolbox(Toolbox state,Action guard){if(state==null||state.HideEntered||guard==null)throw new InvalidOperationException("Native Toolbox hide claim invalid or already entered.");guard();RequireToolbox(state,true);state.HideEntered=true;state.HidePreviousVisibility=ShowWindow(new IntPtr(state.Handle),0);state.HideReturned=true;RequireToolbox(state,false);state.HiddenVerified=true;}
 public void RestoreToolbox(Toolbox state,Action guard){if(state==null||!state.HideEntered||!state.HideReturned||state.RestoreEntered||guard==null)throw new InvalidOperationException("Native Toolbox restore claim pending/invalid/already entered.");guard();RequireToolbox(state,false);state.RestoreEntered=true;state.RestorePreviousVisibility=ShowWindow(new IntPtr(state.Handle),8);state.RestoreReturned=true;RequireToolbox(state,true);var after=new Placement{Length=Marshal.SizeOf(typeof(Placement))};if(!GetWindowPlacement(new IntPtr(state.Handle),ref after))throw new Win32Exception(Marshal.GetLastWin32Error());if(!Equal(state.BeforePlacement,after))throw new InvalidOperationException("Original complete Toolbox WINDOWPLACEMENT changed.");state.RestoredVerified=true;}
 public Layout Backup(){RequireOwner();var p=new Placement{Length=Marshal.SizeOf(typeof(Placement))};if(!GetWindowPlacement(root,ref p))throw new Win32Exception(Marshal.GetLastWin32Error());var mi=new MonitorInfo{Size=Marshal.SizeOf(typeof(MonitorInfo))};if(!GetMonitorInfo(MonitorFromWindow(root,2),ref mi))throw new Win32Exception(Marshal.GetLastWin32Error());int width=Math.Min(1400,mi.Work.Right-mi.Work.Left-160),height=Math.Min(850,mi.Work.Bottom-mi.Work.Top-160);if(width<640||height<480)throw new InvalidOperationException("Finite monitor work area too small.");return new Layout{Owner=this,Before=p,Work=mi.Work,Target=new Rect{Left=mi.Work.Left+80,Top=mi.Work.Top+80,Right=mi.Work.Left+80+width,Bottom=mi.Work.Top+80+height}};}
 public void Apply(Layout state,Action guard){if(state==null||state.Owner!=this||state.ShowRestoreEntered||state.SetPositionEntered||guard==null)throw new InvalidOperationException("Layout ownership/action claim invalid.");guard();RequireOwner();state.ShowRestoreEntered=true;ShowWindow(root,9);guard();RequireOwner();state.SetPositionEntered=true;if(!SetWindowPos(root,IntPtr.Zero,state.Target.Left,state.Target.Top,state.Target.Right-state.Target.Left,state.Target.Bottom-state.Target.Top,0x14))throw new Win32Exception(Marshal.GetLastWin32Error());RequireOwner();}
 static bool Equal(Placement a,Placement b){return a.Flags==b.Flags&&a.ShowCmd==b.ShowCmd&&a.MinPosition.X==b.MinPosition.X&&a.MinPosition.Y==b.MinPosition.Y&&a.MaxPosition.X==b.MaxPosition.X&&a.MaxPosition.Y==b.MaxPosition.Y&&a.NormalPosition.Left==b.NormalPosition.Left&&a.NormalPosition.Top==b.NormalPosition.Top&&a.NormalPosition.Right==b.NormalPosition.Right&&a.NormalPosition.Bottom==b.NormalPosition.Bottom;}
 public void Restore(Layout state,Action guard){if(state==null||state.Owner!=this||state.RestoreEntered||!state.ShowRestoreEntered||guard==null)throw new InvalidOperationException("Original layout restore already entered/unneeded.");guard();RequireOwner();state.RestoreEntered=true;var original=state.Before;original.Length=Marshal.SizeOf(typeof(Placement));if(!SetWindowPlacement(root,ref original))throw new Win32Exception(Marshal.GetLastWin32Error());var after=new Placement{Length=Marshal.SizeOf(typeof(Placement))};if(!GetWindowPlacement(root,ref after))throw new Win32Exception(Marshal.GetLastWin32Error());state.RestoreVerified=Equal(state.Before,after);if(!state.RestoreVerified)throw new InvalidOperationException("Original complete WINDOWPLACEMENT not restored exactly.");RequireOwner();}
}
