using System;using System.ComponentModel;using System.Diagnostics;using System.Drawing;using System.Drawing.Imaging;using System.IO;using System.Runtime.InteropServices;using System.Text;using System.Threading;

// Actual composed screen pixels of one foreground owned target ROI. Manual review is still required.
public static class Q020NativeCapture {
 public sealed class CaptureState {
  public volatile bool Completed;public bool Returned,DpiRestored,ForegroundVerifiedBefore,ForegroundVerifiedAfter,FGReturn;
  public int BarrierAttempts,BarrierLastError;public long ImmediateForegroundRoot,SettledForegroundRoot,BarrierReturn;
  public string Error,OutputPath,Desktop,ThreadDesktop,InputBefore,InputAfter,StartedUtc,FinishedUtc;
  public int Width,Height,Left,Top,HostPid,FocusAttempts;public uint OwnerThreadId;public long Window,Bytes;
 }
 [StructLayout(LayoutKind.Sequential)]struct Rect{public int Left,Top,Right,Bottom;}
 delegate bool Visitor(IntPtr w,IntPtr unused);
 [DllImport("kernel32.dll")]static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")]static extern IntPtr GetThreadDesktop(uint thread);
 [DllImport("user32.dll")]static extern IntPtr GetProcessWindowStation();
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)]static extern bool GetUserObjectInformation(IntPtr h,int type,StringBuilder value,uint size,out uint needed);
 [DllImport("user32.dll",SetLastError=true)]static extern IntPtr OpenInputDesktop(uint flags,bool inherit,uint access);
 [DllImport("user32.dll",SetLastError=true)]static extern bool CloseDesktop(IntPtr h);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr w,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr w,StringBuilder value,int size);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr w);
 [DllImport("user32.dll")]static extern bool IsWindowEnabled(IntPtr w);
 [DllImport("user32.dll")]static extern bool IsWindow(IntPtr w);
 [DllImport("user32.dll")]static extern bool IsIconic(IntPtr w);
 [DllImport("user32.dll")]static extern IntPtr GetAncestor(IntPtr w,uint type);
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern bool SetForegroundWindow(IntPtr w);
 [DllImport("kernel32.dll",EntryPoint="SetLastError")]static extern void ClearLastError(uint error);
 [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)]static extern IntPtr SendMessageTimeoutW(IntPtr window,uint message,UIntPtr wparam,IntPtr lparam,uint flags,uint timeout,out UIntPtr result);
 [DllImport("user32.dll")]static extern int GetSystemMetrics(int type);
 [DllImport("user32.dll",SetLastError=true)]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]static extern IntPtr GetThreadDpiAwarenessContext();
 [DllImport("user32.dll")]static extern bool AreDpiAwarenessContextsEqual(IntPtr a,IntPtr b);
 [DllImport("user32.dll")]static extern IntPtr GetWindow(IntPtr window,uint command);
 [DllImport("user32.dll",SetLastError=true)]static extern bool EnumDesktopWindows(IntPtr desktop,Visitor callback,IntPtr unused);
 [DllImport("dwmapi.dll",EntryPoint="DwmGetWindowAttribute")]static extern int Frame(IntPtr w,int attribute,out Rect rect,int size);
 [DllImport("dwmapi.dll",EntryPoint="DwmGetWindowAttribute")]static extern int Cloaked(IntPtr w,int attribute,out int value,int size);
 static string Name(IntPtr h){var b=new StringBuilder(256);uint n;if(h==IntPtr.Zero||!GetUserObjectInformation(h,2,b,512,out n))throw new Win32Exception(Marshal.GetLastWin32Error());return b.ToString();}
 static string Input(){IntPtr h=OpenInputDesktop(0,false,1);if(h==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());try{return Name(h);}finally{if(!CloseDesktop(h))throw new Win32Exception(Marshal.GetLastWin32Error());}}
 static void Context(){if(Name(GetThreadDesktop(GetCurrentThreadId()))!="Default"||Name(GetProcessWindowStation())!="WinSta0"||Input()!="Default")throw new InvalidOperationException("Exact active main desktop required.");}
 static uint Owner(int pid,string birth,string image,IntPtr window){Context();using(var p=Process.GetProcessById(pid)){if(p.HasExited||p.StartTime.ToUniversalTime().ToString("o")!=birth||!string.Equals(p.MainModule.FileName,image,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Original capture generation changed.");}uint actual;uint tid=GetWindowThreadProcessId(window,out actual);var cls=new StringBuilder(256);if(tid==0||actual!=(uint)pid||GetClassName(window,cls,256)==0||cls.ToString()!="wndclass_desked_gsk"||GetAncestor(window,2)!=window||!IsWindowVisible(window)||!IsWindowEnabled(window)||IsIconic(window))throw new InvalidOperationException("Exact enabled visible owned VBE root unavailable.");bool found=false;int count=0;Exception inventoryError=null;Visitor visitor=(w,s)=>{try{if(++count>8192)throw new InvalidOperationException("Default root inventory bound exceeded.");if(w==window){uint ownerPid;uint ownerTid=GetWindowThreadProcessId(w,out ownerPid);if(ownerPid!=(uint)pid||ownerTid!=tid||GetAncestor(w,2)!=w)throw new InvalidOperationException("Owned root identity changed during inventory.");found=true;}return true;}catch(Exception e){inventoryError=e;return false;}};bool complete=EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()),visitor,IntPtr.Zero);GC.KeepAlive(visitor);if(inventoryError!=null)throw inventoryError;if(!complete||!found)throw new InvalidOperationException("Exact owned root membership in Default desktop not proven.");Context();return tid;}
 static Rect Bounds(IntPtr w){Rect r;Marshal.ThrowExceptionForHR(Frame(w,9,out r,Marshal.SizeOf(typeof(Rect))));if(r.Right<=r.Left||r.Bottom<=r.Top)throw new InvalidOperationException("Physical frame bounds unavailable.");return r;}
 static bool Equal(Rect a,Rect b){return a.Left==b.Left&&a.Top==b.Top&&a.Right==b.Right&&a.Bottom==b.Bottom;}
 static void Unoccluded(IntPtr target,Rect roi){
  var seen=new System.Collections.Generic.HashSet<IntPtr>();int count=0;
  // GW_HWNDPREV traverses windows preceding the target in the actual Z order.
  for(IntPtr w=GetWindow(target,3);w!=IntPtr.Zero;w=GetWindow(w,3)){
   if(++count>8192||!seen.Add(w))throw new InvalidOperationException("Foreground Z-order prefix is incomplete or cyclic.");
   if(!IsWindow(w))throw new InvalidOperationException("Z-order topology changed during admission.");
   if(!IsWindowVisible(w)||IsIconic(w))continue;
   int cloaked;Marshal.ThrowExceptionForHR(Cloaked(w,14,out cloaked,4));if(cloaked!=0)continue;
   Rect r=Bounds(w);if(r.Left<roi.Right&&r.Right>roi.Left&&r.Top<roi.Bottom&&r.Bottom>roi.Top)throw new InvalidOperationException("A visible window above the owned ROI would contaminate screen evidence.");
  }
  if(!IsWindow(target)||GetAncestor(GetForegroundWindow(),2)!=target)throw new InvalidOperationException("Owned foreground root changed during Z-order admission.");
 }
 public static CaptureState Begin(int pid,string birthUtc,string imagePath,string desktop,long rootWindow,string outputPath){if(pid<=0||rootWindow<=0||desktop!="Default"||!Path.IsPathRooted(imagePath)||!Path.IsPathRooted(outputPath)||File.Exists(outputPath)||!Directory.Exists(Path.GetDirectoryName(outputPath)))throw new ArgumentException("Fresh explicit owned Default capture required.");var state=new CaptureState{HostPid=pid,Window=rootWindow,Desktop=desktop,OutputPath=outputPath,StartedUtc=DateTime.UtcNow.ToString("o")};var thread=new Thread(()=>{IntPtr previousDpi=IntPtr.Zero;var elapsed=Stopwatch.StartNew();try{
  Context();state.ThreadDesktop=Name(GetThreadDesktop(GetCurrentThreadId()));state.InputBefore=Input();IntPtr window=new IntPtr(rootWindow);state.OwnerThreadId=Owner(pid,birthUtc,imagePath,window);
  previousDpi=SetThreadDpiAwarenessContext(new IntPtr(-4));if(previousDpi==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());
  state.FocusAttempts=1;state.FGReturn=SetForegroundWindow(window);state.ImmediateForegroundRoot=GetAncestor(GetForegroundWindow(),2).ToInt64();
  if(!state.FGReturn)throw new InvalidOperationException("The one allowed foreground claim returned false.");
  long remaining=20000-elapsed.ElapsedMilliseconds;if(remaining<=0)throw new InvalidOperationException("Capture deadline exhausted before foreground barrier.");
  state.BarrierAttempts=1;UIntPtr barrierResult;ClearLastError(0);IntPtr barrierReturn=SendMessageTimeoutW(window,0,UIntPtr.Zero,IntPtr.Zero,0x22,(uint)Math.Min(5000,remaining),out barrierResult);
  int barrierError=Marshal.GetLastWin32Error();state.BarrierReturn=barrierReturn.ToInt64();state.BarrierLastError=barrierReturn==IntPtr.Zero?barrierError:0;
  if(barrierReturn==IntPtr.Zero)throw new InvalidOperationException("The one foreground barrier failed; Win32 error="+state.BarrierLastError+".");
  if(elapsed.ElapsedMilliseconds>=20000)throw new InvalidOperationException("Capture deadline exhausted after foreground barrier.");
  if(Owner(pid,birthUtc,imagePath,window)!=state.OwnerThreadId)throw new InvalidOperationException("Owned VBE thread changed after foreground barrier.");
  state.SettledForegroundRoot=GetAncestor(GetForegroundWindow(),2).ToInt64();if(state.SettledForegroundRoot!=rootWindow)throw new InvalidOperationException("Owned foreground root not verified after the one barrier.");state.ForegroundVerifiedBefore=true;
  Rect roi=Bounds(window);state.Left=roi.Left;state.Top=roi.Top;state.Width=roi.Right-roi.Left;state.Height=roi.Bottom-roi.Top;
  int x=GetSystemMetrics(76),y=GetSystemMetrics(77),width=GetSystemMetrics(78),height=GetSystemMetrics(79);
  if(state.Width<64||state.Height<64||state.Width>8192||state.Height>8192||(long)state.Width*state.Height>16777216||roi.Left<x||roi.Top<y||roi.Right>(long)x+width||roi.Bottom>(long)y+height)throw new InvalidOperationException("Owned physical ROI exceeds finite virtual-screen bounds.");
  Unoccluded(window,roi);
  using(var bitmap=new Bitmap(state.Width,state.Height,PixelFormat.Format32bppArgb))using(var graphics=Graphics.FromImage(bitmap)){
   graphics.CopyFromScreen(state.Left,state.Top,0,0,new Size(state.Width,state.Height),CopyPixelOperation.SourceCopy);state.Returned=true;
   if(Owner(pid,birthUtc,imagePath,window)!=state.OwnerThreadId||GetAncestor(GetForegroundWindow(),2)!=window||!Equal(roi,Bounds(window)))throw new InvalidOperationException("Owner, foreground or physical ROI changed during screen capture.");Unoccluded(window,roi);state.ForegroundVerifiedAfter=true;state.InputAfter=Input();if(state.InputAfter!=state.InputBefore)throw new InvalidOperationException("Input desktop changed.");
   using(var output=new FileStream(outputPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read)){bitmap.Save(output,ImageFormat.Png);output.Flush(true);state.Bytes=output.Length;}
  }
 }catch(Exception e){state.Error=e.ToString();}finally{if(previousDpi!=IntPtr.Zero){IntPtr restored=SetThreadDpiAwarenessContext(previousDpi);state.DpiRestored=restored!=IntPtr.Zero&&AreDpiAwarenessContextsEqual(GetThreadDpiAwarenessContext(),previousDpi);if(!state.DpiRestored)state.Error=(state.Error??"")+" | Thread DPI restore unproved.";}state.FinishedUtc=DateTime.UtcNow.ToString("o");state.Completed=true;}});thread.IsBackground=true;thread.SetApartmentState(ApartmentState.MTA);thread.Start();return state;}
}
