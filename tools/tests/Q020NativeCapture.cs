using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

// Read-only pixel evidence from one explicitly owned VBE window on an inactive desktop.
// PrintWindow can block inside the host. The caller must impose a deadline, record
// uncertainty, retain its original host handles, and never replay a pending capture.
public static class Q020NativeCapture
{
    public sealed class CaptureState
    {
        public volatile bool Completed;
        public bool Returned;
        public string Error, OutputPath, Desktop, ThreadDesktop, InputBefore, InputAfter, StartedUtc, FinishedUtc;
        public int Width, Height, HostPid;
        public uint OwnerThreadId;
        public long Window, Bytes;
    }
    [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll", SetLastError=true)] static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("user32.dll", SetLastError=true)] static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool GetUserObjectInformation(IntPtr obj, int index, StringBuilder value, uint bytes, out uint needed);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder value, int count);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll", SetLastError=true)] static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll", SetLastError=true)] static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

    static string DesktopName(IntPtr desktop)
    {
        var value=new StringBuilder(256); uint needed;
        if(!GetUserObjectInformation(desktop, 2, value, 512, out needed)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        return value.ToString();
    }
    static string InputName()
    {
        IntPtr input=OpenInputDesktop(0, false, 1);
        if(input==IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        try { return DesktopName(input); } finally { CloseDesktop(input); }
    }
    static uint RequireOwner(int pid, string birth, string image, string desktop, IntPtr window)
    {
        using(var host=Process.GetProcessById(pid))
        {
            if(host.HasExited || host.StartTime.ToUniversalTime().ToString("o")!=birth || !string.Equals(host.MainModule.FileName,image,StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Original capture host identity changed.");
        }
        uint actual; uint tid=GetWindowThreadProcessId(window,out actual);
        var name=new StringBuilder(256); GetClassName(window,name,name.Capacity);
        if(tid==0 || actual!=(uint)pid || name.ToString()!="wndclass_desked_gsk" || GetAncestor(window,2)!=window || !IsWindowVisible(window) || !IsWindowEnabled(window))
            throw new InvalidOperationException("Exact visible enabled VBE capture root changed.");
        if(InputName()==desktop) throw new InvalidOperationException("Qualification desktop became the input desktop.");
        return tid;
    }
    public static CaptureState Begin(int pid, string birthUtc, string imagePath, string desktop, long rootWindow, string outputPath)
    {
        if(pid<=0 || rootWindow<=0 || !Regex.IsMatch(desktop??"", @"^VBAiTests_[0-9a-f]{32}$") || !Path.IsPathRooted(imagePath) || !Path.IsPathRooted(outputPath) || File.Exists(outputPath) || !Directory.Exists(Path.GetDirectoryName(outputPath)))
            throw new ArgumentException("Fresh explicit owned capture arguments required.");
        var state=new CaptureState {HostPid=pid,Window=rootWindow,Desktop=desktop,OutputPath=outputPath,StartedUtc=DateTime.UtcNow.ToString("o")};
        var thread=new Thread(()=> {
            try
            {
                // Native-only MTA work inherits the private-born process desktop.
                // GetThreadDesktop returns a borrowed handle: never rebind or close it.
                IntPtr inherited=GetThreadDesktop(GetCurrentThreadId());
                if(inherited==IntPtr.Zero) throw new InvalidOperationException("Capture thread desktop unavailable.");
                state.ThreadDesktop=DesktopName(inherited);
                if(state.ThreadDesktop!=desktop) throw new InvalidOperationException("Capture inherited desktop differs.");
                state.InputBefore=InputName();
                IntPtr window=new IntPtr(rootWindow);
                state.OwnerThreadId=RequireOwner(pid,birthUtc,imagePath,desktop,window);
                Rect rect;
                if(!GetWindowRect(window,out rect)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                state.Width=rect.Right-rect.Left; state.Height=rect.Bottom-rect.Top;
                if(state.Width<64 || state.Height<64 || state.Width>8192 || state.Height>8192 || (long)state.Width*state.Height>16777216)
                    throw new InvalidOperationException("Capture geometry exceeds the finite evidence budget.");
                using(var bitmap=new Bitmap(state.Width,state.Height,PixelFormat.Format32bppArgb))
                using(var graphics=Graphics.FromImage(bitmap))
                {
                    IntPtr dc=graphics.GetHdc();
                    try { state.Returned=PrintWindow(window,dc,2); }
                    finally { graphics.ReleaseHdc(dc); }
                    if(!state.Returned) throw new InvalidOperationException("The one native PrintWindow call returned false.");
                    if(RequireOwner(pid,birthUtc,imagePath,desktop,window)!=state.OwnerThreadId) throw new InvalidOperationException("Capture owner thread changed.");
                    state.InputAfter=InputName();
                    if(state.InputAfter!=state.InputBefore) throw new InvalidOperationException("Input desktop changed during capture.");
                    using(var output=new FileStream(outputPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read))
                    {bitmap.Save(output,ImageFormat.Png);output.Flush(true);state.Bytes=output.Length;}
                }
            }
            catch(Exception error) { state.Error=error.ToString(); }
            finally
            {
                state.FinishedUtc=DateTime.UtcNow.ToString("o");state.Completed=true;
            }
        });
        thread.IsBackground=true; thread.SetApartmentState(ApartmentState.MTA); thread.Start();
        return state;
    }
}
