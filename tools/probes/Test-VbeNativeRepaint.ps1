param(
    [Parameter(Mandatory=$true)][int]$HostProcessId,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
public static class VbeRepaintProbe {
    private delegate bool EnumProc(IntPtr window, IntPtr unused);
    [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr unused);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect bounds);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);
    [DllImport("user32.dll")] private static extern bool RedrawWindow(IntPtr window, IntPtr rect, IntPtr region, uint flags);
    public static IntPtr Find(int processId) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((window, unused) => {
            uint process; GetWindowThreadProcessId(window, out process);
            var name=new StringBuilder(256); GetClassName(window,name,name.Capacity);
            if(process==processId && name.ToString()=="wndclass_desked_gsk" && IsWindowVisible(window)) { found=window; return false; }
            return true;
        },IntPtr.Zero);
        if(found==IntPtr.Zero) throw new InvalidOperationException("No visible VBE for the requested process.");
        return found;
    }
    public static void Capture(IntPtr window,string path,bool live) {
        Rect rect;
        if(!GetWindowRect(window,out rect)) throw new InvalidOperationException("Window bounds unavailable.");
        using(var bitmap=new Bitmap(rect.Right-rect.Left,rect.Bottom-rect.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) {
                if(live) graphics.CopyFromScreen(rect.Left,rect.Top,0,0,bitmap.Size);
                else {
                    IntPtr dc=graphics.GetHdc();
                    try { if(!PrintWindow(window,dc,2)) throw new InvalidOperationException("PrintWindow failed."); }
                    finally { graphics.ReleaseHdc(dc); }
                }
            }
            bitmap.Save(path,ImageFormat.Png);
        }
    }
    public static void Redraw(IntPtr window) {
        if(!RedrawWindow(window,IntPtr.Zero,IntPtr.Zero,0x585)) throw new InvalidOperationException("Native redraw failed.");
    }
}
'@ -ReferencedAssemblies System.Drawing
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
$window=[VbeRepaintProbe]::Find($HostProcessId)
[VbeRepaintProbe]::Capture($window,(Join-Path $OutputDirectory 'before-live.png'),$true)
[VbeRepaintProbe]::Redraw($window)
Start-Sleep -Milliseconds 350
[VbeRepaintProbe]::Capture($window,(Join-Path $OutputDirectory 'after-live.png'),$true)
[VbeRepaintProbe]::Capture($window,(Join-Path $OutputDirectory 'after-print.png'),$false)
@{ HostProcessId=$HostProcessId; Window=$window.ToInt64(); Operation='Native redraw only; no macro edit or execution.' } |
    ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'repaint.json')
