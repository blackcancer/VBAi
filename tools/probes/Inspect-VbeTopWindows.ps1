param([Parameter(Mandatory=$true)][int]$HostProcessId)
$ErrorActionPreference='Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VbeTopWindows {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc,IntPtr data);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  public static string[] Read(int pid){var result=new List<string>();EnumWindows((hwnd,data)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner==pid){var title=new StringBuilder(250);GetWindowText(hwnd,title,title.Capacity);var kind=new StringBuilder(100);GetClassName(hwnd,kind,kind.Capacity);result.Add(hwnd+" "+kind+" Visible="+IsWindowVisible(hwnd)+" Title="+title);}return true;},IntPtr.Zero);return result.ToArray();}
}
'@
[VbeTopWindows]::Read($HostProcessId)
