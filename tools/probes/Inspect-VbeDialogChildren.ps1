param([Parameter(Mandatory=$true)][long]$WindowHandle)
$ErrorActionPreference='Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VbeDialogChildren {
  public delegate bool EnumProc(IntPtr hwnd,IntPtr data);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent,EnumProc proc,IntPtr data);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  public static string[] Read(IntPtr dialog){var rows=new List<string>();EnumChildWindows(dialog,(hwnd,data)=>{var title=new StringBuilder(250);GetWindowText(hwnd,title,title.Capacity);var kind=new StringBuilder(100);GetClassName(hwnd,kind,kind.Capacity);rows.Add(hwnd+" "+kind+" Visible="+IsWindowVisible(hwnd)+" Title="+title);return true;},IntPtr.Zero);return rows.ToArray();}
}
'@
[VbeDialogChildren]::Read([IntPtr]$WindowHandle)
