param([Parameter(Mandatory=$true)][int]$HostProcessId)
$ErrorActionPreference='Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class ChatTranscript {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr data);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder text,int size);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,StringBuilder output);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
  public static string Read(int pid) {
    IntPtr assistant=IntPtr.Zero;
    EnumWindows((hwnd,data)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner==pid){var title=new StringBuilder(200);GetWindowText(hwnd,title,title.Capacity);if(title.ToString().StartsWith("CodexVBE")&&title.ToString().Contains("Assistant"))assistant=hwnd;}return true;},IntPtr.Zero);
    if(assistant==IntPtr.Zero)throw new InvalidOperationException("Assistant absent");
    IntPtr transcript=IntPtr.Zero;int edits=0;
    EnumChildWindows(assistant,(hwnd,data)=>{var kind=new StringBuilder(100);GetClassName(hwnd,kind,kind.Capacity);if(kind.ToString().Contains(".EDIT.")){edits++;if(edits==1)transcript=hwnd;}return true;},IntPtr.Zero);
    if(transcript==IntPtr.Zero)throw new InvalidOperationException("Transcript absent");
    int length=SendMessage(transcript,0x000E,IntPtr.Zero,IntPtr.Zero).ToInt32();
    var buffer=new StringBuilder(length+1);SendMessage(transcript,0x000D,new IntPtr(buffer.Capacity),buffer);return buffer.ToString();
  }
}
'@
[ChatTranscript]::Read($HostProcessId)
