param([string]$Path = (Join-Path $PSScriptRoot '../../bin/Debug/net48/CodexVBE.dll'))
$ErrorActionPreference = 'Stop'
Add-Type @"
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
public static class NativeThemeFileUsers {
 [StructLayout(LayoutKind.Sequential)] public struct UniqueProcess { public uint Id; public System.Runtime.InteropServices.ComTypes.FILETIME StartTime; }
 [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] public struct ProcessInfo {
  public UniqueProcess Process;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string Name;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=64)] public string Service;
  public int Type; public uint Status, Session;
  [MarshalAs(UnmanagedType.Bool)] public bool Restartable;
 }
 [DllImport("rstrtmgr.dll",CharSet=CharSet.Unicode)] static extern int RmStartSession(out uint session, uint flags, StringBuilder key);
 [DllImport("rstrtmgr.dll",CharSet=CharSet.Unicode)] static extern int RmRegisterResources(uint session,uint count,string[] paths,uint apps,IntPtr processes,uint services,IntPtr names);
 [DllImport("rstrtmgr.dll")] static extern int RmGetList(uint session,out uint needed,ref uint count,[In,Out] ProcessInfo[] data,ref uint reasons);
 [DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint session);
 public static ProcessInfo[] Read(string path) {
  uint session; int error=RmStartSession(out session,0,new StringBuilder(33));
  if(error!=0) throw new Win32Exception(error);
  try {
   error=RmRegisterResources(session,1,new[]{path},0,IntPtr.Zero,0,IntPtr.Zero);
   if(error!=0) throw new Win32Exception(error);
   uint needed=0,count=0,reasons=0; ProcessInfo[] rows=null;
   for(int attempt=0;attempt<5;attempt++) {
    error=RmGetList(session,out needed,ref count,rows,ref reasons);
    if(error==0) { if(rows==null) return new ProcessInfo[0]; Array.Resize(ref rows,(int)count); return rows; }
    if(error!=234) throw new Win32Exception(error);
    count=needed; rows=new ProcessInfo[count];
   }
   throw new InvalidOperationException("File users changed during enumeration.");
  } finally { RmEndSession(session); }
 }
}
"@
[NativeThemeFileUsers]::Read([IO.Path]::GetFullPath($Path)) | ForEach-Object {
 [pscustomobject]@{ Id=$_.Process.Id; Name=$_.Name; Restartable=$_.Restartable }
}
