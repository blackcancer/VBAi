using System;using System.Collections.Generic;using System.Runtime.InteropServices;using System.Text;
public static class IsolatedHelperProcessNative {
 [DllImport("kernel32.dll",SetLastError=true)]static extern uint GetProcessId(IntPtr h);
 [DllImport("kernel32.dll",SetLastError=true)]static extern bool GetProcessTimes(IntPtr h,out long c,out long e,out long k,out long u);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)]static extern bool QueryFullProcessImageNameW(IntPtr h,uint flags,StringBuilder path,ref uint size);
 [DllImport("kernel32.dll",SetLastError=true)]public static extern uint WaitForSingleObject(IntPtr h,uint milliseconds);
 [DllImport("kernel32.dll",SetLastError=true)]static extern bool GetExitCodeProcess(IntPtr h,out uint code);
 public static Dictionary<string,object> Describe(IntPtr h,uint pid,string expected){
  long c,e,k,u;var image=new StringBuilder(32768);uint size=(uint)image.Capacity;
  if(h==IntPtr.Zero||GetProcessId(h)!=pid||!GetProcessTimes(h,out c,out e,out k,out u)||!QueryFullProcessImageNameW(h,0,image,ref size)||!string.Equals(image.ToString(),expected,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Original created helper handle/PID/birth/image refused.");
  return new Dictionary<string,object>{{"ProcessId",pid},{"OriginalProcessHandle",h.ToInt64()},{"StartedUtc",DateTime.FromFileTimeUtc(c).ToString("o")},{"Image",image.ToString()}};
 }
 public static uint ExitCode(IntPtr h){uint code;if(WaitForSingleObject(h,0)!=0||!GetExitCodeProcess(h,out code))throw new InvalidOperationException("Original held helper exit was not observed.");return code;}
}
