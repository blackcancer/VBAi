using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
internal static class OwnedTeardownTraceHelper
{
    [StructLayout(LayoutKind.Sequential)]
    private struct ExceptionRecord
    {
        public uint Code, Flags;
        public IntPtr Record, Address;
        public uint NumberParameters;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst=15)] public UIntPtr[] Information;
    }
    [DllImport("kernel32.dll")] private static extern void RaiseFailFastException(ref ExceptionRecord record, IntPtr context, uint flags);
    private static void Main()
    {
        using (var process=Process.GetCurrentProcess())
            Console.WriteLine("READY " + process.Id + " " + process.StartTime.ToUniversalTime().ToString("o"));
        string command=Console.ReadLine();
        if(command=="STOP"){Console.WriteLine("STOPPED");return;}
        if(command!="FAILFAST")throw new InvalidOperationException("Only one owned STOP or synthetic FAILFAST is allowed.");
        var record=new ExceptionRecord{Code=0xC0000409,Flags=1,NumberParameters=1,Information=new UIntPtr[15]};
        record.Information[0]=new UIntPtr(7); // Synthetic fatal-app-exit, not a reproduction of the Excel cause.
        RaiseFailFastException(ref record,IntPtr.Zero,0);
    }
}
