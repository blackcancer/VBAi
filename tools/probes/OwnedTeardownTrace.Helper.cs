using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
internal static class OwnedTeardownTraceHelper
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct ExceptionRecord
    {
        public uint Code, Flags;
        public IntPtr Record, Address;
        public uint NumberParameters;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst=15)] public UIntPtr[] Information;
    }
    // Windows SDK winnt.h AMD64 CONTEXT: size 1232, 16-byte alignment.
    [StructLayout(LayoutKind.Explicit, Size=1232)]
    internal struct Context64
    {
        [FieldOffset(48)] public uint Flags;
        [FieldOffset(152)] public ulong StackPointer;
        [FieldOffset(248)] public ulong InstructionPointer;
    }
    [DllImport("kernel32.dll")] private static extern void RtlCaptureContext(IntPtr context);
    [DllImport("kernel32.dll")] private static extern void RaiseFailFastException(ref ExceptionRecord record, IntPtr context, uint flags);
    internal static IntPtr AlignContext(IntPtr allocation)
    {
        if(IntPtr.Size!=8 || allocation==IntPtr.Zero)throw new InvalidOperationException("Owned helper requires an allocated x64 context.");
        return new IntPtr(checked(allocation.ToInt64()+15)&~15L);
    }
    internal static ExceptionRecord CreateRecord(IntPtr context)
    {
        if(IntPtr.Size!=8 || context==IntPtr.Zero || (context.ToInt64()&15)!=0)
            throw new InvalidOperationException("Captured x64 context must be non-null and aligned.");
        var captured=(Context64)Marshal.PtrToStructure(context,typeof(Context64));
        if((captured.Flags&0x100003)!=0x100003 || captured.InstructionPointer==0 || captured.StackPointer==0)
            throw new InvalidOperationException("Captured x64 control/integer registers are incomplete.");
        var record=new ExceptionRecord{Code=0xC0000409,Flags=1,Address=new IntPtr(unchecked((long)captured.InstructionPointer)),NumberParameters=1,Information=new UIntPtr[15]};
        record.Information[0]=new UIntPtr(7); // Synthetic fatal-app-exit, not a reproduction of the Excel cause.
        return record;
    }
    private static void Main()
    {
        using (var process=Process.GetCurrentProcess())
            Console.WriteLine("READY " + process.Id + " " + process.StartTime.ToUniversalTime().ToString("o"));
        string command=Console.ReadLine();
        if(command=="STOP"){Console.WriteLine("STOPPED");return;}
        if(command!="FAILFAST")throw new InvalidOperationException("Only one owned STOP or synthetic FAILFAST is allowed.");
        IntPtr allocation=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Context64))+15);
        try
        {
            IntPtr context=AlignContext(allocation);
            Marshal.Copy(new byte[Marshal.SizeOf(typeof(Context64))],0,context,Marshal.SizeOf(typeof(Context64)));
            RtlCaptureContext(context);
            var record=CreateRecord(context);
            RaiseFailFastException(ref record,context,0);
            throw new InvalidOperationException("Fatal exception unexpectedly returned; preflight must not pass.");
        }
        finally{Marshal.FreeHGlobal(allocation);}
    }
}
