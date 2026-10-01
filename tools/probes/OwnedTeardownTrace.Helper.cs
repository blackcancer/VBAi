using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Security;
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
    [DllImport("kernel32.dll")] private static extern void RaiseException(uint code, uint flags, uint count, UIntPtr[] arguments);
    internal static IntPtr AlignContext(IntPtr allocation)
    {
        if(IntPtr.Size!=8 || allocation==IntPtr.Zero)throw new InvalidOperationException("Owned helper requires an allocated x64 context.");
        return new IntPtr(checked(allocation.ToInt64()+15)&~15L);
    }
    internal static ExceptionRecord CreateRecord(IntPtr context, uint code=0xC0000409)
    {
        if(IntPtr.Size!=8 || context==IntPtr.Zero || (context.ToInt64()&15)!=0)
            throw new InvalidOperationException("Captured x64 context must be non-null and aligned.");
        var captured=(Context64)Marshal.PtrToStructure(context,typeof(Context64));
        if((captured.Flags&0x100003)!=0x100003 || captured.InstructionPointer==0 || captured.StackPointer==0)
            throw new InvalidOperationException("Captured x64 control/integer registers are incomplete.");
        if(code!=0xC0000409 && code!=0xC0000005)throw new InvalidOperationException("Only the two synthetic fatal codes are allowed.");
        var record=new ExceptionRecord{Code=code,Flags=1,Address=new IntPtr(unchecked((long)captured.InstructionPointer)),NumberParameters=code==0xC0000409?1U:2U,Information=new UIntPtr[15]};
        // Software-generated diagnostics, never a reproduction of the Excel cause or a real memory write.
        if(code==0xC0000409)record.Information[0]=new UIntPtr(7);
        else record.Information[1]=new UIntPtr(1);
        return record;
    }
    [HandleProcessCorruptedStateExceptions, SecurityCritical]
    private static void HandledFirstChanceAv()
    {
        Console.WriteLine("FIRSTCHANCE_AV_REQUESTED");
        try{RaiseException(0xC0000005,0,2,new[]{UIntPtr.Zero,new UIntPtr(1)});}
        catch(AccessViolationException){Console.WriteLine("FIRSTCHANCE_AV_HANDLED");return;}
        throw new InvalidOperationException("Synthetic AV unexpectedly returned without the local handler.");
    }
    private static void Main()
    {
        using (var process=Process.GetCurrentProcess())
            Console.WriteLine("READY " + process.Id + " " + process.StartTime.ToUniversalTime().ToString("o"));
        string command=Console.ReadLine();
        if(command=="STOP"){Console.WriteLine("STOPPED");return;}
        if(command=="FIRSTCHANCEAV"){HandledFirstChanceAv();return;}
        if(command!="FAILFAST" && command!="SECONDCHANCEAV")throw new InvalidOperationException("Only one owned STOP or named synthetic exception is allowed.");
        IntPtr allocation=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Context64))+15);
        try
        {
            IntPtr context=AlignContext(allocation);
            Marshal.Copy(new byte[Marshal.SizeOf(typeof(Context64))],0,context,Marshal.SizeOf(typeof(Context64)));
            RtlCaptureContext(context);
            var record=CreateRecord(context,command=="FAILFAST"?0xC0000409U:0xC0000005U);
            RaiseFailFastException(ref record,context,0);
            throw new InvalidOperationException("Fatal exception unexpectedly returned; preflight must not pass.");
        }
        finally{Marshal.FreeHGlobal(allocation);}
    }
}
