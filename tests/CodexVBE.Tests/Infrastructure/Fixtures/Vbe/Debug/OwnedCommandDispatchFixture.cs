using System;
using System.Runtime.InteropServices;

namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>An owned native IDispatch identity, backed by local command state; never an Office object.</summary>
    internal sealed class OwnedCommandDispatchFixture : IDisposable
    {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Query(IntPtr self, ref Guid iid, out IntPtr result);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate uint Reference(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int TypeCount(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int TypeInfo(IntPtr self, uint index, uint locale, out IntPtr info);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Names(IntPtr self, ref Guid iid, IntPtr names, uint count, uint locale, IntPtr ids);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Invoke(IntPtr self, int id, ref Guid iid, uint locale, ushort flags, IntPtr arguments, IntPtr result, IntPtr exception, IntPtr error);
        private readonly Delegate[] methods;
        private readonly IntPtr table, identity;
        private int references = 1;
        internal readonly object Control;
        internal int Id = 2578, Type = 1, Executions;
        internal string Caption = "Project Properties...";
        internal bool Enabled = true;
        internal OwnedCommandDispatchFixture()
        {
            methods = new Delegate[] { new Query(QueryInterface), new Reference(_ => (uint)System.Threading.Interlocked.Increment(ref references)), new Reference(_ => (uint)System.Threading.Interlocked.Decrement(ref references)),
                new TypeCount(GetTypeCount), new TypeInfo(GetTypeInfo), new Names(GetNames), new Invoke(Dispatch) };
            table = Marshal.AllocHGlobal(IntPtr.Size * methods.Length);
            for (int i = 0; i < methods.Length; i++) Marshal.WriteIntPtr(table, i * IntPtr.Size, Marshal.GetFunctionPointerForDelegate(methods[i]));
            identity = Marshal.AllocHGlobal(IntPtr.Size); Marshal.WriteIntPtr(identity, table);
            Control = Marshal.GetObjectForIUnknown(identity);
            if (!Marshal.IsComObject(Control)) throw new InvalidOperationException("The owned identity must be a native RCW.");
        }
        private int QueryInterface(IntPtr self, ref Guid iid, out IntPtr result)
        {
            result = iid == new Guid("00000000-0000-0000-C000-000000000046") || iid == new Guid("00020400-0000-0000-C000-000000000046") ? identity : IntPtr.Zero;
            if (result != IntPtr.Zero) System.Threading.Interlocked.Increment(ref references);
            return result == IntPtr.Zero ? unchecked((int)0x80004002) : 0;
        }
        private static int GetTypeCount(IntPtr self, out uint count) { count = 0; return 0; }
        private static int GetTypeInfo(IntPtr self, uint index, uint locale, out IntPtr info) { info = IntPtr.Zero; return unchecked((int)0x8002000B); }
        private static int GetNames(IntPtr self, ref Guid iid, IntPtr names, uint count, uint locale, IntPtr ids)
        {
            for (int i = 0; i < count; i++)
            {
                string name = Marshal.PtrToStringUni(Marshal.ReadIntPtr(names, i * IntPtr.Size));
                int id = string.Equals(name, "Id", StringComparison.OrdinalIgnoreCase) ? 1 : string.Equals(name, "Caption", StringComparison.OrdinalIgnoreCase) ? 2 :
                    string.Equals(name, "Enabled", StringComparison.OrdinalIgnoreCase) ? 3 : string.Equals(name, "Type", StringComparison.OrdinalIgnoreCase) ? 4 :
                    string.Equals(name, "Execute", StringComparison.OrdinalIgnoreCase) ? 5 : -1;
                Marshal.WriteInt32(ids, i * 4, id); if (id < 0) return unchecked((int)0x80020006);
            }
            return 0;
        }
        private int Dispatch(IntPtr self, int id, ref Guid iid, uint locale, ushort flags, IntPtr arguments, IntPtr result, IntPtr exception, IntPtr error)
        {
            if (id == 5) { Executions++; return 0; }
            object value = id == 1 ? (object)Id : id == 2 ? Caption : id == 3 ? (object)Enabled : id == 4 ? (object)Type : null;
            if (value == null) return unchecked((int)0x80020003);
            if (result != IntPtr.Zero) Marshal.GetNativeVariantForObject(value, result);
            return 0;
        }
        public void Dispose()
        {
            Marshal.FinalReleaseComObject(Control); Marshal.Release(identity); Marshal.FreeHGlobal(identity); Marshal.FreeHGlobal(table); GC.KeepAlive(methods);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System.Collections.Generic;
    public sealed partial class VbeProjectProtectionTests
    {
        public sealed class DialogHost
        {
            public List<object> VBProjects { get; } = new List<object>();
            public object ActiveVBProject { get; set; }
            public List<DialogBar> CommandBars { get; } = new List<DialogBar>();
        }
        public sealed class DialogBar
        {
            public string Name { get; set; } = "Tools";
            public List<object> Controls { get; } = new List<object>();
        }
    }
}
