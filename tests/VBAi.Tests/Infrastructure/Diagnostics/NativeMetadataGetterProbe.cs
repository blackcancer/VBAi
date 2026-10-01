using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;
using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;

namespace VBAi.Tests.Integration
{
    /// <summary>Reads only HelpFile/HelpContextID through three getters and the same object's runtime type information.</summary>
    internal static class NativeMetadataGetterProbe
    {
        internal const int VariantBytes = 24; // x64 VARIANT includes its two-pointer BRECORD union.
        private const long Canary = 0x1837462518374625;
        [DllImport("oleaut32.dll", PreserveSig = true)] private static extern int VariantClear(IntPtr variant);
        [DllImport("oleaut32.dll")] private static extern uint SysStringLen(IntPtr value);
        [DllImport("oleaut32.dll")] private static extern uint SysStringByteLen(IntPtr value);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetTypeInfoCountDelegate(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetTypeInfoDelegate(IntPtr self, uint index, uint locale, out IntPtr info);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int InvokeDelegate(IntPtr self, int id, ref Guid iid, uint locale, ushort flags,
            ref DISPPARAMS arguments, IntPtr result, ref EXCEPINFO exception, out uint argumentError);

        internal static IDictionary<string, object> Read(object target, Action<string> pending,
            Action<IDictionary<string, object>> begin = null)
        {
            if (IntPtr.Size != 8) throw new PlatformNotSupportedException("This diagnostic requires the x64 VARIANT ABI.");
            var report = new Dictionary<string, object> { ["VariantBufferBytes"] = VariantBytes,
                ["DispatchFlags"] = "DISPATCH_PROPERTYGET only", ["NativeWrites"] = 0,
                ["RawLocale"] = 0, ["ClrBinderCulture"] = "InvariantCulture" };
            begin?.Invoke(report);
            pending("AcquireIDispatch");
            IntPtr dispatch = Marshal.GetIDispatchForObject(target);
            try
            {
                pending("RuntimeTypeInfo");
                try { report["RuntimeTypeInfo"] = ReadTypeInfo(dispatch); }
                catch (Exception error) { report["RuntimeTypeInfo"] = new Dictionary<string, object> {
                    ["State"] = "ERROR", ["Error"] = error.Message, ["HResult"] = Hex(error.HResult) }; }
                var values = new List<object>(); report["Properties"] = values;
                foreach (var member in new[] { new { Name = "HelpFile", Id = 116 }, new { Name = "HelpContextID", Id = 117 } })
                {
                    var row = new Dictionary<string, object> { ["Name"] = member.Name, ["Dispid"] = member.Id };
                    values.Add(row);
                    pending(member.Name + ".Descriptor");
                    row["Descriptor"] = Observe(() => {
                        var descriptor = TypeDescriptor.GetProperties(target).Find(member.Name, false);
                        if (descriptor == null) throw new MissingMemberException(member.Name);
                        row["DescriptorType"] = descriptor.PropertyType.FullName;
                        row["DescriptorReadOnly"] = descriptor.IsReadOnly;
                        return descriptor.GetValue(target);
                    });
                    pending(member.Name + ".ClrBinder");
                    row["ClrBinder"] = Observe(() => target.GetType().InvokeMember(member.Name,
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, target, null, CultureInfo.InvariantCulture));
                    pending(member.Name + ".RawDispatch");
                    var invoke = Slot<InvokeDelegate>(dispatch, 6);
                    row["RawDispatch"] = ReadVariant(pointer => {
                        var parameters = new DISPPARAMS(); var exception = new EXCEPINFO();
                        Guid iid = Guid.Empty; uint argument;
                        int result = invoke(dispatch, member.Id, ref iid, 0, 2 /* PROPERTYGET */,
                            ref parameters, pointer, ref exception, out argument);
                        row["RawInvokeArgumentError"] = result == unchecked((int)0x80020005) ? (object)argument : null;
                        row["RawExceptionScode"] = Hex(exception.scode);
                        row["DeferredExceptionCallbackPresent"] = exception.pfnDeferredFillIn != IntPtr.Zero;
                        // Never execute a deferred callback or a setter/method from the probe.
                        return result;
                    });
                    row["ExactGetterEquality"] = EqualGetters(row);
                }
                return report;
            }
            finally { Marshal.Release(dispatch); }
        }

        private static object Observe(Func<object> read)
        {
            try
            {
                object value = read();
                if (value != null && !(value is string) && !(value is int))
                    throw new InvalidOperationException("Unexpected non-I4/BSTR metadata getter type.");
                if (value is string text && text.Length > 4096)
                    throw new InvalidOperationException("Owned metadata string exceeds the bounded diagnostic read.");
                return new Dictionary<string, object> { ["State"] = "READ", ["Value"] = value,
                    ["ClrType"] = value?.GetType().FullName, ["StringLength"] = (value as string)?.Length };
            }
            catch (Exception error)
            {
                if (error is TargetInvocationException invocation && invocation.InnerException != null) error = invocation.InnerException;
                return new Dictionary<string, object> { ["State"] = "ERROR", ["Error"] = error.Message,
                    ["ExceptionType"] = error.GetType().FullName, ["HResult"] = Hex(error.HResult) };
            }
        }

        /// <summary>Uses a 24-byte result plus a checked canary; clears OLE-owned result data on every outcome.</summary>
        internal static IDictionary<string, object> ReadVariant(Func<IntPtr, int> invoke)
        {
            if (IntPtr.Size != 8) throw new PlatformNotSupportedException("x64 result ABI required.");
            var result = new Dictionary<string, object> { ["State"] = "PENDING", ["BufferBytes"] = VariantBytes };
            IntPtr buffer = Marshal.AllocHGlobal(VariantBytes + sizeof(long));
            Marshal.Copy(new byte[VariantBytes], 0, buffer, VariantBytes);
            Marshal.WriteInt64(buffer, VariantBytes, Canary);
            try
            {
                int hr = invoke(buffer);
                result["HResult"] = Hex(hr);
                result["CanaryIntact"] = Marshal.ReadInt64(buffer, VariantBytes) == Canary;
                if (!(bool)result["CanaryIntact"]) throw new InvalidOperationException("Native result crossed the x64 VARIANT boundary.");
                ushort type = unchecked((ushort)Marshal.ReadInt16(buffer));
                result["VariantType"] = type; result["VariantTypeName"] = ((VarEnum)type).ToString();
                if (hr < 0) { result["State"] = "INVOKE_FAILED"; return result; }
                if (type != 0 && type != 1 && type != 3 && type != 8)
                    throw new InvalidOperationException("Unexpected result VARTYPE; no BYREF/array/object getter dereference is attempted.");
                if (type == 8)
                {
                    IntPtr bstr = Marshal.ReadIntPtr(buffer, 8);
                    uint bytes = bstr == IntPtr.Zero ? 0 : SysStringByteLen(bstr);
                    uint chars = bstr == IntPtr.Zero ? 0 : SysStringLen(bstr);
                    result["BstrByteLength"] = bytes; result["BstrCharLength"] = chars;
                    if (bytes > 8192) throw new InvalidOperationException("Owned metadata BSTR exceeds the bounded diagnostic read.");
                    byte[] raw = new byte[bytes]; if (bytes != 0) Marshal.Copy(bstr, raw, 0, raw.Length);
                    result["BstrBytesHex"] = BitConverter.ToString(raw).Replace("-", "");
                }
                result["Value"] = Marshal.GetObjectForNativeVariant(buffer);
                result["State"] = "READ";
            }
            catch (Exception error) { result["State"] = "ERROR"; result["Error"] = error.Message; result["ReadErrorHResult"] = Hex(error.HResult); }
            finally
            {
                int clear = VariantClear(buffer);
                result["VariantClearHResult"] = Hex(clear);
                Marshal.FreeHGlobal(buffer);
                if (clear < 0) { result["BeforeClearFailureState"] = result["State"]; result["State"] = "CLEAR_FAILED"; }
            }
            return result;
        }

        private static bool EqualGetters(IDictionary<string, object> row)
        {
            var descriptor = (IDictionary<string, object>)row["Descriptor"];
            var binder = (IDictionary<string, object>)row["ClrBinder"];
            var raw = (IDictionary<string, object>)row["RawDispatch"];
            return Equals(descriptor["State"], "READ") && Equals(binder["State"], "READ") && Equals(raw["State"], "READ") &&
                Equals(descriptor["Value"], binder["Value"]) && Equals(binder["Value"], raw["Value"]);
        }

        private static IDictionary<string, object> ReadTypeInfo(IntPtr dispatch)
        {
            var result = new Dictionary<string, object> { ["State"] = "PENDING" };
            uint count; Marshal.ThrowExceptionForHR(Slot<GetTypeInfoCountDelegate>(dispatch, 3)(dispatch, out count));
            result["Count"] = count;
            if (count != 1) throw new InvalidOperationException("One runtime dispatch type-info entry is required.");
            IntPtr pointer; Marshal.ThrowExceptionForHR(Slot<GetTypeInfoDelegate>(dispatch, 4)(dispatch, 0, 0, out pointer));
            ITypeInfo info = null; ITypeLib library = null;
            try
            {
                info = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(pointer, typeof(ITypeInfo));
                IntPtr attribute; info.GetTypeAttr(out attribute);
                try
                {
                    var type = (TYPEATTR)Marshal.PtrToStructure(attribute, typeof(TYPEATTR));
                    result["Guid"] = type.guid.ToString("D"); result["TypeKind"] = type.typekind.ToString();
                    if (type.cFuncs < 0 || type.cFuncs > 256) throw new InvalidOperationException("Runtime member count exceeds the bounded read.");
                    var members = new List<object>(); result["Members"] = members;
                    var observed = new HashSet<string>(StringComparer.Ordinal);
                    for (int index = 0; index < type.cFuncs; index++)
                    {
                        IntPtr function; info.GetFuncDesc(index, out function);
                        try
                        {
                            var entry = (FUNCDESC)Marshal.PtrToStructure(function, typeof(FUNCDESC));
                            if (entry.memid != 116 && entry.memid != 117) continue;
                            string name, documentation, help; int context;
                            info.GetDocumentation(entry.memid, out name, out documentation, out context, out help);
                            if (entry.cParams < 0 || entry.cParams > 4) throw new InvalidOperationException("Unexpected metadata parameter count.");
                            var parameters = new List<object>();
                            for (int parameter = 0; parameter < entry.cParams; parameter++)
                            {
                                var item = (ELEMDESC)Marshal.PtrToStructure(IntPtr.Add(entry.lprgelemdescParam,
                                    parameter * Marshal.SizeOf(typeof(ELEMDESC))), typeof(ELEMDESC));
                                parameters.Add(new { Type = DescribeType(item.tdesc, 0), Flags = item.desc.paramdesc.wParamFlags.ToString() });
                            }
                            members.Add(new { Name = name, Dispid = entry.memid, InvokeKind = entry.invkind.ToString(),
                                ReturnType = DescribeType(entry.elemdescFunc.tdesc, 0), Parameters = parameters });
                            observed.Add(entry.memid + ":" + entry.invkind);
                        }
                        finally { info.ReleaseFuncDesc(function); }
                    }
                    result["BothGettersAndSettersObserved"] = observed.Contains("116:INVOKE_PROPERTYGET") &&
                        observed.Contains("116:INVOKE_PROPERTYPUT") && observed.Contains("117:INVOKE_PROPERTYGET") &&
                        observed.Contains("117:INVOKE_PROPERTYPUT");
                }
                finally { info.ReleaseTypeAttr(attribute); }
                int containingIndex; info.GetContainingTypeLib(out library, out containingIndex);
                IntPtr libraryAttribute; library.GetLibAttr(out libraryAttribute);
                try
                {
                    var type = (TYPELIBATTR)Marshal.PtrToStructure(libraryAttribute, typeof(TYPELIBATTR));
                    result["Library"] = new { Guid = type.guid.ToString("D"), Major = type.wMajorVerNum,
                        Minor = type.wMinorVerNum, SystemKind = type.syskind.ToString(), Locale = type.lcid, Index = containingIndex };
                }
                finally { library.ReleaseTLibAttr(libraryAttribute); }
                result["State"] = "READ";
                return result;
            }
            finally
            {
                if (library != null) Marshal.ReleaseComObject(library);
                if (info != null) Marshal.ReleaseComObject(info);
                if (pointer != IntPtr.Zero) Marshal.Release(pointer);
            }
        }

        private static string DescribeType(TYPEDESC type, int depth)
        {
            if (depth > 3) throw new InvalidOperationException("Metadata type indirection exceeds the bounded read.");
            return ((VarEnum)type.vt).ToString() + (type.vt == 26 ? "->" + DescribeType(
                (TYPEDESC)Marshal.PtrToStructure(type.lpValue, typeof(TYPEDESC)), depth + 1) : "");
        }

        private static T Slot<T>(IntPtr dispatch, int index) where T : class =>
            (T)(object)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(dispatch), index * IntPtr.Size), typeof(T));
        private static string Hex(int value) => "0x" + unchecked((uint)value).ToString("X8");
    }
}
