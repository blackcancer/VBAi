using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;
using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

namespace VBAi.Tests.Integration
{
    /// <summary>One-shot diagnostic PROPERTYPUT for only the owned HelpFile/HelpContextID scalar contracts.</summary>
    internal static class NativeMetadataSetterProbe
    {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int InvokeDelegate(IntPtr self, int id, ref Guid iid, uint locale, ushort flags,
            ref DISPPARAMS arguments, IntPtr result, ref EXCEPINFO exception, out uint argumentError);

        internal static IDictionary<string, object> PutRaw(object target, string property, object value)
        {
            Validate(property, value);
            IntPtr dispatch = Marshal.GetIDispatchForObject(target);
            try
            {
                var invoke = (InvokeDelegate)Marshal.GetDelegateForFunctionPointer(
                    Marshal.ReadIntPtr(Marshal.ReadIntPtr(dispatch), 6 * IntPtr.Size), typeof(InvokeDelegate));
                var exception = new EXCEPINFO(); uint argumentError = 0;
                var row = PutRawCore(property, value, (parameters, result) => {
                    Guid iid = Guid.Empty;
                    return invoke(dispatch, property == "HelpFile" ? 116 : 117, ref iid,
                        (uint)CultureInfo.InvariantCulture.LCID, 4 /* DISPATCH_PROPERTYPUT */,
                        ref parameters, result, ref exception, out argumentError);
                });
                row["ExceptionScode"] = Hex(exception.scode);
                row["DeferredExceptionCallbackPresent"] = exception.pfnDeferredFillIn != IntPtr.Zero;
                var returned = (IDictionary<string, object>)row["ReturnVariant"];
                returned.TryGetValue("HResult", out object hr);
                row["ArgumentError"] = Equals(hr, "0x80020005") || Equals(hr, "0x80020004") ? (object)argumentError : null;
                return row;
            }
            finally { Marshal.Release(dispatch); }
        }

        /// <summary>Supplies one named PROPERTYPUT argument with full x64 argument/result allocations; never retries Invoke.</summary>
        internal static IDictionary<string, object> PutRawCore(string property, object value, Func<DISPPARAMS, IntPtr, int> invoke)
        {
            Validate(property, value);
            var row = new Dictionary<string, object> { ["Property"] = property, ["Dispid"] = property == "HelpFile" ? 116 : 117,
                ["Flags"] = "DISPATCH_PROPERTYPUT only", ["Locale"] = CultureInfo.InvariantCulture.LCID,
                ["NamedDispid"] = -3, ["ArgumentCount"] = 1, ["NamedArgumentCount"] = 1,
                ["InvokeEntries"] = 0, ["MutationRetryAllowed"] = false,
                ["ReturnVariant"] = new Dictionary<string, object> { ["State"] = "NOT_CALLED" } };
            IntPtr named = Marshal.AllocHGlobal(sizeof(int));
            try
            {
                Marshal.WriteInt32(named, -3); // DISPID_PROPERTYPUT, required by the IDispatch contract.
                row["ArgumentVariant"] = NativeMetadataGetterProbe.ReadVariant(argument => {
                    Marshal.GetNativeVariantForObject(value, argument);
                    var parameters = new DISPPARAMS { rgvarg = argument, rgdispidNamedArgs = named, cArgs = 1, cNamedArgs = 1 };
                    row["ReturnVariant"] = NativeMetadataGetterProbe.ReadVariant(result => {
                        row["InvokeEntries"] = 1; // Entry attempted, not evidence that the host changed the property.
                        return invoke(parameters, result);
                    });
                    return 0; // Input-buffer observation HRESULT, not the native setter HRESULT.
                });
                row["NamedDispidAfterInvoke"] = Marshal.ReadInt32(named);
            }
            finally { Marshal.FreeHGlobal(named); }
            return row;
        }

        private static void Validate(string property, object value)
        {
            if (IntPtr.Size != 8) throw new PlatformNotSupportedException("x64 diagnostic ABI required.");
            if ((property != "HelpFile" || !(value is string text) || text.Length > 4096) &&
                (property != "HelpContextID" || !(value is int)))
                throw new ArgumentException("Only bounded HelpFile BSTR or HelpContextID I4 values are allowed; no coercion is performed.");
        }

        private static string Hex(int value) => "0x" + unchecked((uint)value).ToString("X8");
    }
}
