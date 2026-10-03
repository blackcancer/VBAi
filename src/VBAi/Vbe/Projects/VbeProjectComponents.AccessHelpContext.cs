using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        internal Func<bool> AccessHelpContextHost = () => string.Equals(Process.GetCurrentProcess().ProcessName, "MSACCESS", StringComparison.OrdinalIgnoreCase);
        internal Func<object, bool> AccessHelpContextNativeProject = Marshal.IsComObject;
        internal Func<object, object, bool> AccessHelpContextIdentity = SameAccessHelpContextIdentity;
        internal Func<object, IntPtr, AccessHelpContextDispatch> AccessHelpContextFactory = (project, window) =>
            new AccessHelpContextDispatch(new NativeAccessHelpContextCalls(project, window));

        private bool TrySetAccessHelpContext(Request request, object project)
        {
            if (!string.Equals(request.Property, "HelpContextID", StringComparison.OrdinalIgnoreCase) ||
                !AccessHelpContextHost() || !AccessHelpContextNativeProject(project)) return false;
            if (request.Value == null) throw new ArgumentException("Property and non-null Value are required.");
            var descriptor = TypeDescriptor.GetProperties(project).Find("HelpContextID", false);
            if (descriptor == null || descriptor.IsReadOnly || descriptor.PropertyType != typeof(int))
                throw new InvalidOperationException("The Access VBProject HelpContextID must expose a writable Int32 scalar.");
            int converted = (int)Convert.ChangeType(request.Value, typeof(int), CultureInfo.InvariantCulture);
            IntPtr window = new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd));
            Action<bool> requireTarget = revision => {
                request.RevalidateProjectPropertyAuthorization?.Invoke(true); // Scope validation may read/pump the host; perform it before target/version reads.
                if (!AccessHelpContextHost() || !AccessHelpContextNativeProject(project) ||
                    !AccessHelpContextIdentity(project, (object)GetDesignProject(request.Project)) ||
                    Convert.ToInt32(((dynamic)project).Protection) != 0 ||
                    new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)) != window)
                    throw new InvalidOperationException("The approved Access VBProject identity, mode, protection or VBE owner changed.");
                if (revision) AssertProjectVersion(request, project);
                // Project reads can pump messages; resolve the selector again before the final revision check.
                if (!AccessHelpContextHost() || !AccessHelpContextNativeProject(project) ||
                    !AccessHelpContextIdentity(project, (object)GetDesignProject(request.Project)) ||
                    Convert.ToInt32(((dynamic)project).Protection) != 0 ||
                    new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)) != window)
                    throw new InvalidOperationException("The approved Access project changed during final metadata authorization.");
                if (revision) AssertProjectVersion(request, project);
                if (!AccessHelpContextIdentity(project, (object)GetDesignProject(request.Project)))
                    throw new InvalidOperationException("The final metadata revision read resolved a different Access project.");
                request.RevalidateProjectPropertyAuthorization?.Invoke(false); // Cached authorization only, after all project COM reads.
            };
            AccessHelpContextDispatch dispatch = null;
            try
            {
                dispatch = AccessHelpContextFactory(project, window);
                dispatch.Put(converted, () => requireTarget(true));
            }
            catch (Exception error)
            {
                try { error.Data["AccessHelpContextInvokeEntries"] = dispatch?.InvokeEntries ?? 0; } catch (Exception) { }
                if (dispatch != null && dispatch.InvokeEntries != 0)
                    VbeScalarProperty.AnnotateFailure(error, VbeScalarProperty.FailurePhase.SetterInvocation);
                throw;
            }
            try
            {
                requireTarget(false);
                object actual = descriptor.GetValue(project);
                if (!(actual is int retained) || retained != converted)
                    throw new InvalidOperationException("The VBE did not retain property HelpContextID as the exact Int32 value.");
            }
            catch (Exception error) { VbeScalarProperty.AnnotateFailure(error, VbeScalarProperty.FailurePhase.RetentionReadback); throw; }
            return true;
        }

        private static bool SameAccessHelpContextIdentity(object first, object second)
        {
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct AccessHelpContextExceptionInfo
        {
            internal ushort Code, Reserved;
            internal IntPtr Source, Description, HelpFile;
            internal uint HelpContext;
            internal IntPtr ReservedPointer, DeferredCallback;
            internal int Scode;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct AccessHelpContextParameters
        {
            internal IntPtr Arguments, NamedArguments;
            internal uint ArgumentCount, NamedArgumentCount;
        }

        internal interface IAccessHelpContextCalls : IDisposable
        {
            int Prepare();
            void RequireOwner();
            int Invoke(int member, ushort flags, IntPtr result, ref AccessHelpContextParameters parameters, IntPtr exception, out uint argumentError);
            void FreeExceptionString(IntPtr text);
        }

        /// <summary>One owned I4 PROPERTYPUT, with no alternative setter or mutation replay.</summary>
        internal sealed class AccessHelpContextDispatch
        {
            internal const int VariantBytes = 24;
            internal const long Canary = 0x47E2195A0D62B38C;
            private readonly IAccessHelpContextCalls calls;
            private bool consumed;
            internal int InvokeEntries { get; private set; }
            internal AccessHelpContextDispatch(IAccessHelpContextCalls calls) { this.calls = calls ?? throw new ArgumentNullException(nameof(calls)); }

            internal void Put(int value, Action revalidate)
            {
                if (consumed) throw new InvalidOperationException("The original HelpContextID write was already consumed; no retry is allowed.");
                consumed = true;
                IntPtr argument = IntPtr.Zero, named = IntPtr.Zero, exceptionBuffer = IntPtr.Zero;
                var exception = new AccessHelpContextExceptionInfo();
                Exception failure = null;
                bool outputTrusted = true;
                try
                {
                    RequireAccessHelpContextAbi();
                    if (revalidate == null) throw new ArgumentNullException(nameof(revalidate));
                    calls.RequireOwner();
                    int member = calls.Prepare();
                    if (member != 117) throw new InvalidOperationException("The exact VBProject HelpContextID dispatch contract was not bound.");
                    argument = Marshal.AllocHGlobal(VariantBytes + sizeof(long));
                    Marshal.Copy(new byte[VariantBytes], 0, argument, VariantBytes);
                    Marshal.WriteInt16(argument, 3); // VT_I4, not BYREF or a marshaled object.
                    Marshal.WriteInt32(argument, 8, value);
                    Marshal.WriteInt64(argument, VariantBytes, Canary);
                    byte[] expected = new byte[VariantBytes]; Marshal.Copy(argument, expected, 0, expected.Length);
                    named = Marshal.AllocHGlobal(sizeof(int)); Marshal.WriteInt32(named, -3);
                    exceptionBuffer = Marshal.AllocHGlobal(64 + sizeof(long));
                    Marshal.Copy(new byte[64], 0, exceptionBuffer, 64); Marshal.WriteInt64(exceptionBuffer, 64, Canary);
                    var parameters = new AccessHelpContextParameters { Arguments = argument, NamedArguments = named, ArgumentCount = 1, NamedArgumentCount = 1 };
                    revalidate(); // Final live project/revision/runtime approval after readonly preparation.
                    calls.RequireOwner(); // Exact held IDispatch/project and native VBE UI-thread/PID immediately before Invoke.
                    InvokeEntries = 1;
                    outputTrusted = false; // A thrown/unknown Invoke cannot authorize output pointer cleanup.
                    int hr = calls.Invoke(member, 4, IntPtr.Zero, ref parameters, exceptionBuffer, out uint argumentError);
                    bool exceptionBoundary = Marshal.ReadInt64(exceptionBuffer, 64) == Canary;
                    if (exceptionBoundary) exception = (AccessHelpContextExceptionInfo)Marshal.PtrToStructure(exceptionBuffer, typeof(AccessHelpContextExceptionInfo));
                    byte[] actual = new byte[VariantBytes]; Marshal.Copy(argument, actual, 0, actual.Length);
                    bool unchanged = true; for (int index = 0; index < expected.Length; index++) unchanged &= expected[index] == actual[index];
                    outputTrusted = exceptionBoundary && Marshal.ReadInt64(argument, VariantBytes) == Canary && unchanged &&
                        Marshal.ReadInt32(named) == -3 && parameters.Arguments == argument && parameters.NamedArguments == named &&
                        parameters.ArgumentCount == 1 && parameters.NamedArgumentCount == 1;
                    if (hr != 0)
                        throw new COMException("Access HelpContextID PROPERTYPUT failed. Native HRESULT 0x" + unchecked((uint)hr).ToString("X8", CultureInfo.InvariantCulture) +
                            "; EXCEPINFO scode 0x" + unchecked((uint)exception.Scode).ToString("X8", CultureInfo.InvariantCulture) +
                            "; deferred callback " + (exception.DeferredCallback != IntPtr.Zero) +
                            ((hr == unchecked((int)0x80020005) || hr == unchecked((int)0x80020004)) ? "; argument " + argumentError : "") + ". Do not retry automatically.", hr);
                    if (!outputTrusted)
                    {
                        outputTrusted = false;
                        throw new InvalidOperationException("Native HelpContextID changed its input/ABI boundary; the mutation outcome is uncertain and must not be retried.");
                    }
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    // I4 has no nested allocation; never VariantClear a possibly corrupted native argument.
                    if (named != IntPtr.Zero) Marshal.FreeHGlobal(named);
                    if (argument != IntPtr.Zero) Marshal.FreeHGlobal(argument);
                    if (exceptionBuffer != IntPtr.Zero) Marshal.FreeHGlobal(exceptionBuffer);
                    if (outputTrusted)
                    {
                        var freed = new System.Collections.Generic.HashSet<IntPtr>();
                        foreach (IntPtr text in new[] { exception.Source, exception.Description, exception.HelpFile })
                            if (text != IntPtr.Zero && freed.Add(text))
                                try { calls.FreeExceptionString(text); }
                                catch (Exception cleanup) { PreserveAccessHelpContextFailure(ref failure, cleanup); }
                    }
                    else if (failure != null)
                    {
                        try { failure.Data["AccessHelpContextOutputCleanup"] = "Untrusted native output pointers were not dereferenced or freed; any native EXCEPINFO BSTR allocation may remain retained. Outcome uncertain; no replay."; }
                        catch (Exception) { }
                    }
                    try { calls.Dispose(); }
                    catch (Exception cleanup) { PreserveAccessHelpContextFailure(ref failure, cleanup); }
                }
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            }
        }

        private static void PreserveAccessHelpContextFailure(ref Exception original, Exception cleanup)
        {
            if (original == null) { original = cleanup; return; }
            // Preserve the first HRESULT/exception; cleanup must not replace it with an aggregate HRESULT.
            try { original.Data["AccessHelpContextCleanupFailure"] = cleanup.GetType().FullName + ": " + cleanup.Message; }
            catch (Exception) { }
        }

        internal static void RequireAccessHelpContextAbi()
        {
            if (IntPtr.Size != 8 || Marshal.SizeOf(typeof(AccessHelpContextParameters)) != 24 ||
                Marshal.SizeOf(typeof(AccessHelpContextExceptionInfo)) != 64 ||
                Marshal.OffsetOf(typeof(AccessHelpContextExceptionInfo), nameof(AccessHelpContextExceptionInfo.Source)).ToInt32() != 8 ||
                Marshal.OffsetOf(typeof(AccessHelpContextExceptionInfo), nameof(AccessHelpContextExceptionInfo.Scode)).ToInt32() != 56)
                throw new PlatformNotSupportedException("The validated Windows x64 HelpContextID native ABI is required.");
        }

        private sealed class NativeAccessHelpContextCalls : IAccessHelpContextCalls
        {
            [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int NamesDelegate(IntPtr self, ref Guid iid, IntPtr names, uint count, uint locale, IntPtr member);
            [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int PutDelegate(IntPtr self, int member, ref Guid iid, uint locale, ushort flags,
                ref AccessHelpContextParameters parameters, IntPtr result, IntPtr exception, out uint argumentError);
            [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int TypeInfoDelegate(IntPtr self, uint index, uint locale, out IntPtr info);
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
            [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
            private readonly object project;
            private readonly IntPtr window;
            private readonly uint pid, thread;
            private IntPtr originalIdentity, defaultDispatch, dispatch, dispatchIdentity;
            private PutDelegate put;
            private bool disposed;

            internal NativeAccessHelpContextCalls(object project, IntPtr window)
            {
                this.project = project; this.window = window;
                pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            }
            public void RequireOwner()
            {
                uint owner;
                if (disposed || window == IntPtr.Zero || GetCurrentProcessId() != pid || GetCurrentThreadId() != thread ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || GetWindowThreadProcessId(window, out owner) != thread || owner != pid)
                    throw new InvalidOperationException("HelpContextID must stay on its original current-process VBE UI STA.");
                if (originalIdentity != IntPtr.Zero)
                {
                    if (dispatchIdentity != originalIdentity) throw new InvalidOperationException("The original held project/IDispatch canonical identity changed.");
                }
            }
            public int Prepare()
            {
                RequireOwner();
                originalIdentity = Marshal.GetIUnknownForObject(project);
                defaultDispatch = Marshal.GetIDispatchForObject(project);
                Guid projectInterface = new Guid("EEE00915-E393-11D1-BB03-00C04FB6C4A6");
                int typedQi = Marshal.QueryInterface(defaultDispatch, ref projectInterface, out dispatch);
                if (typedQi != 0) throw new COMException("The native target does not expose the observed VBProject dispatch interface.", typedQi);
                Guid unknown = new Guid("00000000-0000-0000-C000-000000000046");
                int qi = Marshal.QueryInterface(dispatch, ref unknown, out dispatchIdentity);
                if (qi != 0) throw new COMException("The held project IDispatch has no verified canonical IUnknown.", qi);
                RequireOwner();
                VerifyTypeContract();
                put = (PutDelegate)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(dispatch), 6 * IntPtr.Size), typeof(PutDelegate));
                var namesCall = (NamesDelegate)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(dispatch), 5 * IntPtr.Size), typeof(NamesDelegate));
                IntPtr name = Marshal.StringToCoTaskMemUni("HelpContextID"), names = IntPtr.Zero, member = IntPtr.Zero;
                try
                {
                    names = Marshal.AllocHGlobal(IntPtr.Size); Marshal.WriteIntPtr(names, name);
                    member = Marshal.AllocHGlobal(sizeof(int)); Marshal.WriteInt32(member, -1);
                    Guid iid = Guid.Empty;
                    int hr = namesCall(dispatch, ref iid, names, 1, (uint)CultureInfo.InvariantCulture.LCID, member);
                    if (hr != 0) throw new COMException("HelpContextID name binding failed before any mutation.", hr);
                    return Marshal.ReadInt32(member);
                }
                finally { if (member != IntPtr.Zero) Marshal.FreeHGlobal(member); if (names != IntPtr.Zero) Marshal.FreeHGlobal(names); Marshal.FreeCoTaskMem(name); }
            }
            private void VerifyTypeContract()
            {
                var read = (TypeInfoDelegate)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(dispatch), 4 * IntPtr.Size), typeof(TypeInfoDelegate));
                IntPtr pointer; int hr = read(dispatch, 0, 0, out pointer);
                ITypeInfo info = null;
                Exception failure = null;
                try
                {
                    if (hr != 0) throw new COMException("VBProject type information was unavailable before mutation.", hr);
                    if (pointer == IntPtr.Zero) throw new InvalidOperationException("VBProject type information is missing.");
                    info = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(pointer, typeof(ITypeInfo));
                    IntPtr attribute; info.GetTypeAttr(out attribute);
                    try
                    {
                        var type = (TYPEATTR)Marshal.PtrToStructure(attribute, typeof(TYPEATTR));
                        bool getter = false, setter = false, invalid = false;
                        int getters = 0, setters = 0;
                        if (type.cFuncs < 1 || type.cFuncs > 256) throw new InvalidOperationException("VBProject type information exceeds its bound.");
                        for (int index = 0; index < type.cFuncs; index++)
                        {
                            IntPtr function; info.GetFuncDesc(index, out function);
                            try
                            {
                                var member = (FUNCDESC)Marshal.PtrToStructure(function, typeof(FUNCDESC));
                                if (member.memid != 117) continue;
                                if (member.invkind == INVOKEKIND.INVOKE_PROPERTYGET)
                                {
                                    if (++getters != 1) invalid = true;
                                    getter = member.cParams == 0 && member.elemdescFunc.tdesc.vt == (short)VarEnum.VT_I4;
                                    invalid |= !getter;
                                }
                                else if (member.invkind == INVOKEKIND.INVOKE_PROPERTYPUT && member.cParams == 1 && member.lprgelemdescParam != IntPtr.Zero)
                                {
                                    if (++setters != 1) invalid = true;
                                    var parameter = (ELEMDESC)Marshal.PtrToStructure(member.lprgelemdescParam, typeof(ELEMDESC));
                                    setter = member.elemdescFunc.tdesc.vt == (short)VarEnum.VT_VOID && parameter.tdesc.vt == (short)VarEnum.VT_I4 &&
                                        (parameter.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FIN) != 0 &&
                                        (parameter.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FOUT) == 0;
                                    invalid |= !setter;
                                }
                                else invalid = true;
                            }
                            finally { info.ReleaseFuncDesc(function); }
                        }
                        RequireAccessHelpContextTypeContract(type.guid, type.typekind, getter, setter, invalid);
                    }
                    finally { info.ReleaseTypeAttr(attribute); }
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    if (info != null) try { Marshal.ReleaseComObject(info); } catch (Exception error) { PreserveAccessHelpContextFailure(ref failure, error); }
                    if (pointer != IntPtr.Zero) try { Marshal.Release(pointer); } catch (Exception error) { PreserveAccessHelpContextFailure(ref failure, error); }
                }
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            }
            public int Invoke(int member, ushort flags, IntPtr result, ref AccessHelpContextParameters parameters, IntPtr exception, out uint argumentError)
            {
                Guid iid = Guid.Empty;
                return put(dispatch, member, ref iid, (uint)CultureInfo.InvariantCulture.LCID, flags, ref parameters, result, exception, out argumentError);
            }
            public void FreeExceptionString(IntPtr text) { Marshal.FreeBSTR(text); }
            public void Dispose()
            {
                if (disposed) return; disposed = true;
                Exception failure = null;
                foreach (IntPtr pointer in new[] { dispatchIdentity, dispatch, defaultDispatch, originalIdentity })
                    if (pointer != IntPtr.Zero) try { Marshal.Release(pointer); } catch (Exception error) { PreserveAccessHelpContextFailure(ref failure, error); }
                dispatchIdentity = dispatch = defaultDispatch = originalIdentity = IntPtr.Zero;
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            }
        }

        internal static void RequireAccessHelpContextTypeContract(Guid guid, TYPEKIND kind, bool getter, bool setter, bool invalid)
        {
            if (guid != new Guid("EEE00915-E393-11D1-BB03-00C04FB6C4A6") || kind != TYPEKIND.TKIND_DISPATCH || !getter || !setter || invalid)
                throw new InvalidOperationException("Only the observed VBProject Int32 GET/scalar PROPERTYPUT contract is supported.");
        }
    }
}
