using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Vbe.Interop;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        internal Func<bool> AccessHelpContextHost = () => IsAccessHelpContextHost(Process.GetCurrentProcess().ProcessName);
        internal Func<object, bool> AccessHelpContextNativeProject = Marshal.IsComObject;
        internal Func<object, object, bool> AccessHelpContextIdentity = SameAccessHelpContextIdentity;
        internal Func<object, IntPtr, AccessHelpContextDispatch> AccessHelpContextFactory = (project, window) =>
            new AccessHelpContextDispatch(new NativeAccessHelpContextCalls(project, window));

        internal static Type AccessHelpContextInterfaceType => typeof(_VBProject);

        internal static bool IsAccessHelpContextHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

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

        internal interface IAccessHelpContextCalls : IDisposable
        {
            void Prepare();
            void RequireOwner();
            void Set(int value);
        }

        /// <summary>One declared PIA setter with final authorization; an uncertain outcome has no alternative dispatch or retry.</summary>
        internal sealed class AccessHelpContextDispatch
        {
            private readonly IAccessHelpContextCalls calls;
            private bool consumed;
            internal int InvokeEntries { get; private set; }
            internal AccessHelpContextDispatch(IAccessHelpContextCalls calls) { this.calls = calls ?? throw new ArgumentNullException(nameof(calls)); }
            internal void Put(int value, Action revalidate)
            {
                if (consumed) throw new InvalidOperationException("The original HelpContextID write was already consumed; no retry is allowed.");
                consumed = true;
                Exception failure = null;
                try
                {
                    if (revalidate == null) throw new ArgumentNullException(nameof(revalidate));
                    calls.RequireOwner();
                    calls.Prepare(); // Cast/QI of the borrowed project completes before final live authorization.
                    revalidate();
                    calls.RequireOwner(); // Native owner and cached canonical identity only; no late host/property reads.
                    InvokeEntries = 1;
                    calls.Set(value); // One declared early-bound Int32 setter; failed native HRESULT propagates.
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    try { calls.Dispose(); }
                    catch (Exception cleanup) { PreserveAccessHelpContextFailure(ref failure, cleanup); }
                }
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            }
        }

        private static void PreserveAccessHelpContextFailure(ref Exception original, Exception cleanup)
        {
            if (original == null) { original = cleanup; return; }
            try { original.Data["AccessHelpContextCleanupFailure"] = cleanup.GetType().FullName + ": " + cleanup.Message; }
            catch (Exception) { }
        }

        private sealed class NativeAccessHelpContextCalls : IAccessHelpContextCalls
        {
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
            [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
            private readonly object project;
            private readonly IntPtr window;
            private readonly uint pid, thread;
            private IntPtr originalIdentity, typedIdentity;
            private _VBProject typedProject; // Borrowed shared RCW; never activate or ReleaseComObject/FinalReleaseComObject it.
            private bool prepared, disposed;

            internal NativeAccessHelpContextCalls(object project, IntPtr window)
            {
                this.project = project; this.window = window;
                pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            }
            public void RequireOwner()
            {
                uint owner;
                if (disposed || IntPtr.Size != 8 || window == IntPtr.Zero || GetCurrentProcessId() != pid || GetCurrentThreadId() != thread ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || GetWindowThreadProcessId(window, out owner) != thread || owner != pid)
                    throw new InvalidOperationException("HelpContextID must stay on its original current-process x64 VBE UI STA.");
                if (prepared && (typedProject == null || originalIdentity == IntPtr.Zero || typedIdentity != originalIdentity))
                    throw new InvalidOperationException("The original held VBProject and its typed canonical identity differ.");
            }
            public void Prepare()
            {
                RequireOwner();
                originalIdentity = Marshal.GetIUnknownForObject(project);
                typedProject = (_VBProject)project; // Official modern dual PIA interface; CLR supplies its supported QI/member binding.
                typedIdentity = Marshal.GetIUnknownForObject(typedProject);
                prepared = true;
                RequireOwner();
            }
            public void Set(int value)
            {
                // PIA Void setter has no PreserveSig: CLR translates failed native HRESULTs into exceptions.
                typedProject.HelpContextID = value;
            }
            public void Dispose()
            {
                if (disposed) return; disposed = true;
                Exception failure = null;
                foreach (IntPtr pointer in new[] { typedIdentity, originalIdentity })
                    if (pointer != IntPtr.Zero) try { Marshal.Release(pointer); } catch (Exception error) { PreserveAccessHelpContextFailure(ref failure, error); }
                typedIdentity = originalIdentity = IntPtr.Zero;
                typedProject = null; // Drop the borrowed managed reference, not the shared RCW.
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            }
        }
    }
}
