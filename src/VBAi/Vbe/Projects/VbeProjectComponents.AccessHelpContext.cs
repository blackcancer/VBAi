using Microsoft.Vbe.Interop;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi
{

    /// <summary>Provides the bounded Access/Publisher VBProject HelpContextID setter path.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Host-process predicate for the Access and Publisher HelpContextID compatibility path.</summary>
        internal Func<bool> AccessHelpContextHost = () => IsAccessHelpContextHost(Process.GetCurrentProcess().ProcessName);

        /// <summary>Predicate requiring the project argument to be an existing COM object.</summary>
        internal Func<object, bool> AccessHelpContextNativeProject = Marshal.IsComObject;

        /// <summary>Canonical COM identity comparison used before and after authorization-sensitive reads.</summary>
        internal Func<object, object, bool> AccessHelpContextIdentity = SameAccessHelpContextIdentity;

        /// <summary>Factory for the early-bound VBProject setter adapter, injectable for focused tests.</summary>
        internal Func<object, IntPtr, AccessHelpContextDispatch> AccessHelpContextFactory = (project, window) =>
            new AccessHelpContextDispatch(new NativeAccessHelpContextCalls(project, window));

        /// <summary>Gets the access help context interface type.</summary>
        /// <value>Current access help context interface type exposed by vbe project components.</value>
        internal static Type AccessHelpContextInterfaceType => typeof(_VBProject);

        /// <summary>Checks whether the current host process is Access or Publisher.</summary>
        /// <param name="processName">Process name, compared case-insensitively without a file extension.</param>
        /// <returns>True only for <c>MSACCESS</c> or <c>MSPUB</c>.</returns>
        internal static bool IsAccessHelpContextHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

        /// <summary>Sets and reads back an Access/Publisher project's HelpContextID through one early-bound PIA call.</summary>
        /// <param name="request">Authorized scalar property-write request with exact project revision and Int32 value.</param>
        /// <param name="project">Held canonical VBProject COM object that must retain identity and remain unprotected.</param>
        /// <returns>True when this compatibility setter handled the property; false when the property, host, or object is outside its scope.</returns>
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
            void requireTarget(bool revision)
            {
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
            }
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

        /// <summary>Compares the canonical IUnknown identity of two project RCWs.</summary>
        /// <param name="first">First COM project object.</param>
        /// <param name="second">Second COM project object.</param>
        /// <returns>True when both objects resolve to the same IUnknown pointer.</returns>
        private static bool SameAccessHelpContextIdentity(object first, object second)
        {
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }

        /// <summary>Owner-thread calls required to prepare and invoke the one Access HelpContextID setter.</summary>
        internal interface IAccessHelpContextCalls : IDisposable
        {

            /// <summary>Performs interface preparation and captures canonical COM identity before final authorization.</summary>
            void Prepare();

            /// <summary>Requires the original current-process x64 VBE UI STA and, after preparation, unchanged COM identity.</summary>
            void RequireOwner();

            /// <summary>Invokes the declared VBProject HelpContextID Int32 property setter.</summary>
            /// <param name="value">Scalar HelpContextID value to write.</param>
            void Set(int value);
        }

        /// <summary>One declared PIA setter with final authorization; an uncertain outcome has no alternative dispatch or retry.</summary>
        internal sealed class AccessHelpContextDispatch
        {

            /// <summary>Native owner-bound calls used for preparation, setter invocation, and cleanup.</summary>
            private readonly IAccessHelpContextCalls calls;

            /// <summary>One-use guard set before validation or native invocation so the write cannot be retried.</summary>
            private bool consumed;

            /// <summary>Gets or sets the invoke entries.</summary>
            /// <value>Current invoke entries exposed by access help context dispatch.</value>
            internal int InvokeEntries { get; private set; }

            /// <summary>Creates a dispatcher around the native call boundary.</summary>
            /// <param name="calls">Non-null owner-thread call implementation.</param>
            internal AccessHelpContextDispatch(IAccessHelpContextCalls calls) { this.calls = calls ?? throw new ArgumentNullException(nameof(calls)); }

            /// <summary>Claims the setter once, revalidates authority immediately before it, then always releases identity pointers.</summary>
            /// <param name="value">Int32 HelpContextID value to send through the declared setter.</param>
            /// <param name="revalidate">Final live target/revision check performed after COM interface preparation.</param>
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

        /// <summary>Attaches cleanup failure details to an earlier primary failure without replacing it.</summary>
        /// <param name="original">Exception describing the original failure.</param>
        /// <param name="cleanup">Exception describing the cleanup failure.</param>
        private static void PreserveAccessHelpContextFailure(ref Exception original, Exception cleanup)
        {
            if (original == null) { original = cleanup; return; }
            try { original.Data["AccessHelpContextCleanupFailure"] = cleanup.GetType().FullName + ": " + cleanup.Message; }
            catch (Exception) { }
        }

        /// <summary>Holds a borrowed VBProject RCW and invokes its official PIA setter on the captured VBE STA.</summary>
        private sealed class NativeAccessHelpContextCalls : IAccessHelpContextCalls
        {

            /// <summary>Reads the current Win32 thread ID for native owner validation.</summary>
            /// <returns>Current native thread ID.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            /// <summary>Reads the current Win32 process ID for native owner validation.</summary>
            /// <returns>Current process ID.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

            /// <summary>Borrowed canonical VBProject RCW; never activated or explicitly released.</summary>
            private readonly object project;

            /// <summary>VBE main-window HWND captured when the adapter is created.</summary>
            private readonly IntPtr window;

            /// <summary>Current process and native UI-thread IDs captured at construction.</summary>
            private readonly uint pid, thread;

            /// <summary>Temporary IUnknown pointers used to verify the queried PIA interface belongs to the original RCW.</summary>
            private IntPtr originalIdentity, typedIdentity;

            /// <summary>Borrowed early-bound PIA interface used for the declared Int32 property setter.</summary>
            private _VBProject typedProject; // Borrowed shared RCW; never activate or ReleaseComObject/FinalReleaseComObject it.

            /// <summary>Tracks whether interface identity was captured and whether temporary pointers were released.</summary>
            private bool prepared, disposed;

            /// <summary>Captures the current process and native thread for a held project and VBE window.</summary>
            /// <param name="project">Borrowed project RCW whose HelpContextID property may be set once.</param>
            /// <param name="window">Nonzero VBE main-window HWND that defines native ownership.</param>
            internal NativeAccessHelpContextCalls(object project, IntPtr window)
            {
                this.project = project; this.window = window;
                pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            }

            /// <summary>Requires the captured x64 process and VBE UI STA; after preparation the PIA interface must retain COM identity.</summary>
            public void RequireOwner()
            {
                if (disposed || IntPtr.Size != 8 || window == IntPtr.Zero || GetCurrentProcessId() != pid || GetCurrentThreadId() != thread ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || GetWindowThreadProcessId(window, out uint owner) != thread || owner != pid)
                    throw new InvalidOperationException("HelpContextID must stay on its original current-process x64 VBE UI STA.");
                if (prepared && (typedProject == null || originalIdentity == IntPtr.Zero || typedIdentity != originalIdentity))
                    throw new InvalidOperationException("The original held VBProject and its typed canonical identity differ.");
            }

            /// <summary>Captures original and PIA IUnknown identities without taking ownership of the shared RCW.</summary>
            public void Prepare()
            {
                RequireOwner();
                originalIdentity = Marshal.GetIUnknownForObject(project);
                typedProject = (_VBProject)project; // Official modern dual PIA interface; CLR supplies its supported QI/member binding.
                typedIdentity = Marshal.GetIUnknownForObject(typedProject);
                prepared = true;
                RequireOwner();
            }

            /// <summary>Invokes the declared early-bound Int32 HelpContextID setter; failed HRESULTs become exceptions.</summary>
            /// <param name="value">HelpContextID value passed to the VBProject property setter.</param>
            public void Set(int value)
            {
                // PIA Void setter has no PreserveSig: CLR translates failed native HRESULTs into exceptions.
                typedProject.HelpContextID = value;
            }

            /// <summary>Releases temporary IUnknown references and drops the borrowed managed RCW reference.</summary>
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
