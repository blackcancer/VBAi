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

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Maintains the access help context host state for vbe project components.</summary>
        internal Func<bool> AccessHelpContextHost = () => IsAccessHelpContextHost(Process.GetCurrentProcess().ProcessName);

        /// <summary>Maintains the access help context native project state for vbe project components.</summary>
        internal Func<object, bool> AccessHelpContextNativeProject = Marshal.IsComObject;

        /// <summary>Maintains the access help context identity state for vbe project components.</summary>
        internal Func<object, object, bool> AccessHelpContextIdentity = SameAccessHelpContextIdentity;

        /// <summary>Maintains the access help context factory state for vbe project components.</summary>
        internal Func<object, IntPtr, AccessHelpContextDispatch> AccessHelpContextFactory = (project, window) =>
            new AccessHelpContextDispatch(new NativeAccessHelpContextCalls(project, window));

        /// <summary>Gets the access help context interface type.</summary>
        /// <value>Current access help context interface type exposed by vbe project components.</value>
        internal static Type AccessHelpContextInterfaceType => typeof(_VBProject);

        /// <summary>Determines whether access help context host for vbe project components.</summary>
        /// <param name="processName">Text that supplies the process name value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is access help context host on vbe project components.</returns>
        internal static bool IsAccessHelpContextHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

        /// <summary>Attempts to set access help context for vbe project components.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <returns>Boolean indicating the result of the check for try set access help context on vbe project components.</returns>
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

        /// <summary>Compares access help context identity for vbe project components.</summary>
        /// <param name="first">object that supplies the first for this operation.</param>
        /// <param name="second">object that supplies the second for this operation.</param>
        /// <returns>Boolean indicating the result of the check for same access help context identity on vbe project components.</returns>
        private static bool SameAccessHelpContextIdentity(object first, object second)
        {
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }

        /// <summary>Defines the i access help context calls contract.</summary>
        internal interface IAccessHelpContextCalls : IDisposable
        {

            /// <summary>Handles prepare for i access help context calls.</summary>
            void Prepare();

            /// <summary>Requires owner for i access help context calls.</summary>
            void RequireOwner();

            /// <summary>Sets  for i access help context calls.</summary>
            /// <param name="value">int that supplies the value for this operation.</param>
            void Set(int value);
        }

        /// <summary>One declared PIA setter with final authorization; an uncertain outcome has no alternative dispatch or retry.</summary>
        internal sealed class AccessHelpContextDispatch
        {

            /// <summary>Maintains the calls state for access help context dispatch.</summary>
            private readonly IAccessHelpContextCalls calls;

            /// <summary>Maintains the consumed state for access help context dispatch.</summary>
            private bool consumed;

            /// <summary>Gets or sets the invoke entries.</summary>
            /// <value>Current invoke entries exposed by access help context dispatch.</value>
            internal int InvokeEntries { get; private set; }

            /// <summary>Initializes a AccessHelpContextDispatch instance with the supplied state.</summary>
            /// <param name="calls">i access help context calls that supplies the calls for this operation.</param>
            internal AccessHelpContextDispatch(IAccessHelpContextCalls calls) { this.calls = calls ?? throw new ArgumentNullException(nameof(calls)); }

            /// <summary>Handles put for access help context dispatch.</summary>
            /// <param name="value">int that supplies the value for this operation.</param>
            /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
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

        /// <summary>Handles preserve access help context failure for vbe project components.</summary>
        /// <param name="original">Exception describing the original failure.</param>
        /// <param name="cleanup">Exception describing the cleanup failure.</param>
        private static void PreserveAccessHelpContextFailure(ref Exception original, Exception cleanup)
        {
            if (original == null) { original = cleanup; return; }
            try { original.Data["AccessHelpContextCleanupFailure"] = cleanup.GetType().FullName + ": " + cleanup.Message; }
            catch (Exception) { }
        }

        /// <summary>Owns the native access help context calls state and operations.</summary>
        private sealed class NativeAccessHelpContextCalls : IAccessHelpContextCalls
        {

            /// <summary>Returns current thread id for native access help context calls.</summary>
            /// <returns>uint produced by the operation for get current thread id on native access help context calls.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            /// <summary>Returns current process id for native access help context calls.</summary>
            /// <returns>uint produced by the operation for get current process id on native access help context calls.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

            /// <summary>Maintains the project state for native access help context calls.</summary>
            private readonly object project;

            /// <summary>Maintains the window state for native access help context calls.</summary>
            private readonly IntPtr window;

            /// <summary>Identifies the pid and thread associated with native access help context calls.</summary>
            private readonly uint pid, thread;

            /// <summary>Maintains the original identity and typed identity state for native access help context calls.</summary>
            private IntPtr originalIdentity, typedIdentity;

            /// <summary>Maintains the typed project state for native access help context calls.</summary>
            private _VBProject typedProject; // Borrowed shared RCW; never activate or ReleaseComObject/FinalReleaseComObject it.

            /// <summary>Maintains the prepared and disposed state for native access help context calls.</summary>
            private bool prepared, disposed;

            /// <summary>Initializes a NativeAccessHelpContextCalls instance with the supplied state.</summary>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            internal NativeAccessHelpContextCalls(object project, IntPtr window)
            {
                this.project = project; this.window = window;
                pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            }

            /// <summary>Requires owner for native access help context calls.</summary>
            public void RequireOwner()
            {
                uint owner;
                if (disposed || IntPtr.Size != 8 || window == IntPtr.Zero || GetCurrentProcessId() != pid || GetCurrentThreadId() != thread ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || GetWindowThreadProcessId(window, out owner) != thread || owner != pid)
                    throw new InvalidOperationException("HelpContextID must stay on its original current-process x64 VBE UI STA.");
                if (prepared && (typedProject == null || originalIdentity == IntPtr.Zero || typedIdentity != originalIdentity))
                    throw new InvalidOperationException("The original held VBProject and its typed canonical identity differ.");
            }

            /// <summary>Handles prepare for native access help context calls.</summary>
            public void Prepare()
            {
                RequireOwner();
                originalIdentity = Marshal.GetIUnknownForObject(project);
                typedProject = (_VBProject)project; // Official modern dual PIA interface; CLR supplies its supported QI/member binding.
                typedIdentity = Marshal.GetIUnknownForObject(typedProject);
                prepared = true;
                RequireOwner();
            }

            /// <summary>Sets  for native access help context calls.</summary>
            /// <param name="value">int that supplies the value for this operation.</param>
            public void Set(int value)
            {
                // PIA Void setter has no PreserveSig: CLR translates failed native HRESULTs into exceptions.
                typedProject.HelpContextID = value;
            }

            /// <summary>Disposes  for native access help context calls.</summary>
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
