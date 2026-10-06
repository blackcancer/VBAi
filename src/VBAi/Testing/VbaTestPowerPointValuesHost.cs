using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace VBAi
{

    /// <summary>PowerPoint's single SAFEARRAY Application.Run transport, restricted to the owning process/thread.</summary>
    internal sealed class VbaTestPowerPointValuesHost : VbeDebug.IProcedureValuesHost
    {

        /// <summary>Process-name reader used to require in-process POWERPNT execution.</summary>
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

        /// <summary>Process-ID reader used to verify ownership of the registered application window.</summary>
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

        /// <summary>Resolver for the registered PowerPoint.Application COM object.</summary>
        internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;

        /// <summary>Reads PowerPoint's application-window HWND without starting another host instance.</summary>
        internal Func<object, IntPtr> ReadApplicationWindow = PowerPointWindow.Read;

        /// <summary>Reads the owning PID for an HWND.</summary>
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };

        /// <summary>Compares managed identity first, then native COM identity.</summary>
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);

        /// <summary>Single Application.Run route using PowerPoint's SAFEARRAY argument convention.</summary>
        internal Func<object, string, object[], object> RunProcedure = NativeRun;

        /// <summary>Managed thread ID required for all PowerPoint COM operations.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Culture-invariant VBA identifier grammar used before constructing a qualified macro name.</summary>
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        /// <summary>Resolved PowerPoint application, presentation, VBProject, and saved path for one run target.</summary>
        internal sealed class OwnedTarget
        {

            /// <summary>Adapter that created and revalidates this target.</summary>
            internal VbaTestPowerPointValuesHost Owner;

            /// <summary>Exact application, presentation, and VBProject COM identities.</summary>
            internal object Application, Presentation, Project;

            /// <summary>Normalized saved presentation path used for identity revalidation.</summary>
            internal string Path;
        }

        /// <summary>Resolves the exact saved presentation that owns the supplied VBProject.</summary>
        /// <param name="project">Live PowerPoint VBProject identity being tested.</param>
        /// <param name="expectedHostPath">Absolute saved presentation path that must match the owning presentation.</param>
        /// <returns>Owned target containing the application, presentation, project, and normalized path.</returns>
        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner();
            if (project == null) throw new InvalidOperationException("An exact PowerPoint project is required.");
            RequireAbsolutePath(expectedHostPath);
            object application = ResolveApplication();
            object match = FindPresentation(application, project, expectedHostPath);
            return new OwnedTarget { Owner = this, Application = application, Presentation = match, Project = project, Path = Path.GetFullPath(expectedHostPath) };
        }

        /// <summary>Invokes one presentation-qualified macro after checking target identity and filename uniqueness.</summary>
        /// <param name="target">Owned target returned by <see cref="ResolveTarget"/>.</param>
        /// <param name="module">Resolved VBA module identifier.</param>
        /// <param name="procedure">Resolved procedure identifier.</param>
        /// <param name="arguments">Optional positional arguments; the array is cloned before dispatch.</param>
        /// <returns>Value returned by PowerPoint Application.Run.</returns>
        /// <remarks>Uses one invocation shape; it does not retry or fall back to the active presentation.</remarks>
        public object Invoke(object target, string module, string procedure, object[] arguments)
        {
            var owned = ValidateTarget(target);
            if (!Identifier.IsMatch(module ?? "") || !Identifier.IsMatch(procedure ?? ""))
                throw new InvalidOperationException("A resolved VBA module and procedure identifier are required.");
            string name = (string)((dynamic)owned.Presentation).Name;
            if (!string.Equals(name, Path.GetFileName(owned.Path), StringComparison.OrdinalIgnoreCase)
                || name.IndexOfAny(new[] { '!', '\r', '\n' }) >= 0)
                throw new InvalidOperationException("The presentation filename cannot qualify this macro safely.");
            int sameNames = 0;
            foreach (dynamic presentation in ((dynamic)owned.Application).Presentations)
                if (string.Equals((string)presentation.Name, name, StringComparison.OrdinalIgnoreCase)) sameNames++;
            if (sameNames != 1) throw new InvalidOperationException("The qualified presentation filename is ambiguous.");
            // No retry, alternate invocation shape, active-presentation fallback or VBA expression evaluation.
            return RunProcedure(owned.Application, name + "!" + module + "." + procedure, arguments == null ? new object[0] : (object[])arguments.Clone());
        }

        /// <summary>Requires the target belong to this adapter and reacquires the same application and presentation.</summary>
        /// <param name="target">Object returned by <see cref="ResolveTarget"/>.</param>
        /// <returns>Validated owned target.</returns>
        internal OwnedTarget ValidateTarget(object target)
        {
            RequireOwner();
            var owned = target as OwnedTarget;
            if (owned == null || !ReferenceEquals(owned.Owner, this)) throw new InvalidOperationException("An owned PowerPoint target is required.");
            object application = ResolveApplication();
            if (!SameIdentity(application, owned.Application)) throw new InvalidOperationException("The owned PowerPoint application identity changed.");
            object presentation = FindPresentation(application, owned.Project, owned.Path);
            if (!SameIdentity(presentation, owned.Presentation)) throw new InvalidOperationException("The owned presentation identity changed.");
            return owned;
        }

        /// <summary>Resolves the registered PowerPoint instance in this process without changing host security policy.</summary>
        /// <returns>Application object whose window belongs to this PID.</returns>
        /// <exception cref="InvalidOperationException">The process, window ownership, registration, or AutomationSecurity state is unsafe.</exception>
        internal object ResolveApplication()
        {
            RequireOwner();
            if (!string.Equals(ReadProcessName(), "POWERPNT", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("PowerPoint returned values require the in-process POWERPNT host.");
            int processId = ReadProcessId();
            object application = ReadActiveApplication("PowerPoint.Application");
            if (application == null) throw new InvalidOperationException("The registered PowerPoint application is unavailable.");
            IntPtr window = ReadApplicationWindow(application);
            if (window == IntPtr.Zero || ReadWindowOwner(window) != (uint)processId)
                throw new InvalidOperationException("The registered PowerPoint application belongs to another PID.");
            // ForceDisable blocks every programmatically opened copy, even when its original can execute.
            // Respect the host's current policy: refuse before dispatch or copy mutation; never lower it here.
            if (Convert.ToInt32(((dynamic)application).AutomationSecurity) == 3)
                throw new InvalidOperationException("PowerPoint AutomationSecurity is ForceDisable; returned VBA values and coverage copies cannot execute under the current host policy.");
            return application;
        }

        /// <summary>Finds presentation for vba test power point values host.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="expectedPath">Path used for the expected path being processed.</param>
        /// <returns>object produced by the operation for find presentation on vba test power point values host.</returns>
        private object FindPresentation(object application, object project, string expectedPath)
        {
            object match = null; int count = 0;
            foreach (dynamic presentation in ((dynamic)application).Presentations)
            {
                if (++count > 1000) throw new InvalidOperationException("Unexpected open PowerPoint presentation count.");
                if (!SameIdentity(project, (object)presentation.VBProject)) continue;
                if (match != null) throw new InvalidOperationException("Multiple presentations share the selected project identity.");
                if (string.IsNullOrWhiteSpace((string)presentation.Path) || !SamePath((string)presentation.FullName, expectedPath))
                    throw new InvalidOperationException("The owned presentation path changed or has never been saved.");
                match = presentation;
            }
            if (match == null) throw new InvalidOperationException("No owned open presentation shares this VBProject identity.");
            return match;
        }

        /// <summary>Requires owner for vba test power point values host.</summary>
        internal void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("PowerPoint COM calls must use their owning thread."); }

        /// <summary>Requires absolute path for vba test power point values host.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        internal static void RequireAbsolutePath(string path)
        { if (!IsAbsolutePath(path)) throw new InvalidOperationException("A saved absolute presentation path is required."); }

        /// <summary>Compares path for vba test power point values host.</summary>
        /// <param name="first">Text that supplies the first value. Use the format required by the calling operation.</param>
        /// <param name="second">Text that supplies the second value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for same path on vba test power point values host.</returns>
        internal static bool SamePath(string first, string second)
        { return IsAbsolutePath(first) && IsAbsolutePath(second)
            && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }

        /// <summary>Determines whether absolute path for vba test power point values host.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>Boolean indicating the result of the check for is absolute path on vba test power point values host.</returns>
        private static bool IsAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) return false;
            string root = Path.GetPathRoot(path);
            return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal);
        }

        /// <summary>Handles native run for vba test power point values host.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <param name="macro">Text that supplies the macro value. Use the format required by the calling operation.</param>
        /// <param name="arguments">object[] that supplies the arguments for this operation.</param>
        /// <returns>object produced by the operation for native run on vba test power point values host.</returns>
        private static object NativeRun(object application, string macro, object[] arguments)
        {
            // The PowerPoint PIA declares Run(string, ref object[]), not Excel's positional argument list.
            var powerPoint = (Microsoft.Office.Interop.PowerPoint._Application)application;
            return powerPoint.Run(macro, ref arguments);
        }
    }
}
