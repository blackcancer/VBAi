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

        /// <summary>Maintains the read process name state for vba test power point values host.</summary>
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

        /// <summary>Identifies the read process id associated with vba test power point values host.</summary>
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

        /// <summary>Maintains the read active application state for vba test power point values host.</summary>
        internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;

        /// <summary>Maintains the read application window state for vba test power point values host.</summary>
        internal Func<object, IntPtr> ReadApplicationWindow = PowerPointWindow.Read;

        /// <summary>Maintains the read window owner state for vba test power point values host.</summary>
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };

        /// <summary>Maintains the same identity state for vba test power point values host.</summary>
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);

        /// <summary>Maintains the run procedure state for vba test power point values host.</summary>
        internal Func<object, string, object[], object> RunProcedure = NativeRun;

        /// <summary>Maintains the owner thread state for vba test power point values host.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Maintains the identifier state for vba test power point values host.</summary>
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        /// <summary>Owns the owned target state and operations.</summary>
        internal sealed class OwnedTarget
        {

            /// <summary>Maintains the owner state for owned target.</summary>
            internal VbaTestPowerPointValuesHost Owner;

            /// <summary>Maintains the application and presentation and project state for owned target.</summary>
            internal object Application, Presentation, Project;

            /// <summary>Keeps the path path available to owned target.</summary>
            internal string Path;
        }

        /// <summary>Resolves target for vba test power point values host.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="expectedHostPath">Path used for the expected host path being processed.</param>
        /// <returns>object produced by the operation for resolve target on vba test power point values host.</returns>
        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner();
            if (project == null) throw new InvalidOperationException("An exact PowerPoint project is required.");
            RequireAbsolutePath(expectedHostPath);
            object application = ResolveApplication();
            object match = FindPresentation(application, project, expectedHostPath);
            return new OwnedTarget { Owner = this, Application = application, Presentation = match, Project = project, Path = Path.GetFullPath(expectedHostPath) };
        }

        /// <summary>Invokes  for vba test power point values host.</summary>
        /// <param name="target">object that supplies the target for this operation.</param>
        /// <param name="module">Text that supplies the module value. Use the format required by the calling operation.</param>
        /// <param name="procedure">Text that supplies the procedure value. Use the format required by the calling operation.</param>
        /// <param name="arguments">object[] that supplies the arguments for this operation.</param>
        /// <returns>object produced by the operation for invoke on vba test power point values host.</returns>
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

        /// <summary>Validates target for vba test power point values host.</summary>
        /// <param name="target">object that supplies the target for this operation.</param>
        /// <returns>owned target produced by the operation for validate target on vba test power point values host.</returns>
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

        /// <summary>Resolves application for vba test power point values host.</summary>
        /// <returns>object produced by the operation for resolve application on vba test power point values host.</returns>
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
