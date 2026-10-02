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
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };
        internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;
        internal Func<object, IntPtr> ReadApplicationWindow = PowerPointWindow.Read;
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);
        internal Func<object, string, object[], object> RunProcedure = NativeRun;
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        internal sealed class OwnedTarget
        {
            internal VbaTestPowerPointValuesHost Owner;
            internal object Application, Presentation, Project;
            internal string Path;
        }

        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner();
            if (project == null) throw new InvalidOperationException("An exact PowerPoint project is required.");
            RequireAbsolutePath(expectedHostPath);
            object application = ResolveApplication();
            object match = FindPresentation(application, project, expectedHostPath);
            return new OwnedTarget { Owner = this, Application = application, Presentation = match, Project = project, Path = Path.GetFullPath(expectedHostPath) };
        }

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

        internal void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("PowerPoint COM calls must use their owning thread."); }
        internal static void RequireAbsolutePath(string path)
        { if (!IsAbsolutePath(path)) throw new InvalidOperationException("A saved absolute presentation path is required."); }
        internal static bool SamePath(string first, string second)
        { return IsAbsolutePath(first) && IsAbsolutePath(second)
            && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }
        private static bool IsAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) return false;
            string root = Path.GetPathRoot(path);
            return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal);
        }
        private static object NativeRun(object application, string macro, object[] arguments)
        {
            // The PowerPoint PIA declares Run(string, ref object[]), not Excel's positional argument list.
            var powerPoint = (Microsoft.Office.Interop.PowerPoint._Application)application;
            return powerPoint.Run(macro, ref arguments);
        }
    }
}
