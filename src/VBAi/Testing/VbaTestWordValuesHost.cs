using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace VBAi
{
    /// <summary>Word Application.Run transport restricted to a saved document in the owning process.</summary>
    internal sealed class VbaTestWordValuesHost : VbeDebug.IProcedureValuesHost
    {
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };
        internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);
        internal Action<object> ActivateDocument = document => ((dynamic)document).Activate();
        internal Func<object, string, object[], object> RunProcedure = NativeRun;
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        internal sealed class OwnedTarget
        {
            internal VbaTestWordValuesHost Owner;
            internal object Application, Document, Project;
            internal string Path;
        }

        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner(); RequireAbsolutePath(expectedHostPath);
            if (project == null) throw new InvalidOperationException("An exact Word project is required.");
            object application = ResolveApplication();
            return new OwnedTarget { Owner = this, Application = application, Document = FindDocument(application, project, expectedHostPath),
                Project = project, Path = Path.GetFullPath(expectedHostPath) };
        }

        public object Invoke(object target, string module, string procedure, object[] arguments)
        {
            var owned = ValidateTarget(target);
            arguments = arguments == null ? new object[0] : (object[])arguments.Clone();
            if (!Identifier.IsMatch(module ?? "") || !Identifier.IsMatch(procedure ?? "") || (arguments.Length != 0 && arguments.Length != 2))
                throw new InvalidOperationException("Word testing requires resolved identifiers and zero or two positional arguments.");
            string name = (string)((dynamic)owned.Document).Name;
            if (!string.Equals(name, Path.GetFileName(owned.Path), StringComparison.OrdinalIgnoreCase)
                || name.IndexOfAny(new[] { '\'', '!', '\r', '\n' }) >= 0)
                throw new InvalidOperationException("The document filename cannot qualify this macro safely.");
            int names = 0;
            foreach (dynamic document in ((dynamic)owned.Application).Documents)
                if (string.Equals((string)document.Name, name, StringComparison.OrdinalIgnoreCase)) names++;
            if (names != 1) throw new InvalidOperationException("The qualified Word document filename is ambiguous.");
            // Word restricts document-qualified macros to the current context. Never run an
            // unqualified macro or substitute a different active document if activation fails.
            try
            {
                ActivateDocument(owned.Document);
                ValidateTarget(owned);
                if (!SameIdentity((object)((dynamic)owned.Application).ActiveDocument, owned.Document))
                    throw new InvalidOperationException("The owned Word document did not become the exact macro context.");
                names = 0;
                foreach (dynamic document in ((dynamic)owned.Application).Documents)
                    if (string.Equals((string)document.Name, name, StringComparison.OrdinalIgnoreCase)) names++;
                if (names != 1) throw new InvalidOperationException("The qualified Word document filename became ambiguous during activation.");
                return RunProcedure(owned.Application, "'" + name + "'!" + module + "." + procedure, arguments);
            }
            catch (Exception error)
            { throw new VbaTestInvocationException("Word activation or macro completion is uncertain; no retry was attempted. " + error.Message, true, error); }
        }

        internal OwnedTarget ValidateTarget(object target)
        {
            RequireOwner();
            var owned = target as OwnedTarget;
            if (owned == null || !ReferenceEquals(owned.Owner, this)) throw new InvalidOperationException("An owned Word target is required.");
            object application = ResolveApplication();
            if (!SameIdentity(application, owned.Application)) throw new InvalidOperationException("The owned Word application identity changed.");
            if (!SameIdentity(FindDocument(application, owned.Project, owned.Path), owned.Document))
                throw new InvalidOperationException("The owned Word document identity changed.");
            return owned;
        }

        internal object ResolveApplication()
        {
            RequireOwner();
            if (!string.Equals(ReadProcessName(), "WINWORD", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Word returned values require the in-process WINWORD host.");
            object application = ReadActiveApplication("Word.Application");
            if (ReadWindowOwner(ReadApplicationWindow(application)) != (uint)ReadProcessId())
                throw new InvalidOperationException("The registered Word application belongs to another PID.");
            return application;
        }

        /// <summary>Reads Word's active document Window.Hwnd without activating or creating a window.</summary>
        internal static IntPtr ReadApplicationWindow(object application)
        {
            if (application == null) throw new InvalidOperationException("The registered Word application is unavailable.");
            object window = ((dynamic)application).ActiveWindow;
            if (window == null) throw new InvalidOperationException("The registered Word application has no active document window.");
            var handle = new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd));
            if (handle == IntPtr.Zero) throw new InvalidOperationException("The active Word document window has no native handle.");
            return handle;
        }
        private object FindDocument(object application, object project, string path)
        {
            object match = null; int count = 0;
            foreach (dynamic document in ((dynamic)application).Documents)
            {
                if (++count > 1000) throw new InvalidOperationException("Unexpected open Word document count.");
                if (!SameIdentity(project, (object)document.VBProject)) continue;
                if (match != null) throw new InvalidOperationException("Multiple documents share the selected project identity.");
                if (string.IsNullOrWhiteSpace((string)document.Path) || !SamePath((string)document.FullName, path))
                    throw new InvalidOperationException("The owned Word document path changed or has never been saved.");
                match = document;
            }
            if (match == null) throw new InvalidOperationException("No owned open document shares this VBProject identity.");
            return match;
        }

        internal void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("Word COM calls must use their owning thread."); }
        internal static void RequireAbsolutePath(string path)
        { if (!IsAbsolutePath(path)) throw new InvalidOperationException("A saved absolute Word document path is required."); }
        internal static bool SamePath(string first, string second)
        { return IsAbsolutePath(first) && IsAbsolutePath(second) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }
        private static bool IsAbsolutePath(string path)
        { if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) return false; string root = Path.GetPathRoot(path); return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal); }
        private static object NativeRun(object application, string macro, object[] arguments)
        {
            var positional = new object[arguments.Length + 1]; positional[0] = macro;
            Array.Copy(arguments, 0, positional, 1, arguments.Length);
            return application.GetType().InvokeMember("Run", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding,
                null, application, positional, System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
