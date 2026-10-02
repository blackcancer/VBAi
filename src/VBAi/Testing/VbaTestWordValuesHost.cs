using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
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
        internal Func<string, object> ReadActiveApplication;
        private readonly Func<string, object> nativeApplicationReader;
        private static readonly ConcurrentBag<object> retainedReferences = new ConcurrentBag<object>();
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);
        internal Func<object, int, object> ReadDocumentItem = (documents, index) => ((dynamic)documents)[index];
        internal Action<object> ActivateDocument = document => ((dynamic)document).Activate();
        internal Func<object, string, object[], object> RunProcedure = NativeRun;
        internal static Func<object, bool> IsComReference = Marshal.IsComObject;
        internal static Func<object, int> ReleaseComReference = Marshal.ReleaseComObject;
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        internal VbaTestWordValuesHost() : this(Marshal.GetActiveObject) { }

        // The stable native resolver identifies acquisitions owned by this transport.
        // Replacing ReadActiveApplication injects a borrowed application instead.
        internal VbaTestWordValuesHost(Func<string, object> nativeApplicationReader)
        {
            this.nativeApplicationReader = nativeApplicationReader;
            ReadActiveApplication = nativeApplicationReader;
        }

        internal sealed class ApplicationLease : IDisposable
        {
            private readonly VbaTestWordValuesHost owner;
            private readonly bool acquired;
            private bool disposed, retained;
            internal object Application { get; private set; }
            internal ApplicationLease(VbaTestWordValuesHost owner, object application, bool acquired)
            { this.owner = owner; Application = application; this.acquired = acquired; }
            internal void RetainOnUncertain()
            {
                owner.RequireOwner();
                if (disposed || retained) return;
                retained = true;
                RetainAcquired(this);
            }
            public void Dispose()
            {
                owner.RequireOwner();
                if (disposed || retained) return;
                disposed = true;
                object application = Application;
                Application = null;
                if (acquired) ReleaseAcquired(application);
            }
        }

        internal sealed class OwnedTarget : IDisposable
        {
            internal VbaTestWordValuesHost Owner;
            internal object Application, Document, Project;
            internal string Path;
            internal ApplicationLease ApplicationOwnership;
            private bool disposed;
            internal bool IsRetained { get; private set; }
            internal void RequireUsable()
            {
                if (disposed || IsRetained)
                    throw new InvalidOperationException("The Word target has been released or retained after an uncertain operation.");
            }
            internal void RetainOnUncertain()
            {
                Owner.RequireOwner();
                if (disposed || IsRetained) return;
                IsRetained = true;
                RetainAcquired(this);
            }
            public void Dispose()
            {
                Owner.RequireOwner();
                if (disposed || IsRetained) return;
                disposed = true;
                object document = Document;
                var application = ApplicationOwnership;
                Application = Document = Project = null;
                ApplicationOwnership = null;
                try { ReleaseAcquired(document); }
                finally { application?.Dispose(); }
            }
        }

        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner(); RequireAbsolutePath(expectedHostPath);
            if (project == null) throw new InvalidOperationException("An exact Word project is required.");
            string normalizedPath = Path.GetFullPath(expectedHostPath);
            var application = ResolveApplicationLease();
            try
            {
                return new OwnedTarget { Owner = this, Application = application.Application,
                    Document = FindDocument(application.Application, project, normalizedPath),
                    Project = project, Path = normalizedPath, ApplicationOwnership = application };
            }
            catch { application.Dispose(); throw; }
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
            RequireUniqueDocumentName(owned.Application, name);
            RequireUniqueModuleOwner(owned, module);
            if (arguments.Length == 0 && string.Equals(module, VbaCoverageInstrumentation.ModuleName, StringComparison.OrdinalIgnoreCase)
                && (string.Equals(procedure, VbaCoverageInstrumentation.ResetProcedure, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(procedure, VbaCoverageInstrumentation.SnapshotProcedure, StringComparison.OrdinalIgnoreCase)))
                arguments = new object[] { false, false };
            // Word resolves module-qualified names. Prove the only matching module belongs to
            // this exact project; neither another document nor a global template may supply it.
            try
            {
                ActivateDocument(owned.Document);
                ValidateTarget(owned);
                object active = null;
                try
                {
                    active = ((dynamic)owned.Application).ActiveDocument;
                    if (!SameIdentity(active, owned.Document))
                        throw new InvalidOperationException("The owned Word document did not become the exact macro context.");
                }
                finally { ReleaseAcquired(active); }
                RequireUniqueDocumentName(owned.Application, name);
                RequireUniqueModuleOwner(owned, module);
                return RunProcedure(owned.Application, module + "." + procedure, arguments);
            }
            catch (Exception error)
            { owned.RetainOnUncertain(); throw new VbaTestInvocationException("Word activation or macro completion is uncertain; no retry was attempted. " + error.Message, true, error); }
        }

        internal OwnedTarget ValidateTarget(object target)
        {
            RequireOwner();
            var owned = target as OwnedTarget;
            if (owned == null || !ReferenceEquals(owned.Owner, this)) throw new InvalidOperationException("An owned Word target is required.");
            owned.RequireUsable();
            using (var application = ResolveApplicationLease())
            {
                if (!SameIdentity(application.Application, owned.Application)) throw new InvalidOperationException("The owned Word application identity changed.");
                object document = FindDocument(application.Application, owned.Project, owned.Path);
                try
                {
                    if (!SameIdentity(document, owned.Document))
                        throw new InvalidOperationException("The owned Word document identity changed.");
                }
                finally { ReleaseAcquired(document); }
            }
            return owned;
        }

        internal ApplicationLease ResolveApplicationLease()
        {
            RequireOwner();
            if (!string.Equals(ReadProcessName(), "WINWORD", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Word returned values require the in-process WINWORD host.");
            bool acquired = ReferenceEquals(ReadActiveApplication, nativeApplicationReader);
            var lease = new ApplicationLease(this, ReadActiveApplication("Word.Application"), acquired);
            try
            {
                if (ReadWindowOwner(ReadApplicationWindow(lease.Application)) != (uint)ReadProcessId())
                    throw new InvalidOperationException("The registered Word application belongs to another PID.");
                // Respect the host policy before dispatch or opening an instrumented coverage copy.
                if (Convert.ToInt32(((dynamic)lease.Application).AutomationSecurity) == 3)
                    throw new InvalidOperationException("Word AutomationSecurity is ForceDisable; returned VBA values and coverage copies cannot execute under the current host policy.");
                return lease;
            }
            catch { lease.Dispose(); throw; }
        }

        /// <summary>Reads Word's active document Window.Hwnd without activating or creating a window.</summary>
        internal static IntPtr ReadApplicationWindow(object application)
        {
            if (application == null) throw new InvalidOperationException("The registered Word application is unavailable.");
            object window = null;
            try
            {
                window = ((dynamic)application).ActiveWindow;
                if (window == null) throw new InvalidOperationException("The registered Word application has no active document window.");
                var handle = new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd));
                if (handle == IntPtr.Zero) throw new InvalidOperationException("The active Word document window has no native handle.");
                return handle;
            }
            finally { ReleaseAcquired(window); }
        }

        private static int CollectionCount(object collection, string kind)
        {
            int count = Convert.ToInt32(((dynamic)collection).Count);
            if (count < 0 || count > 1000) throw new InvalidOperationException("Unexpected Word " + kind + " count.");
            return count;
        }

        internal static void RetainAcquired(object value)
        {
            if (value != null) retainedReferences.Add(value);
        }

        internal static void ReleaseAcquired(object value)
        {
            // Balance this getter/indexer acquisition once, even when its RCW aliases a borrowed target.
            if (value != null && IsComReference(value)) ReleaseComReference(value);
        }

        private void RequireUniqueDocumentName(object application, string name)
        {
            object documents = null;
            try
            {
                documents = ((dynamic)application).Documents;
                int count = CollectionCount(documents, "document"), matches = 0;
                for (int index = 1; index <= count; index++)
                {
                    object document = null;
                    try
                    {
                        document = ReadDocumentItem(documents, index);
                        if (string.Equals((string)((dynamic)document).Name, name, StringComparison.OrdinalIgnoreCase)) matches++;
                    }
                    finally { ReleaseAcquired(document); }
                }
                if (matches != 1) throw new InvalidOperationException("The qualified Word document filename is ambiguous.");
            }
            finally { ReleaseAcquired(documents); }
        }

        private void RequireUniqueModuleOwner(OwnedTarget owned, string module)
        {
            object vbe = null, projects = null;
            try
            {
                vbe = ((dynamic)owned.Application).VBE;
                projects = ((dynamic)vbe).VBProjects;
                int count = CollectionCount(projects, "project"), matches = 0;
                for (int index = 1; index <= count; index++)
                {
                    object project = null, components = null;
                    try
                    {
                        project = ((dynamic)projects)[index];
                        components = ((dynamic)project).VBComponents;
                        int componentCount = CollectionCount(components, "component");
                        for (int componentIndex = 1; componentIndex <= componentCount; componentIndex++)
                        {
                            object component = null;
                            try
                            {
                                component = ((dynamic)components)[componentIndex];
                                if (!string.Equals((string)((dynamic)component).Name, module, StringComparison.OrdinalIgnoreCase)) continue;
                                matches++;
                                if (!SameIdentity(project, owned.Project))
                                    throw new InvalidOperationException("The Word module qualifier also belongs to another loaded project or template.");
                            }
                            finally { ReleaseAcquired(component); }
                        }
                    }
                    finally { ReleaseAcquired(components); ReleaseAcquired(project); }
                }
                if (matches != 1) throw new InvalidOperationException("No unique module in the exact owned Word project qualifies this procedure.");
            }
            finally { ReleaseAcquired(projects); ReleaseAcquired(vbe); }
        }

        private object FindDocument(object application, object project, string path)
        {
            object documents = null, match = null;
            try
            {
                documents = ((dynamic)application).Documents;
                int count = CollectionCount(documents, "document");
                for (int index = 1; index <= count; index++)
                {
                    object document = null, candidateProject = null;
                    bool transferred = false;
                    try
                    {
                        document = ReadDocumentItem(documents, index);
                        candidateProject = ((dynamic)document).VBProject;
                        if (!SameIdentity(project, candidateProject)) continue;
                        if (match != null) throw new InvalidOperationException("Multiple documents share the selected project identity.");
                        if (string.IsNullOrWhiteSpace((string)((dynamic)document).Path) || !SamePath((string)((dynamic)document).FullName, path))
                            throw new InvalidOperationException("The owned Word document path changed or has never been saved.");
                        match = document; transferred = true;
                    }
                    finally { ReleaseAcquired(candidateProject); if (!transferred) ReleaseAcquired(document); }
                }
                if (match == null) throw new InvalidOperationException("No owned open document shares this VBProject identity.");
                return match;
            }
            catch { ReleaseAcquired(match); throw; }
            finally { ReleaseAcquired(documents); }
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
            // Word requires 30 optional by-reference VARIANT slots, not a shortened dispatch argument list.
            var word = (Microsoft.Office.Interop.Word._Application)application;
            var slots = new object[30];
            for (int i = 0; i < slots.Length; i++) slots[i] = Type.Missing;
            Array.Copy(arguments, slots, arguments.Length);
            return word.Run(macro,
                ref slots[0], ref slots[1], ref slots[2], ref slots[3], ref slots[4],
                ref slots[5], ref slots[6], ref slots[7], ref slots[8], ref slots[9],
                ref slots[10], ref slots[11], ref slots[12], ref slots[13], ref slots[14],
                ref slots[15], ref slots[16], ref slots[17], ref slots[18], ref slots[19],
                ref slots[20], ref slots[21], ref slots[22], ref slots[23], ref slots[24],
                ref slots[25], ref slots[26], ref slots[27], ref slots[28], ref slots[29]);
        }
    }
}
