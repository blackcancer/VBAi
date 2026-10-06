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

        /// <summary>Maintains the read process name state for vba test word values host.</summary>
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

        /// <summary>Identifies the read process id associated with vba test word values host.</summary>
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

        /// <summary>Maintains the read active application state for vba test word values host.</summary>
        internal Func<string, object> ReadActiveApplication;

        /// <summary>Maintains the native application reader state for vba test word values host.</summary>
        private readonly Func<string, object> nativeApplicationReader;

        /// <summary>Maintains the retained references state for vba test word values host.</summary>
        private static readonly ConcurrentBag<object> retainedReferences = new ConcurrentBag<object>();

        /// <summary>Maintains the read window owner state for vba test word values host.</summary>
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };

        /// <summary>Maintains the same identity state for vba test word values host.</summary>
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);

        /// <summary>Maintains the read document item state for vba test word values host.</summary>
        internal Func<object, int, object> ReadDocumentItem = (documents, index) => ((dynamic)documents)[index];

        /// <summary>Maintains the activate document state for vba test word values host.</summary>
        internal Action<object> ActivateDocument = document => ((dynamic)document).Activate();

        /// <summary>Maintains the run procedure state for vba test word values host.</summary>
        internal Func<object, string, object[], object> RunProcedure = NativeRun;

        /// <summary>Tracks the is com reference state of vba test word values host.</summary>
        internal static Func<object, bool> IsComReference = Marshal.IsComObject;

        /// <summary>Maintains the release com reference state for vba test word values host.</summary>
        internal static Func<object, int> ReleaseComReference = Marshal.ReleaseComObject;

        /// <summary>Maintains the owner thread state for vba test word values host.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Maintains the identifier state for vba test word values host.</summary>
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        /// <summary>Initializes a VbaTestWordValuesHost instance with the supplied state.</summary>
        internal VbaTestWordValuesHost() : this(Marshal.GetActiveObject) { }

        // The stable native resolver identifies acquisitions owned by this transport.
        // Replacing ReadActiveApplication injects a borrowed application instead.
        /// <summary>Initializes a VbaTestWordValuesHost instance with the supplied state.</summary>
        /// <param name="nativeApplicationReader">func&lt;string, object&gt; that supplies the native application reader for this operation.</param>
        internal VbaTestWordValuesHost(Func<string, object> nativeApplicationReader)
        {
            this.nativeApplicationReader = nativeApplicationReader;
            ReadActiveApplication = nativeApplicationReader;
        }

        /// <summary>Owns the application lease state and operations.</summary>
        internal sealed class ApplicationLease : IDisposable
        {

            /// <summary>Maintains the owner state for application lease.</summary>
            private readonly VbaTestWordValuesHost owner;

            /// <summary>Maintains the acquired state for application lease.</summary>
            private readonly bool acquired;

            /// <summary>Maintains the disposed and retained state for application lease.</summary>
            private bool disposed, retained;

            /// <summary>Gets or sets the application.</summary>
            /// <value>Current application exposed by application lease.</value>
            internal object Application { get; private set; }

            /// <summary>Initializes a ApplicationLease instance with the supplied state.</summary>
            /// <param name="owner">vba test word values host that supplies the owner for this operation.</param>
            /// <param name="application">object that supplies the application for this operation.</param>
            /// <param name="acquired">Indicates whether acquired is enabled.</param>
            internal ApplicationLease(VbaTestWordValuesHost owner, object application, bool acquired)
            { this.owner = owner; Application = application; this.acquired = acquired; }

            /// <summary>Handles retain on uncertain for application lease.</summary>
            internal void RetainOnUncertain()
            {
                owner.RequireOwner();
                if (disposed || retained) return;
                retained = true;
                RetainAcquired(this);
            }

            /// <summary>Disposes  for application lease.</summary>
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

        /// <summary>Owns the owned target state and operations.</summary>
        internal sealed class OwnedTarget : IDisposable
        {

            /// <summary>Maintains the owner state for owned target.</summary>
            internal VbaTestWordValuesHost Owner;

            /// <summary>Maintains the application and document and project state for owned target.</summary>
            internal object Application, Document, Project;

            /// <summary>Keeps the path path available to owned target.</summary>
            internal string Path;

            /// <summary>Maintains the application ownership state for owned target.</summary>
            internal ApplicationLease ApplicationOwnership;

            /// <summary>Maintains the disposed state for owned target.</summary>
            private bool disposed;

            /// <summary>Gets or sets the is retained.</summary>
            /// <value>Current is retained exposed by owned target.</value>
            internal bool IsRetained { get; private set; }

            /// <summary>Requires usable for owned target.</summary>
            internal void RequireUsable()
            {
                if (disposed || IsRetained)
                    throw new InvalidOperationException("The Word target has been released or retained after an uncertain operation.");
            }

            /// <summary>Handles retain on uncertain for owned target.</summary>
            internal void RetainOnUncertain()
            {
                Owner.RequireOwner();
                if (disposed || IsRetained) return;
                IsRetained = true;
                RetainAcquired(this);
            }

            /// <summary>Disposes  for owned target.</summary>
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

        /// <summary>Resolves target for vba test word values host.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="expectedHostPath">Path used for the expected host path being processed.</param>
        /// <returns>object produced by the operation for resolve target on vba test word values host.</returns>
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

        /// <summary>Invokes  for vba test word values host.</summary>
        /// <param name="target">object that supplies the target for this operation.</param>
        /// <param name="module">Text that supplies the module value. Use the format required by the calling operation.</param>
        /// <param name="procedure">Text that supplies the procedure value. Use the format required by the calling operation.</param>
        /// <param name="arguments">object[] that supplies the arguments for this operation.</param>
        /// <returns>object produced by the operation for invoke on vba test word values host.</returns>
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

        /// <summary>Validates target for vba test word values host.</summary>
        /// <param name="target">object that supplies the target for this operation.</param>
        /// <returns>owned target produced by the operation for validate target on vba test word values host.</returns>
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

        /// <summary>Resolves application lease for vba test word values host.</summary>
        /// <returns>application lease produced by the operation for resolve application lease on vba test word values host.</returns>
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
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <returns>int ptr produced by the operation for read application window on vba test word values host.</returns>
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

        /// <summary>Handles collection count for vba test word values host.</summary>
        /// <param name="collection">object that supplies the collection for this operation.</param>
        /// <param name="kind">Text that supplies the kind value. Use the format required by the calling operation.</param>
        /// <returns>int produced by the operation for collection count on vba test word values host.</returns>
        private static int CollectionCount(object collection, string kind)
        {
            int count = Convert.ToInt32(((dynamic)collection).Count);
            if (count < 0 || count > 1000) throw new InvalidOperationException("Unexpected Word " + kind + " count.");
            return count;
        }

        /// <summary>Handles retain acquired for vba test word values host.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        internal static void RetainAcquired(object value)
        {
            if (value != null) retainedReferences.Add(value);
        }

        /// <summary>Releases acquired for vba test word values host.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        internal static void ReleaseAcquired(object value)
        {
            // Balance this getter/indexer acquisition once, even when its RCW aliases a borrowed target.
            if (value != null && IsComReference(value)) ReleaseComReference(value);
        }

        /// <summary>Requires unique document name for vba test word values host.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
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

        /// <summary>Requires unique module owner for vba test word values host.</summary>
        /// <param name="owned">owned target that supplies the owned for this operation.</param>
        /// <param name="module">Text that supplies the module value. Use the format required by the calling operation.</param>
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

        /// <summary>Finds document for vba test word values host.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>object produced by the operation for find document on vba test word values host.</returns>
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

        /// <summary>Requires owner for vba test word values host.</summary>
        internal void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("Word COM calls must use their owning thread."); }

        /// <summary>Requires absolute path for vba test word values host.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        internal static void RequireAbsolutePath(string path)
        { if (!IsAbsolutePath(path)) throw new InvalidOperationException("A saved absolute Word document path is required."); }

        /// <summary>Compares path for vba test word values host.</summary>
        /// <param name="first">Text that supplies the first value. Use the format required by the calling operation.</param>
        /// <param name="second">Text that supplies the second value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for same path on vba test word values host.</returns>
        internal static bool SamePath(string first, string second)
        { return IsAbsolutePath(first) && IsAbsolutePath(second) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }

        /// <summary>Determines whether absolute path for vba test word values host.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>Boolean indicating the result of the check for is absolute path on vba test word values host.</returns>
        private static bool IsAbsolutePath(string path)
        { if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) return false; string root = Path.GetPathRoot(path); return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal); }

        /// <summary>Handles native run for vba test word values host.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <param name="macro">Text that supplies the macro value. Use the format required by the calling operation.</param>
        /// <param name="arguments">object[] that supplies the arguments for this operation.</param>
        /// <returns>object produced by the operation for native run on vba test word values host.</returns>
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
