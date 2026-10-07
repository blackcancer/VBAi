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

        /// <summary>Process-name reader used to require in-process WINWORD dispatch.</summary>
        internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

        /// <summary>Process-ID reader used to verify that the registered Word window belongs to this process.</summary>
        internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

        /// <summary>Resolver for the registered Word.Application object; replacing it can inject a borrowed instance.</summary>
        internal Func<string, object> ReadActiveApplication;

        /// <summary>Original stable resolver used to distinguish transport-owned COM references from injected ones.</summary>
        private readonly Func<string, object> nativeApplicationReader;

        /// <summary>Process-lifetime roots for acquired Word COM targets retained after uncertain native completion.</summary>
        private static readonly ConcurrentBag<object> retainedReferences = new ConcurrentBag<object>();

        /// <summary>Reads the owning process ID for Word's application-window HWND.</summary>
        internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { VbeDebugWindows.GetWindowThreadProcessId(hwnd, out uint owner); return owner; };

        /// <summary>Compares managed identity first, then native COM identity.</summary>
        internal Func<object, object, bool> SameIdentity = (first, second) => ReferenceEquals(first, second) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second);

        /// <summary>Reads a one-based Word Documents collection item.</summary>
        internal Func<object, int, object> ReadDocumentItem = (documents, index) => ((dynamic)documents)[index];

        /// <summary>Activates the exact target document before qualified Word macro dispatch.</summary>
        internal Action<object> ActivateDocument = document => ((dynamic)document).Activate();

        /// <summary>Single Word Application.Run route; a thrown call is not replayed.</summary>
        internal Func<object, string, object[], object> RunProcedure = NativeRun;

        /// <summary>Tracks the is com reference state of vba test word values host.</summary>
        internal static Func<object, bool> IsComReference = Marshal.IsComObject;

        /// <summary>COM release callback used only for references acquired by this transport.</summary>
        internal static Func<object, int> ReleaseComReference = Marshal.ReleaseComObject;

        /// <summary>Managed thread ID that owns every Word COM operation performed by this adapter.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Unicode VBA identifier rule used before building Word's qualified macro name.</summary>
        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        /// <summary>Creates the Word adapter with the production active-object resolver.</summary>
        internal VbaTestWordValuesHost() : this(Marshal.GetActiveObject) { }

        // The stable native resolver identifies acquisitions owned by this transport.
        // Replacing ReadActiveApplication injects a borrowed application instead.
        /// <summary>Creates the adapter with an injectable Word application resolver.</summary>
        /// <param name="nativeApplicationReader">Resolver used for stable production acquisitions and ownership-aware release.</param>
        internal VbaTestWordValuesHost(Func<string, object> nativeApplicationReader)
        {
            this.nativeApplicationReader = nativeApplicationReader;
            ReadActiveApplication = nativeApplicationReader;
        }

        /// <summary>Tracks whether a Word application COM reference is borrowed, acquired, disposed, or retained.</summary>
        internal sealed class ApplicationLease : IDisposable
        {

            /// <summary>Adapter that created and validates this lease.</summary>
            private readonly VbaTestWordValuesHost owner;

            /// <summary>True only when this lease owns a reference acquired through the native resolver.</summary>
            private readonly bool acquired;

            /// <summary>Prevents release after disposal or uncertain-call retention.</summary>
            private bool disposed, retained;

            /// <summary>Word.Application COM object held by this lease.</summary>
            /// <value>Null after successful disposal; retained leases keep the reference rooted.</value>
            internal object Application { get; private set; }

            /// <summary>Creates an application lease with explicit COM-reference ownership.</summary>
            /// <param name="owner">Adapter responsible for owner-thread validation and release.</param>
            /// <param name="application">Word.Application reference held for this lease.</param>
            /// <param name="acquired">Whether this adapter acquired and therefore must release the reference.</param>
            internal ApplicationLease(VbaTestWordValuesHost owner, object application, bool acquired)
            { this.owner = owner; Application = application; this.acquired = acquired; }

            /// <summary>Keeps an acquired application reference rooted after uncertain macro completion.</summary>
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

        /// <summary>Validated Word application, document, and VBProject identity retained for one run target.</summary>
        internal sealed class OwnedTarget : IDisposable
        {

            /// <summary>Adapter that created this target and validates future operations.</summary>
            internal VbaTestWordValuesHost Owner;

            /// <summary>Exact application, document, and VBProject COM identities resolved for the target.</summary>
            internal object Application, Document, Project;

            /// <summary>Normalized saved document path used to reacquire and validate the target.</summary>
            internal string Path;

            /// <summary>Lease governing release or retention of the application reference.</summary>
            internal ApplicationLease ApplicationOwnership;

            /// <summary>Prevents a released target from being reused or disposed twice.</summary>
            private bool disposed;

            /// <summary>Indicates whether this target was rooted after uncertain macro completion.</summary>
            /// <value>True means the target must not be disposed or reused.</value>
            internal bool IsRetained { get; private set; }

            /// <summary>Rejects a target already disposed or retained after uncertain completion.</summary>
            internal void RequireUsable()
            {
                if (disposed || IsRetained)
                    throw new InvalidOperationException("The Word target has been released or retained after an uncertain operation.");
            }

            /// <summary>Roots the complete target and its COM references after uncertain macro completion.</summary>
            internal void RetainOnUncertain()
            {
                Owner.RequireOwner();
                if (disposed || IsRetained) return;
                IsRetained = true;
                RetainAcquired(this);
            }

            /// <summary>Releases the document and its owned application lease unless retained after uncertainty.</summary>
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

        /// <summary>Resolves the exact saved Word document that owns the supplied VBProject.</summary>
        /// <param name="project">Live Word VBProject identity being tested.</param>
        /// <param name="expectedHostPath">Absolute saved document path that must match the owning document.</param>
        /// <returns>Owned target retaining the application lease, document, project, and normalized path.</returns>
        public object ResolveTarget(object project, string expectedHostPath)
        {
            RequireOwner(); RequireAbsolutePath(expectedHostPath);
            if (project == null) throw new InvalidOperationException("An exact Word project is required.");
            string normalizedPath = Path.GetFullPath(expectedHostPath);
            var application = ResolveApplicationLease();
            try
            {
                return new OwnedTarget
                {
                    Owner = this,
                    Application = application.Application,
                    Document = FindDocument(application.Application, project, normalizedPath),
                    Project = project,
                    Path = normalizedPath,
                    ApplicationOwnership = application
                };
            }
            catch { application.Dispose(); throw; }
        }

        /// <summary>Invokes one module-qualified macro only after validating and activating the exact owned document.</summary>
        /// <param name="target">Owned target returned by <see cref="ResolveTarget"/>.</param>
        /// <param name="module">Resolved VBA module identifier.</param>
        /// <param name="procedure">Resolved procedure identifier.</param>
        /// <param name="arguments">Zero arguments or the supported two positional arguments; the array is cloned.</param>
        /// <returns>Value returned by Word Application.Run.</returns>
        /// <exception cref="VbaTestInvocationException">Activation or dispatch throws; the outcome is marked uncertain and is not retried.</exception>
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

        /// <summary>Reacquires the exact Word application and document and checks their COM identities.</summary>
        /// <param name="target">Target returned by ResolveTarget for this adapter.</param>
        /// <returns>Validated owned target whose application and document still match the captured identities.</returns>
        internal OwnedTarget ValidateTarget(object target)
        {
            RequireOwner();
            if (!(target is OwnedTarget owned) || !ReferenceEquals(owned.Owner, this)) throw new InvalidOperationException("An owned Word target is required.");
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
        /// <param name="application">Word.Application instance whose active document window is inspected.</param>
        /// <returns>Active document-window HWND owned by the registered Word process.</returns>
        internal static IntPtr ReadApplicationWindow(object application)
        {
            if (application == null) throw new InvalidOperationException("The registered Word application is unavailable.");
            object window = null;
            try
            {
                window = ((dynamic)application).ActiveWindow ?? throw new InvalidOperationException("The registered Word application has no active document window.");
                var handle = new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd));
                if (handle == IntPtr.Zero) throw new InvalidOperationException("The active Word document window has no native handle.");
                return handle;
            }
            finally { ReleaseAcquired(window); }
        }

        /// <summary>Reads a bounded COM collection count before indexed enumeration.</summary>
        /// <param name="collection">Word Documents or VBE VBProjects collection being enumerated.</param>
        /// <param name="kind">Collection label included in the refusal message when its count is outside 0 through 1,000.</param>
        /// <returns>Validated collection item count.</returns>
        private static int CollectionCount(object collection, string kind)
        {
            int count = Convert.ToInt32(((dynamic)collection).Count);
            if (count < 0 || count > 1000) throw new InvalidOperationException("Unexpected Word " + kind + " count.");
            return count;
        }

        /// <summary>Roots an acquired COM reference so it remains alive after uncertain completion.</summary>
        /// <param name="value">Acquired COM object to keep alive.</param>
        internal static void RetainAcquired(object value)
        {
            if (value != null) retainedReferences.Add(value);
        }

        /// <summary>Releases acquired for vba test word values host.</summary>
        /// <param name="value">Acquired COM object to keep alive.</param>
        internal static void ReleaseAcquired(object value)
        {
            // Balance this getter/indexer acquisition once, even when its RCW aliases a borrowed target.
            if (value != null && IsComReference(value)) ReleaseComReference(value);
        }

        /// <summary>Requires exactly one open Word document to match the macro-qualified filename.</summary>
        /// <param name="application">Word.Application instance whose active document window is inspected.</param>
        /// <param name="name">Saved document filename used in Word's qualified macro name.</param>
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

        /// <summary>Requires exactly one loaded VBE module with this name, owned by the target project.</summary>
        /// <param name="owned">Validated Word target whose project must be the sole owner of the module.</param>
        /// <param name="module">Module identifier that Word will resolve in the qualified macro name.</param>
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

        /// <summary>Finds the single saved open document whose VBProject identity and full path both match.</summary>
        /// <param name="application">Word.Application instance whose active document window is inspected.</param>
        /// <param name="project">Exact VBProject identity used to select its owning document.</param>
        /// <param name="path">Expected normalized absolute path of the saved document.</param>
        /// <returns>Owned Document COM reference; its caller is responsible for releasing it once.</returns>
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

        /// <summary>Requires all Word automation calls to run on the adapter's constructing thread.</summary>
        internal void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("Word COM calls must use their owning thread."); }

        /// <summary>Rejects unsaved or nonabsolute host-document paths before COM resolution.</summary>
        /// <param name="path">Expected normalized absolute path of the saved document.</param>
        internal static void RequireAbsolutePath(string path)
        { if (!IsAbsolutePath(path)) throw new InvalidOperationException("A saved absolute Word document path is required."); }

        /// <summary>Compares path for vba test word values host.</summary>
        /// <param name="first">First candidate path.</param>
        /// <param name="second">Expected path.</param>
        /// <returns>True when both paths are absolute and normalize to the same path, case-insensitively.</returns>
        internal static bool SamePath(string first, string second)
        { return IsAbsolutePath(first) && IsAbsolutePath(second) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }

        /// <summary>Checks whether a path is rooted at a drive or network share.</summary>
        /// <param name="path">Expected normalized absolute path of the saved document.</param>
        /// <returns>True for a rooted drive or UNC path; false for empty, relative, or drive-relative paths.</returns>
        private static bool IsAbsolutePath(string path)
        { if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) return false; string root = Path.GetPathRoot(path); return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal); }

        /// <summary>Calls Word Application.Run using its required 30 optional by-reference argument slots.</summary>
        /// <param name="application">Word.Application instance whose active document window is inspected.</param>
        /// <param name="macro">Module-qualified Word macro name.</param>
        /// <param name="arguments">Zero or two populated positional arguments copied into the first slots; the remainder are Type.Missing.</param>
        /// <returns>Value returned by Word Application.Run.</returns>
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
