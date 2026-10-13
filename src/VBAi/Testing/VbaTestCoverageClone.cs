using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Owned coverage workbook; only the copy's source may be instrumented.</summary>
    internal sealed class VbaTestCoverageClone : IDisposable
    {

        /// <summary>Gets or sets the VBProject belonging to the owned disposable coverage copy.</summary>
        /// <value>Cloned project instrumented instead of the user's original.</value>
        internal object Project { get; set; }

        /// <summary>Gets or sets the retained filesystem path of the disposable copy.</summary>
        /// <value>Path reported if preparation or close has an uncertain outcome.</value>
        internal string Path { get; set; }

        /// <summary>Gets or sets the one-shot close action for the owned host document copy.</summary>
        /// <value>Cleanup action, cleared before invocation so disposal cannot retry it.</value>
        internal Action Close { get; set; }

        /// <summary>Invokes the owned copy's close action at most once.</summary>
        public void Dispose() { var close = Close; Close = null; close?.Invoke(); }

        /// <summary>Supplies exact-process/COM identity checks and host-specific disposable-copy factories.</summary>
        internal sealed class HostBoundary
        {

            /// <summary>Reads the current executable name to select the supported host adapter.</summary>
            internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

            /// <summary>Reads the current host process identifier for in-process Excel ownership checks.</summary>
            internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

            /// <summary>Resolves the Excel.Application object belonging to the current process ID.</summary>
            internal Func<int, object> ResolveExcel = pid => ExcelOwnedApplication.Resolve(pid, () => Marshal.GetActiveObject("Excel.Application"));

            /// <summary>Reads the owning process ID of a native application window.</summary>
            internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { VbeDebugWindows.GetWindowThreadProcessId(hwnd, out uint owner); return owner; };

            /// <summary>Compares COM object identity rather than display names or paths.</summary>
            internal Func<object, object, bool> SameIdentity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;

            /// <summary>Releases exactly one acquired COM reference after its owned operation is settled.</summary>
            internal Action<object> ReleaseAcquired = VbaTestComReferences.Release;

            /// <summary>Selects one-based native workbook indexing without creating an unreleasable enumeration adapter.</summary>
            internal Func<object, bool> IndexedCollection = Marshal.IsComObject;

            /// <summary>Creates and owns a separate Word coverage document copy.</summary>
            internal Func<object, string, string, VbaTestCoverageClone> CreateWord = VbaTestWordCoverageClone.CreateWord;

            /// <summary>Creates and owns a separate PowerPoint coverage presentation copy.</summary>
            internal Func<object, string, string, VbaTestCoverageClone> CreatePowerPoint = VbaTestPowerPointCoverageClone.CreatePowerPoint;
        }

        /// <summary>Creates a host-specific disposable project copy for coverage instrumentation.</summary>
        /// <param name="project">Live source VBProject whose identity is verified against its owning document.</param>
        /// <param name="sourcePath">Expected saved host-document path.</param>
        /// <param name="folder">Fresh directory that will contain the owned copy.</param>
        /// <returns>Owned copy and one-shot close action.</returns>
        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder)
        { return CreateOwned(project, sourcePath, folder, new HostBoundary()); }

        /// <summary>Selects the Word, PowerPoint, or in-process Excel cloning path by host executable.</summary>
        /// <param name="project">Source VBProject to match to the host document.</param>
        /// <param name="sourcePath">Expected source document path.</param>
        /// <param name="folder">Directory for the disposable copy.</param>
        /// <param name="boundary">Process, COM identity, and host-copy operations; defaults to production host services.</param>
        /// <returns>Host-specific owned copy.</returns>
        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder, HostBoundary boundary)
        {
            boundary = boundary ?? new HostBoundary();
            string name = boundary.ReadProcessName();
            if (name.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return boundary.CreateWord(project, sourcePath, folder);
            if (name.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)) return boundary.CreatePowerPoint(project, sourcePath, folder);
            return CreateExcel(project, sourcePath, folder, boundary);
        }

        /// <summary>Creates a SaveCopyAs workbook in the current Excel process with events suppressed during preparation.</summary>
        /// <param name="project">Source VBProject whose owning workbook must be uniquely identified.</param>
        /// <param name="sourcePath">Expected saved workbook path.</param>
        /// <param name="folder">Directory for the fresh coverage workbook.</param>
        /// <param name="boundary">Optional injected host operations for isolated contract checks.</param>
        /// <returns>Distinct project object and close action for the cloned workbook.</returns>
        /// <exception cref="InvalidOperationException">The host, process ownership, source identity, path, or supported workbook extension does not match.</exception>
        /// <exception cref="VbaTestInvocationException">Copy preparation or close/event restoration is uncertain; retained paths are included.</exception>
        internal static VbaTestCoverageClone CreateExcel(object project, string sourcePath, string folder, HostBoundary boundary = null)
        {
            boundary = boundary ?? new HostBoundary();
            var acquired = new ExcelCloneReferences(boundary.ReleaseAcquired);
            bool transferred = false;
            try
            {
                int processId = boundary.ReadProcessId();
                if (!boundary.ReadProcessName().Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Coverage cloning requires the in-process Excel host.");
                dynamic application = acquired.Add(boundary.ResolveExcel(processId));
                uint owner = boundary.ReadWindowOwner(new IntPtr(Convert.ToInt64(application.Hwnd)));
                if (owner != processId) throw new InvalidOperationException("The Excel application belongs to another PID.");
                object source = null;
                dynamic books = acquired.Add((object)application.Workbooks);
                using (var enumeration = new VbaTestComReferences(boundary.ReleaseAcquired, boundary.IndexedCollection))
                foreach (dynamic book in enumeration.ItemsFromBorrowed((object)books))
                {
                    bool matches;
                    using (var candidate = new VbaTestComReferences(boundary.ReleaseAcquired))
                        matches = boundary.SameIdentity(project, candidate.Add((object)book.VBProject));
                    if (!matches) continue;
                    if (source != null) throw new InvalidOperationException("The source workbook identity is ambiguous.");
                    if (!string.Equals(System.IO.Path.GetFullPath((string)book.FullName), System.IO.Path.GetFullPath(sourcePath), StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The source workbook path changed.");
                    enumeration.Transfer((object)book);
                    source = acquired.Add((object)book);
                }
                if (source == null) throw new InvalidOperationException("The exact source workbook is unavailable.");
                string extension = System.IO.Path.GetExtension(sourcePath);
                if (!new[] { ".xlsm", ".xlsb", ".xls" }.Contains(extension, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Coverage requires a saved XLSM, XLSB or XLS workbook.");
                Directory.CreateDirectory(folder);
                string copyPath = System.IO.Path.Combine(folder, "coverage" + extension);
                if (File.Exists(copyPath)) throw new InvalidOperationException("The coverage copy path is already occupied.");
                object copy = null;
                bool events = (bool)application.EnableEvents;
                try
                {
                    // Suppress BeforeSave and the copied Workbook_Open during preparation.
                    // Macro trust/security policies are retained and may refuse the copy.
                    application.EnableEvents = false;
                    if ((bool)application.EnableEvents) throw new InvalidOperationException("Excel events could not be disabled for coverage preparation.");
                    ((dynamic)source).SaveCopyAs(copyPath);
                    copy = acquired.Add((object)books.Open(copyPath, UpdateLinks: 0, ReadOnly: false, AddToMru: false));
                    object copyProject = acquired.Add((object)((dynamic)copy).VBProject);
                    if (boundary.SameIdentity(project, copyProject))
                        throw new InvalidOperationException("Excel did not create a distinct coverage project.");
                    transferred = true;
                    return new VbaTestCoverageClone
                    {
                        Project = copyProject,
                        Path = copyPath,
                        Close = () =>
                        {
                            bool previous = (bool)application.EnableEvents;
                            try
                            {
                                if (boundary.ReadWindowOwner(new IntPtr(Convert.ToInt64(application.Hwnd))) != processId
                                    || boundary.SameIdentity(project, copyProject)
                                    || !string.Equals(System.IO.Path.GetFullPath((string)((dynamic)copy).FullName), copyPath, StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("The owned Excel copy identity changed before Close.");
                                application.EnableEvents = false;
                                if ((bool)application.EnableEvents) throw new InvalidOperationException("Excel events could not be disabled before closing the coverage copy.");
                                // Verify with the live project first, then release our acquisition
                                // before Excel decides whether its VBA project can be unloaded.
                                acquired.ReleaseTracked(copyProject);
                                ((dynamic)copy).Close(false);
                                using (var remainingBooks = new VbaTestComReferences(boundary.ReleaseAcquired, boundary.IndexedCollection))
                                {
                                    try
                                    {
                                        foreach (object remaining in remainingBooks.ItemsFromBorrowed((object)books))
                                            if (boundary.SameIdentity(remaining, copy))
                                                throw new InvalidOperationException("The owned copy remains open after Close returned.");
                                    }
                                    catch { remainingBooks.Retain(); throw; }
                                }
                            }
                            catch (Exception error) { acquired.Retain(); throw new VbaTestInvocationException("Closing the owned workbook copy is uncertain. Retained path: " + copyPath + ". " + error.Message, true, error); }
                            finally
                            {
                                try { RestoreEvents((object)application, previous, copyPath); }
                                catch { acquired.Retain(); throw; }
                            }
                            acquired.Dispose();
                        }
                    };
                }
                catch (Exception error)
                {
                    throw new VbaTestInvocationException("Coverage copy preparation is uncertain; no close or retry was attempted. Retained copy: " + copyPath + ". " + error.Message, true, error);
                }
                finally { RestoreEvents((object)application, events, copyPath); }
            }
            catch (VbaTestInvocationException) { acquired.Retain(); throw; }
            finally { if (!transferred) acquired.Dispose(); }
        }

        /// <summary>Owns the Excel interfaces acquired by cloning until verified close or explicit uncertainty retention.</summary>
        private sealed class ExcelCloneReferences : IDisposable
        {
            private readonly Action<object> release;
            private readonly List<object> references = new List<object>();
            private static readonly List<ExcelCloneReferences> retainedReferences = new List<ExcelCloneReferences>();
            private bool retained, disposed;
            internal ExcelCloneReferences(Action<object> release) { this.release = release; }
            internal object Add(object value) { if (value != null) references.Add(value); return value; }
            /// <summary>Releases one known acquisition before its owning workbook closes.</summary>
            internal void ReleaseTracked(object value)
            {
                int index = references.FindLastIndex(item => ReferenceEquals(item, value));
                if (index < 0) throw new InvalidOperationException("The copy project acquisition is not owned.");
                try { release(value); references.RemoveAt(index); }
                catch { Retain(); throw; }
            }
            internal void Retain()
            {
                if (retained) return;
                retained = true;
                lock (retainedReferences) retainedReferences.Add(this);
            }
            public void Dispose()
            {
                if (retained || disposed) return;
                disposed = true;
                for (int index=references.Count-1; index>=0; index--)
                {
                    object value=references[index];
                    try { release(value); references[index]=null; }
                    catch { Retain(); throw; }
                }
                references.Clear();
            }
        }

        /// <summary>Restores and reads back Excel's previous event state, reporting uncertainty if verification fails.</summary>
        /// <param name="application">Owning Excel Application object.</param>
        /// <param name="expected">Event-enabled value captured before copy preparation or close.</param>
        /// <param name="retainedPath">Copy path to include in any recovery diagnostic.</param>
        private static void RestoreEvents(dynamic application, bool expected, string retainedPath)
        {
            try
            {
                application.EnableEvents = expected;
                if ((bool)application.EnableEvents != expected) throw new InvalidOperationException("Excel event state did not match its original value.");
            }
            catch (Exception error) { throw new VbaTestInvocationException("Excel event-state restoration could not be verified. Inspect the host and " + retainedPath + ". " + error.Message, true, error); }
        }
    }
}
