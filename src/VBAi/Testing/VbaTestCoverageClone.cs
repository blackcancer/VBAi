using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Owned coverage workbook; only the copy's source may be instrumented.</summary>
    internal sealed class VbaTestCoverageClone : IDisposable
    {

        /// <summary>Gets or sets the project.</summary>
        /// <value>Current project exposed by vba test coverage clone.</value>
        internal object Project { get; set; }

        /// <summary>Gets or sets the path.</summary>
        /// <value>Current path exposed by vba test coverage clone.</value>
        internal string Path { get; set; }

        /// <summary>Gets or sets the close.</summary>
        /// <value>Current close exposed by vba test coverage clone.</value>
        internal Action Close { get; set; }

        /// <summary>Disposes  for vba test coverage clone.</summary>
        public void Dispose() { var close = Close; Close = null; close?.Invoke(); }

        /// <summary>Owns the host boundary state and operations.</summary>
        internal sealed class HostBoundary
        {

            /// <summary>Maintains the read process name state for host boundary.</summary>
            internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };

            /// <summary>Identifies the read process id associated with host boundary.</summary>
            internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };

            /// <summary>Maintains the resolve excel state for host boundary.</summary>
            internal Func<int, object> ResolveExcel = pid => ExcelOwnedApplication.Resolve(pid, () => Marshal.GetActiveObject("Excel.Application"));

            /// <summary>Maintains the read window owner state for host boundary.</summary>
            internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };

            /// <summary>Maintains the same identity state for host boundary.</summary>
            internal Func<object, object, bool> SameIdentity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;

            /// <summary>Maintains the create word state for host boundary.</summary>
            internal Func<object, string, string, VbaTestCoverageClone> CreateWord = VbaTestWordCoverageClone.CreateWord;

            /// <summary>Maintains the create power point state for host boundary.</summary>
            internal Func<object, string, string, VbaTestCoverageClone> CreatePowerPoint = VbaTestPowerPointCoverageClone.CreatePowerPoint;
        }

        /// <summary>Creates owned for vba test coverage clone.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="sourcePath">Path used for the source path being processed.</param>
        /// <param name="folder">Text that supplies the folder value. Use the format required by the calling operation.</param>
        /// <returns>vba test coverage clone produced by the operation for create owned on vba test coverage clone.</returns>
        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder)
        { return CreateOwned(project, sourcePath, folder, new HostBoundary()); }

        /// <summary>Creates owned for vba test coverage clone.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="sourcePath">Path used for the source path being processed.</param>
        /// <param name="folder">Text that supplies the folder value. Use the format required by the calling operation.</param>
        /// <param name="boundary">host boundary that supplies the boundary for this operation.</param>
        /// <returns>vba test coverage clone produced by the operation for create owned on vba test coverage clone.</returns>
        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder, HostBoundary boundary)
        {
            boundary = boundary ?? new HostBoundary();
            string name = boundary.ReadProcessName();
            if (name.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return boundary.CreateWord(project, sourcePath, folder);
            if (name.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)) return boundary.CreatePowerPoint(project, sourcePath, folder);
            return CreateExcel(project, sourcePath, folder, boundary);
        }

        /// <summary>Creates excel for vba test coverage clone.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="sourcePath">Path used for the source path being processed.</param>
        /// <param name="folder">Text that supplies the folder value. Use the format required by the calling operation.</param>
        /// <param name="boundary">host boundary that supplies the boundary for this operation.</param>
        /// <returns>vba test coverage clone produced by the operation for create excel on vba test coverage clone.</returns>
        internal static VbaTestCoverageClone CreateExcel(object project, string sourcePath, string folder, HostBoundary boundary = null)
        {
            boundary = boundary ?? new HostBoundary();
            {
                int processId = boundary.ReadProcessId();
                if (!boundary.ReadProcessName().Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Coverage cloning requires the in-process Excel host.");
                dynamic application = boundary.ResolveExcel(processId);
                uint owner = boundary.ReadWindowOwner(new IntPtr(Convert.ToInt64(application.Hwnd)));
                if (owner != processId) throw new InvalidOperationException("The Excel application belongs to another PID.");
                object source = null;
                foreach (dynamic book in application.Workbooks)
                {
                    if (!boundary.SameIdentity(project, (object)book.VBProject)) continue;
                    if (source != null) throw new InvalidOperationException("The source workbook identity is ambiguous.");
                    if (!string.Equals(System.IO.Path.GetFullPath((string)book.FullName), System.IO.Path.GetFullPath(sourcePath), StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The source workbook path changed.");
                    source = book;
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
                    copy = application.Workbooks.Open(copyPath, UpdateLinks: 0, ReadOnly: false, AddToMru: false);
                    if (boundary.SameIdentity(project, (object)((dynamic)copy).VBProject))
                        throw new InvalidOperationException("Excel did not create a distinct coverage project.");
                    return new VbaTestCoverageClone { Project = ((dynamic)copy).VBProject, Path = copyPath,
                        Close = () => {
                            bool previous = (bool)application.EnableEvents;
                            try {
                                application.EnableEvents = false;
                                if ((bool)application.EnableEvents) throw new InvalidOperationException("Excel events could not be disabled before closing the coverage copy.");
                                ((dynamic)copy).Close(false);
                                foreach (dynamic remaining in application.Workbooks)
                                    if (boundary.SameIdentity((object)remaining, copy))
                                        throw new InvalidOperationException("The owned copy remains open after Close returned.");
                            }
                            catch (Exception error) { throw new VbaTestInvocationException("Closing the owned workbook copy is uncertain. Retained path: " + copyPath + ". " + error.Message, true, error); }
                            finally { RestoreEvents((object)application, previous, copyPath); }
                        } };
                }
                catch (Exception error)
                {
                    throw new VbaTestInvocationException("Coverage copy preparation is uncertain; no close or retry was attempted. Retained copy: " + copyPath + ". " + error.Message, true, error);
                }
                finally { RestoreEvents((object)application, events, copyPath); }
            }
        }

        /// <summary>Handles restore events for vba test coverage clone.</summary>
        /// <param name="application">dynamic that supplies the application for this operation.</param>
        /// <param name="expected">Indicates whether expected is enabled.</param>
        /// <param name="retainedPath">Path used for the retained path being processed.</param>
        private static void RestoreEvents(dynamic application, bool expected, string retainedPath)
        {
            try {
                application.EnableEvents = expected;
                if ((bool)application.EnableEvents != expected) throw new InvalidOperationException("Excel event state did not match its original value.");
            }
            catch (Exception error) { throw new VbaTestInvocationException("Excel event-state restoration could not be verified. Inspect the host and " + retainedPath + ". " + error.Message, true, error); }
        }
    }
}
