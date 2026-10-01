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
        internal object Project { get; set; }
        internal string Path { get; set; }
        internal Action Close { get; set; }
        public void Dispose() { var close = Close; Close = null; close?.Invoke(); }

        internal sealed class HostBoundary
        {
            internal Func<string> ReadProcessName = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName; };
            internal Func<int> ReadProcessId = () => { using (var process = Process.GetCurrentProcess()) return process.Id; };
            internal Func<int, object> ResolveExcel = pid => ExcelOwnedApplication.Resolve(pid, () => Marshal.GetActiveObject("Excel.Application"));
            internal Func<IntPtr, uint> ReadWindowOwner = hwnd => { uint owner; VbeDebugWindows.GetWindowThreadProcessId(hwnd, out owner); return owner; };
            internal Func<object, object, bool> SameIdentity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            internal Func<object, string, string, VbaTestCoverageClone> CreateWord = VbaTestWordCoverageClone.CreateWord;
            internal Func<object, string, string, VbaTestCoverageClone> CreatePowerPoint = VbaTestPowerPointCoverageClone.CreatePowerPoint;
        }

        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder)
        { return CreateOwned(project, sourcePath, folder, new HostBoundary()); }

        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder, HostBoundary boundary)
        {
            boundary = boundary ?? new HostBoundary();
            string name = boundary.ReadProcessName();
            if (name.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return boundary.CreateWord(project, sourcePath, folder);
            if (name.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)) return boundary.CreatePowerPoint(project, sourcePath, folder);
            return CreateExcel(project, sourcePath, folder, boundary);
        }

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
