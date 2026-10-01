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

        internal static VbaTestCoverageClone CreateOwned(object project, string sourcePath, string folder)
        {
            using (var process = Process.GetCurrentProcess())
            {
                if (process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
                    return VbaTestWordCoverageClone.CreateWord(project, sourcePath, folder);
                if (process.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase))
                    return VbaTestPowerPointCoverageClone.CreatePowerPoint(project, sourcePath, folder);
            }
            return CreateExcel(project, sourcePath, folder);
        }

        internal static VbaTestCoverageClone CreateExcel(object project, string sourcePath, string folder)
        {
            using (var process = Process.GetCurrentProcess())
            {
                if (!process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Coverage cloning requires the in-process Excel host.");
                dynamic application = ExcelOwnedApplication.Resolve(process.Id, () => Marshal.GetActiveObject("Excel.Application"));
                uint owner;
                VbeDebugWindows.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(application.Hwnd)), out owner);
                if (owner != process.Id) throw new InvalidOperationException("The Excel application belongs to another PID.");
                object source = null;
                foreach (dynamic book in application.Workbooks)
                {
                    if (!VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, (object)book.VBProject)) continue;
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
                    if (VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, (object)((dynamic)copy).VBProject))
                        throw new InvalidOperationException("Excel did not create a distinct coverage project.");
                    return new VbaTestCoverageClone { Project = ((dynamic)copy).VBProject, Path = copyPath,
                        Close = () => {
                            bool previous = (bool)application.EnableEvents;
                            try {
                                application.EnableEvents = false;
                                if ((bool)application.EnableEvents) throw new InvalidOperationException("Excel events could not be disabled before closing the coverage copy.");
                                ((dynamic)copy).Close(false);
                                foreach (dynamic remaining in application.Workbooks)
                                    if (VbeDebug.NativeProcedureValuesHost.SameComIdentity((object)remaining, copy))
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
