using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Reads exact owned-host and form identities without exporting or invoking macros.</summary>
        internal IDictionary<string, object> ReadGitExportContext(string form)
        {
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            object project = null, components = null, component = null, editor = null, mainWindow = null;
            try
            {
                uint excelPid, editorPid;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out excelPid);
                editor = ((dynamic)application).VBE; mainWindow = ((dynamic)editor).MainWindow;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)mainWindow).HWnd)), out editorPid);
                Assert.AreEqual((uint)ProcessId, excelPid); Assert.AreEqual(excelPid, editorPid);
                project = ((dynamic)workbook).VBProject;
                string hostPath = Path.GetFullPath(Convert.ToString(((dynamic)workbook).FullName));
                Assert.AreEqual(hostPath, Path.GetFullPath(VbeProjectHostPath.Read(project)), true);
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode));
                Assert.AreEqual(0, Convert.ToInt32(((dynamic)project).Protection));
                components = ((dynamic)project).VBComponents; component = ((dynamic)components).Item(form);
                Assert.AreEqual(form, ((dynamic)component).Name); Assert.AreEqual(3, Convert.ToInt32(((dynamic)component).Type));
                return new Dictionary<string, object> {
                    ["HostProcessId"] = ProcessId, ["EditorProcessId"] = editorPid,
                    ["ProjectName"] = ((dynamic)project).Name, ["HostPath"] = hostPath,
                    ["ComponentName"] = ((dynamic)component).Name, ["ComponentType"] = ((dynamic)component).Type,
                    ["Mode"] = ((dynamic)project).Mode, ["Protection"] = ((dynamic)project).Protection,
                    ["HasOpenDesigner"] = ((dynamic)component).HasOpenDesigner,
                    ["Apartment"] = Thread.CurrentThread.GetApartmentState().ToString()
                };
            }
            finally { Release(mainWindow); Release(editor); Release(component); Release(components); Release(project); }
        }

        /// <summary>Performs exactly one external native Export call after the owned identity checks.</summary>
        internal void ExportGitFormOnce(string form, string destination)
        {
            ReadGitExportContext(form);
            Assert.IsFalse(System.IO.File.Exists(destination));
            Assert.IsFalse(System.IO.File.Exists(Path.ChangeExtension(destination, ".frx")));
            object project = null, components = null, component = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents; component = ((dynamic)components).Item(form);
                ((dynamic)component).Export(destination);
            }
            finally { Release(component); Release(components); Release(project); }
        }
    }
}
