using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Reads the actual VBE window owning native TID, independently of the external testhost STA.</summary>
        internal uint ReadPairedExportOwnerThread()
        {
            object editor = null, window = null;
            try
            {
                editor = ((dynamic)application).VBE; window = ((dynamic)editor).MainWindow;
                uint pid;
                uint tid = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).HWnd)), out pid);
                Assert.AreEqual((uint)ProcessId, pid); Assert.AreNotEqual(0u, tid);
                return tid;
            }
            finally { Release(window); Release(editor); }
        }

        /// <summary>Independently reads the sole synthetic Label without creating, saving, exporting or executing a macro.</summary>
        internal IDictionary<string, object> ReadPairedExportLabel(string form)
        {
            object project = null, components = null, component = null, designer = null, controls = null, label = null;
            try
            {
                project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form); designer = ((dynamic)component).Designer;
                controls = ((dynamic)designer).Controls;
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)controls).Count));
                label = ((dynamic)controls).Item("SyntheticLabel");
                return new Dictionary<string, object>
                {
                    ["Name"] = ((dynamic)label).Name,
                    ["Caption"] = ((dynamic)label).Caption,
                    ["Left"] = Convert.ToDouble(((dynamic)label).Left),
                    ["Top"] = Convert.ToDouble(((dynamic)label).Top),
                    ["Width"] = Convert.ToDouble(((dynamic)label).Width),
                    ["Height"] = Convert.ToDouble(((dynamic)label).Height)
                };
            }
            finally { Release(label); Release(controls); Release(designer); Release(component); Release(components); Release(project); }
        }
    }
}
