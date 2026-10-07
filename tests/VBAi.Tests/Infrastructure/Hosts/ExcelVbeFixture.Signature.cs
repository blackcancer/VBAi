using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Emits one signing request with enough time for the native protected certificate picker.</summary>
        internal IDictionary<string, object> SignOnce(object request)
        {
            bool terminal = false;
            try
            {
                return RecordCommand(request, () =>
                {
                    var reply = VbeBridgeClient.Read("VBAi." + ProcessId, request, 180000);
                    terminal = reply != null;
                    return reply;
                });
            }
            finally { if (!terminal) PreserveForDiagnosticRecovery = true; }
        }

        /// <summary>Opens only a saved synthetic workbook read-only, with events and macros disabled.</summary>
        internal void OpenOwnedReadOnlyWorkbook(string path)
        {
            path = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(path);
            Assert.IsTrue(System.IO.File.Exists(path));
            Assert.AreEqual(".xlsm", Path.GetExtension(path), true);
            ((dynamic)application).EnableEvents = false;
            ((dynamic)application).AutomationSecurity = 3;
            bool terminal = false;
            try
            {
                ((dynamic)workbook).Close(false);
                Release(workbook); workbook = null;
                workbook = ((dynamic)workbooks).Open(path, 0, true);
                Assert.AreEqual(path, Convert.ToString(((dynamic)workbook).FullName), true);
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)workbooks).Count));
                terminal = true;
            }
            finally { if (!terminal) PreserveForDiagnosticRecovery = true; }
        }
    }
}
