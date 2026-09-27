using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    internal sealed class ExcelVbeFixture : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        private object application;
        private object workbooks;
        private object workbook;
        private bool owned;

        private ExcelVbeFixture() { }

        internal int ProcessId { get; private set; }
        internal string Root { get; private set; }

        internal static ExcelVbeFixture Start()
        {
            if (Environment.GetEnvironmentVariable("CODEXVBE_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set CODEXVBE_RUN_EXCEL_TESTS=1.");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is unavailable.");
            var existing = Process.GetProcessesByName("EXCEL");
            int[] existingIds = existing.Select(process => process.Id).ToArray();
            foreach (var process in existing) process.Dispose();
            var fixture = new ExcelVbeFixture {
                Root = Path.Combine(Path.GetTempPath(), "CodexVBE-VSTest", Guid.NewGuid().ToString("N")) };
            try
            {
                Directory.CreateDirectory(fixture.Root);
                fixture.application = Activator.CreateInstance(excelType);
                dynamic excel = fixture.application;
                uint processId;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out processId);
                fixture.ProcessId = (int)processId;
                fixture.owned = processId != 0 && !existingIds.Contains((int)processId);
                if (!fixture.owned)
                    Assert.Inconclusive("Excel returned an existing session; no workbook was opened.");
                excel.Visible = true;
                excel.DisplayAlerts = false;
                fixture.workbooks = excel.Workbooks;
                fixture.workbook = ((dynamic)fixture.workbooks).Add();
                excel.CommandBars.ExecuteMso("VisualBasic");
                var status = fixture.Command("status");
                Assert.IsNotNull(status, "The isolated Excel VBE has no CodexVBE bridge.");
                Assert.AreEqual(true, status["Ok"]);
                return fixture;
            }
            catch
            {
                fixture.Dispose();
                throw;
            }
        }

        internal IDictionary<string, object> Command(string name)
        {
            return VbeBridgeClient.Read(ProcessId, name);
        }

        internal IDictionary<string, object> Command(object request)
        {
            return VbeBridgeClient.Read(ProcessId, request);
        }

        internal string File(string name) { return Path.Combine(Root, name); }

        public void Dispose()
        {
            if (owned && workbook != null)
                try { ((dynamic)workbook).Close(false); } catch { }
            if (owned && application != null)
                try { ((dynamic)application).Quit(); } catch { }
            Release(workbook);
            Release(workbooks);
            Release(application);
            workbook = workbooks = application = null;
            if (owned && ProcessId != 0)
            {
                try
                {
                    using (var process = Process.GetProcessById(ProcessId))
                        if (!process.WaitForExit(10000)) { process.Kill(); process.WaitForExit(10000); }
                }
                catch (ArgumentException) { }
            }
            if (string.IsNullOrWhiteSpace(Root) || !Directory.Exists(Root)) return;
            try
            {
                foreach (string file in Directory.GetFiles(Root, "*", SearchOption.TopDirectoryOnly))
                    System.IO.File.Delete(file);
                Directory.Delete(Root, false);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void Release(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
            catch (COMException) { }
        }
    }
}
