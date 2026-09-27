using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    [TestClass]
    [TestCategory("Excel")]
    public sealed class ExcelHostTests
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        [TestMethod]
        [STATestMethod]
        public void IsolatedExcelCanCreateAndReadWorkbookWithoutTouchingUserSession()
        {
            if (Environment.GetEnvironmentVariable("CODEXVBE_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set CODEXVBE_RUN_EXCEL_TESTS=1.");

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is unavailable.");

            var processes = Process.GetProcessesByName("EXCEL");
            var existing = processes.Select(p => p.Id).ToArray();
            foreach (var process in processes) process.Dispose();
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-VSTest", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, "isolated.xlsx");
            object application = null, books = null, workbook = null, sheet = null, cell = null;
            uint processId = 0;
            bool owned = false;
            try
            {
                application = Activator.CreateInstance(excelType);
                dynamic excel = application;
                GetWindowThreadProcessId(new IntPtr((int)excel.Hwnd), out processId);
                owned = processId != 0 && !existing.Contains((int)processId);
                if (!owned)
                    Assert.Inconclusive("Excel returned an existing session; no workbook was opened.");

                excel.Visible = true;
                excel.DisplayAlerts = false;
                books = excel.Workbooks;
                workbook = ((dynamic)books).Add();
                sheet = ((dynamic)workbook).Worksheets[1];
                cell = ((dynamic)sheet).Cells[1, 1];
                ((dynamic)cell).Value2 = "CodexVBE été";
                Assert.AreEqual("CodexVBE été", (string)((dynamic)cell).Value2);
                ((dynamic)workbook).SaveAs(path, 51);
                Assert.IsTrue(File.Exists(path));
                excel.CommandBars.ExecuteMso("VisualBasic");
                var environment = Bridge((int)processId, "vbe_environment");
                Assert.IsNotNull(environment, "The native VBE command ran, but this Excel PID has no CodexVBE bridge.");
                Assert.AreEqual(true, environment["Ok"]);
                var fields = Obj(Obj(environment["Data"])["Properties"]);
                Assert.IsTrue(Convert.ToInt32(fields["ProjectCount"]) >= 1);
                var projects = Obj(Bridge((int)processId, "list_projects"));
                Assert.AreEqual(true, projects["Ok"]);
                Assert.IsTrue(((object[])projects["Data"]).Any(p =>
                    string.Equals(Convert.ToString(Obj(p)["FileName"]), path, StringComparison.OrdinalIgnoreCase)));
                var addIns = Obj(Bridge((int)processId, "list_addins"));
                Assert.AreEqual(true, addIns["Ok"]);
                Assert.IsTrue(((object[])Obj(addIns["Data"])["AddIns"]).Any(a => {
                    var properties = Obj(Obj(a)["Properties"]);
                    return Convert.ToString(properties["ProgId"]) == "CodexVBE.AddIn" &&
                        Convert.ToBoolean(properties["Connect"]);
                }));
            }
            finally
            {
                if (owned && workbook != null)
                    try { ((dynamic)workbook).Close(false); } catch { }
                if (owned && application != null)
                    try { ((dynamic)application).Quit(); } catch { }
                ReleaseSafely(cell); ReleaseSafely(sheet); ReleaseSafely(workbook); ReleaseSafely(books); ReleaseSafely(application);
                if (owned && processId != 0)
                {
                    try { using (var process = Process.GetProcessById((int)processId)) {
                        if (!process.WaitForExit(10000)) { process.Kill(); process.WaitForExit(10000); }
                    } }
                    catch (Exception) { }
                }
                try { if (Directory.Exists(root)) Directory.Delete(root, true); } catch (Exception) { }
            }
        }

        private static void ReleaseSafely(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
            catch (COMException) { }
        }

        private static System.Collections.Generic.IDictionary<string, object> Obj(object value)
        {
            return (System.Collections.Generic.IDictionary<string, object>)value;
        }

        private static System.Collections.Generic.IDictionary<string, object> Bridge(int processId, string command)
        {
            var json = new JavaScriptSerializer();
            for (int attempt = 0; attempt < 20; attempt++)
            {
                try
                {
                    using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                    {
                        pipe.Connect(250);
                        using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                        {
                            writer.WriteLine(json.Serialize(new { Command = command }));
                            return Obj(json.DeserializeObject(reader.ReadLine()));
                        }
                    }
                }
                catch (TimeoutException) { Thread.Sleep(250); }
                catch (IOException) { Thread.Sleep(250); }
            }
            return null;
        }
    }
}
