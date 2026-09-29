using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Vérifie l’automatisation Excel isolée et son pont VBE sans réutiliser une session utilisateur.</summary>
    [TestClass]
    [TestCategory("Excel")]
    public sealed class ExcelHostTests
    {
        /// <summary>Résout l’identifiant de processus propriétaire d’une fenêtre Win32.</summary>
        /// <param name="window">Handle de la fenêtre.</param>
        /// <param name="processId">Reçoit l’identifiant du processus.</param>
        /// <returns>Identifiant du thread propriétaire de la fenêtre.</returns>
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        /// <summary>Démarre une instance Excel séparée, crée un classeur et vérifie le pont VBE connecté.</summary>
        [TestMethod]
        [STATestMethod]
        public void IsolatedExcelCanCreateAndReadWorkbookWithoutTouchingUserSession()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set VBAi_RUN_EXCEL_TESTS=1.");

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is unavailable.");

            var processes = Process.GetProcessesByName("EXCEL");
            var existing = processes.Select(p => p.Id).ToArray();
            foreach (var process in processes) process.Dispose();
            var root = Path.Combine(Path.GetTempPath(), "VBAi-VSTest", Guid.NewGuid().ToString("N"));
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
                ((dynamic)cell).Value2 = "VBAi été";
                Assert.AreEqual("VBAi été", (string)((dynamic)cell).Value2);
                ((dynamic)workbook).SaveAs(path, 51);
                Assert.IsTrue(File.Exists(path));
                excel.CommandBars.ExecuteMso("VisualBasic");
                var environment = VbeBridgeClient.Read((int)processId, "vbe_environment");
                Assert.IsNotNull(environment, "The native VBE command ran, but this Excel PID has no VBAi bridge.");
                Assert.AreEqual(true, environment["Ok"]);
                var fields = VbeBridgeClient.Object(VbeBridgeClient.Object(environment["Data"])["Properties"]);
                Assert.IsTrue(Convert.ToInt32(fields["ProjectCount"]) >= 1);
                var projects = VbeBridgeClient.Read((int)processId, "list_projects");
                Assert.IsNotNull(projects, "The VBAi bridge disconnected while reading Excel projects.");
                Assert.AreEqual(true, projects["Ok"]);
                Assert.IsTrue(((object[])projects["Data"]).Any(p =>
                    string.Equals(Convert.ToString(VbeBridgeClient.Object(p)["FileName"]), path, StringComparison.OrdinalIgnoreCase)));
                var addIns = VbeBridgeClient.Read((int)processId, "list_addins");
                Assert.IsNotNull(addIns, "The VBAi bridge disconnected while reading Excel VBE add-ins.");
                Assert.AreEqual(true, addIns["Ok"]);
                Assert.IsTrue(((object[])VbeBridgeClient.Object(addIns["Data"])["AddIns"]).Any(a => {
                    var properties = VbeBridgeClient.Object(VbeBridgeClient.Object(a)["Properties"]);
                    return Convert.ToString(properties["ProgId"]) == "VBAi.AddIn" &&
                        Convert.ToBoolean(properties["Connect"]);
                }));

                var debugWindows = VbeBridgeClient.Read((int)processId, "debug_windows");
                Assert.IsNotNull(debugWindows, "The VBAi bridge disconnected while reading native debug windows.");
                Assert.AreEqual(true, debugWindows["Ok"], "The native debug window snapshot failed.");
                var snapshot = VbeBridgeClient.Object(debugWindows["Data"]);
                Assert.AreEqual((int)processId, Convert.ToInt32(snapshot["HostProcessId"]));
                foreach (var pane in new[] { "Locals", "Watches", "Immediate" })
                    Assert.IsTrue(snapshot.ContainsKey(pane), "The debug snapshot omitted " + pane + ".");
                StringAssert.Contains(Convert.ToString(snapshot["Limits"]), "Native UI accessibility");
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

        /// <summary>Libère une référence COM lorsqu’elle en est une et ignore une erreur de libération COM.</summary>
        /// <param name="value">Objet COM éventuellement nul à libérer.</param>
        private static void ReleaseSafely(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
            catch (COMException) { }
        }

    }
}
