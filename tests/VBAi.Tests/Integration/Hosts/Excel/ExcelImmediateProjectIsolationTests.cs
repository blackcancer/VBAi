using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Excel"), TestCategory("ExcelImmediate"), DoNotParallelize]
    public sealed class ExcelImmediateProjectIsolationTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void InstalledBridgeRejectsOtherActiveProjectAndExecutesOnlyMatchingDisposableWorkbook()
        {
            // ExcelVbeFixture requires VBAi_RUN_EXCEL_TESTS=1 and verifies the loaded assembly MVID.
            var host = ExcelVbeFixture.Start();
            object second = null;
            var evidence = new List<object>();
            var json = new JavaScriptSerializer();
            string firstPath = host.File("ImmediateFirst.xlsm"), secondPath = host.File("ImmediateSecond.xlsm");
            const string module = "ImmediateIsolation";
            const string statement = "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1";
            Func<object, IDictionary<string, object>> request = value => {
                var timer = Stopwatch.StartNew();
                var response = VbeBridgeClient.Read("VBAi." + host.ProcessId, value, 20000);
                evidence.Add(new { Request = value, Response = response, ElapsedMilliseconds = timer.ElapsedMilliseconds });
                Assert.IsNotNull(response, "The owned Excel bridge did not respond.");
                return response;
            };
            Func<object, IDictionary<string, object>> success = value => {
                var response = request(value);
                Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                return response;
            };
            try
            {
                object first = UiInvoke.Field<object>(host, "workbook");
                ((dynamic)first).SaveAs(firstPath, 52);
                second = ((dynamic)UiInvoke.Field<object>(host, "workbooks")).Add();
                ((dynamic)second).SaveAs(secondPath, 52);
                WriteCell(first, 100);
                WriteCell(second, 200);
                success(new { Command = "status" });
                var projects = ((object[])success(new { Command = "list_projects" })["Data"])
                    .Select(VbeBridgeClient.Object).ToArray();
                string firstName = Convert.ToString(projects.Single(p => SamePath(p["FileName"], firstPath))["Name"]);
                string secondName = Convert.ToString(projects.Single(p => SamePath(p["FileName"], secondPath))["Name"]);
                Assert.AreEqual(firstName, secondName, "This adversarial fixture requires identical default VBA project names and distinct paths.");
                var hashes = new Dictionary<string, object>();
                foreach (string path in new[] { firstPath, secondPath })
                {
                    success(new { Command = "create_module", Project = path, Module = module, ExpectedMode = 2 });
                    var empty = VbeBridgeClient.Object(success(new { Command = "read_module", Project = path, Module = module })["Data"]);
                    success(new { Command = "replace_lines", Project = path, Module = module,
                        ExpectedSha256 = empty["Sha256"], StartLine = 1, Count = 0,
                        Text = "Option Explicit\r\n' Disposable Immediate project isolation fixture." });
                    hashes[path] = VbeBridgeClient.Object(success(new { Command = "read_module", Project = path, Module = module })["Data"])["Sha256"];
                }
                success(new { Command = "open_debug_pane", Action = "immediate" });
                Action<string> select = path => {
                    success(new { Command = "select_code", Project = path, Module = module,
                        StartLine = 1, ExpectedSha256 = hashes[path] });
                    var timer = Stopwatch.StartNew();
                    while (true)
                    {
                        var state = VbeBridgeClient.Object(success(new { Command = "debug_state", Project = path })["Data"]);
                        Assert.AreEqual(2, Convert.ToInt32(state["Mode"]));
                        if (SamePath(state["SelectedProjectPath"], path) && Convert.ToString(state["ActiveModule"]) == module) break;
                        Assert.IsTrue(timer.ElapsedMilliseconds < 5000, "The native code pane did not settle on " + path);
                        Thread.Sleep(50); // Read-only observation; native navigation is never replayed.
                    }
                };
                Action<string, int, int> reject = (path, expectedFirst, expectedSecond) => {
                    var state = VbeBridgeClient.Object(success(new { Command = "debug_state", Project = path })["Data"]);
                    Assert.IsNull(state["SelectedProjectPath"], "The requested project must differ from the active code pane by COM identity.");
                    var denied = request(new { Command = "immediate_execute", Project = path, ExpectedMode = 2, Text = statement });
                    Assert.AreEqual(false, denied["Ok"], json.Serialize(denied));
                    StringAssert.Contains(Convert.ToString(denied["Error"]), "requested project must be active");
                    // An independent Excel readback checks both possible recipients, including delayed native input.
                    var observation = Stopwatch.StartNew();
                    do
                    {
                        Assert.AreEqual(expectedFirst, Convert.ToInt32(host.ReadCell("A1")));
                        Assert.AreEqual(expectedSecond, Convert.ToInt32(ReadCell(second)));
                        Thread.Sleep(25);
                    } while (observation.ElapsedMilliseconds < 500);
                    evidence.Add(new { RejectedProject = path, FirstCell = host.ReadCell("A1"), SecondCell = ReadCell(second) });
                };
                Action<string, object, int> execute = (path, book, expected) => {
                    var data = VbeBridgeClient.Object(success(new { Command = "immediate_execute", Project = path,
                        ExpectedMode = 2, Text = statement })["Data"]);
                    Assert.AreEqual(true, data["CommandEchoObserved"]);
                    var timer = Stopwatch.StartNew();
                    while (Convert.ToInt32(ReadCell(book)) != expected)
                    {
                        Assert.IsTrue(timer.ElapsedMilliseconds < 5000, "Immediate side effect was not independently observed. " + json.Serialize(data));
                        Thread.Sleep(50); // Observe only. Never repeat an uncertain Immediate mutation.
                    }
                    evidence.Add(new { ExecutedProject = path, FirstCell = host.ReadCell("A1"), SecondCell = ReadCell(second) });
                };

                select(secondPath);
                reject(firstPath, 100, 200);
                select(firstPath);
                execute(firstPath, first, 101);
                reject(secondPath, 101, 200);
                select(secondPath);
                execute(secondPath, second, 201);
                Assert.AreEqual(101, Convert.ToInt32(host.ReadCell("A1")));
                Assert.AreEqual(201, Convert.ToInt32(ReadCell(second)));
                foreach (string path in new[] { firstPath, secondPath })
                {
                    var after = VbeBridgeClient.Object(success(new { Command = "read_module", Project = path, Module = module })["Data"]);
                    Assert.AreEqual(hashes[path], after["Sha256"], "Immediate execution must preserve fixture source.");
                    var state = VbeBridgeClient.Object(success(new { Command = "debug_state", Project = path })["Data"]);
                    Assert.AreEqual(2, Convert.ToInt32(state["Mode"]));
                }
            }
            finally
            {
                try
                {
                    if (second != null)
                    {
                        try { ((dynamic)second).Close(false); }
                        finally { Marshal.FinalReleaseComObject(second); }
                    }
                }
                finally
                {
                    try
                    {
                        host.Dispose();
                        evidence.Add(new { Shutdown = "CloseQuitAndExitCodeZero", HostProcessId = host.ProcessId });
                    }
                    finally
                    {
                        Directory.CreateDirectory(TestContext.TestResultsDirectory);
                        string report = Path.Combine(TestContext.TestResultsDirectory, "excel-immediate-isolation-" + host.ProcessId + ".json");
                        File.WriteAllText(report, json.Serialize(new { HostProcessId = host.ProcessId,
                            AssemblyMvid = typeof(VBAi.VbeSession).Module.ModuleVersionId,
                            Scope = "Installed native bridge, same-name disposable Excel projects, independent worksheet side effects. Chat UI is not exercised.",
                            Evidence = evidence }));
                        TestContext.AddResultFile(report);
                    }
                }
            }
        }

        private static bool SamePath(object value, string path) =>
            string.Equals(Convert.ToString(value), path, StringComparison.OrdinalIgnoreCase);

        private static object ReadCell(object workbook) => Cell(workbook, null);
        private static void WriteCell(object workbook, int value) => Cell(workbook, value);

        private static object Cell(object workbook, int? value)
        {
            object sheets = null, sheet = null, cell = null;
            try
            {
                sheets = ((dynamic)workbook).Worksheets;
                sheet = ((dynamic)sheets)[1];
                cell = ((dynamic)sheet).Range["A1"];
                if (value.HasValue) ((dynamic)cell).Value2 = value.Value;
                return ((dynamic)cell).Value2;
            }
            finally
            {
                foreach (object item in new[] { cell, sheet, sheets })
                    if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
            }
        }
    }
}
