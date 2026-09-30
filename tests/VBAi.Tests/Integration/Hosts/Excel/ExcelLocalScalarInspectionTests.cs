using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelLocalScalarInspectionTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void InstalledBridgeReadsDeclaredScalarsInOwnedPausedWorkbook()
        {
            var host = ExcelVbeFixture.Start();
            string projectPath = host.File("LocalScalars.xlsm");
            string report = Path.Combine(TestContext.TestResultsDirectory, "excel-local-scalars-" + host.ProcessId + "-" + Path.GetFileName(host.Root) + ".json");
            var evidence = new ExcelScalarQualificationEvidence(report, host.ProcessId, host.Root);
            bool canClose = true;
            bool bridgeAvailable = true;
            const string module = "ScalarAudit";
            var json = new JavaScriptSerializer();
            Func<object, IDictionary<string, object>> send = request => {
                try
                {
                    var result = evidence.Send(request, () => VbeBridgeClient.Read("VBAi." + host.ProcessId, request, 20000), response => {
                        Assert.IsNotNull(response, "The owned host bridge did not respond.");
                        Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                    });
                    bridgeAvailable = true;
                    return result;
                }
                catch { bridgeAvailable = false; throw; }
            };
            Func<IDictionary<string, object>> state = () => VbeBridgeClient.Object(send(new { Command = "debug_state", Project = projectPath })["Data"]);
            evidence.Run(() =>
            {
                // Save only the new owned workbook; all VBE references stay inside the installed add-in.
                ((dynamic)UiInvoke.Field<object>(host, "workbook")).SaveAs(projectPath, 52);
                var status = send(new { Command = "status" });
                var projects = (object[])send(new { Command = "list_projects" })["Data"];
                var target = projects.Select(VbeBridgeClient.Object).Single(p =>
                    string.Equals(Convert.ToString(p["FileName"]), projectPath, StringComparison.OrdinalIgnoreCase));
                string projectName = (string)target["Name"];
                const string source = "Option Explicit\r\nPublic Sub AuditLocals()\r\n" +
                    "    Dim auditCount As Long\r\n    Dim auditText As String\r\n    Dim auditFlag As Boolean\r\n" +
                    "    Dim auditValues(1 To 2) As Long\r\n    Dim auditUnknown As Variant\r\n    Dim auditObject As Object\r\n" +
                    "    auditCount = 42\r\n    auditText = \"VBAi scalar probe\"\r\n    auditFlag = True\r\n" +
                    "    auditValues(1) = 7\r\n    auditUnknown = Empty\r\n    Stop\r\nEnd Sub";
                send(new { Command = "create_module", Project = projectPath, Module = module, ExpectedMode = 2 });
                var empty = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                send(new { Command = "replace_lines", Project = projectPath, Module = module,
                    ExpectedSha256 = empty["Sha256"], StartLine = 1, Count = 0, Text = source });
                var original = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                // Let native navigation settle before the single execution request.
                // A refused cold-start Run Sub remains a separate qualification gap.
                send(new { Command = "select_code", Project = projectPath, Module = module,
                    StartLine = 2, ExpectedSha256 = original["Sha256"] });
                var navigation = Stopwatch.StartNew();
                while (!string.Equals(Convert.ToString(state()["ActiveModule"]), module, StringComparison.Ordinal))
                {
                    if (navigation.ElapsedMilliseconds > 5000)
                        Assert.Fail("The native fixture code pane did not become active before execution.");
                    Thread.Sleep(50);
                }
                canClose = false;
                send(new { Command = "run_sub", Project = projectPath, Module = module, Procedure = "AuditLocals",
                    ExpectedMode = 2, ExpectedSha256 = original["Sha256"] });
                WaitMode(state, 1);
                send(new { Command = "open_debug_pane", Action = "locals" });
                var inspected = VbeBridgeClient.Object(send(new { Command = "inspect_local_scalars", Project = projectPath,
                    Module = module, Procedure = "AuditLocals", ExpectedMode = 1, ExpectedSha256 = original["Sha256"], Limit = 8 })["Data"]);
                Assert.AreEqual(false, inspected["Aborted"], json.Serialize(inspected));
                Assert.AreEqual(false, inspected["RuntimeInventoryComplete"]);
                Assert.AreEqual(projectName + "." + module + ".AuditLocals", inspected["Context"]);
                Assert.AreEqual(6, Convert.ToInt32(inspected["TotalCandidates"]));
                Assert.AreEqual(3, Convert.ToInt32(inspected["EligibleCandidates"]));
                var rows = ((object[])inspected["Items"]).Select(VbeBridgeClient.Object).ToDictionary(r => (string)r["Name"]);
                Assert.AreEqual("42", rows["auditCount"]["Value"]);
                Assert.AreEqual("\"VBAi scalar probe\"", rows["auditText"]["Value"]);
                Assert.IsTrue(new[] { "True", "Vrai" }.Contains(Convert.ToString(rows["auditFlag"]["Value"])), json.Serialize(inspected));
                foreach (string name in new[] { "auditCount", "auditText", "auditFlag" }) Assert.AreEqual("Read", rows[name]["Status"]);
                foreach (string name in new[] { "auditValues", "auditUnknown", "auditObject" }) Assert.AreEqual("Skipped", rows[name]["Status"]);
                Assert.AreEqual(true, inspected["SelectionRestored"], json.Serialize(inspected));
                Assert.AreEqual(true, inspected["FocusRestored"], json.Serialize(inspected));
                Assert.AreEqual(1, Convert.ToInt32(state()["Mode"]));
                var after = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                Assert.AreEqual(original["Sha256"], after["Sha256"]);
            }, () =>
            {
                if (!canClose || !bridgeAvailable)
                {
                    int mode = Convert.ToInt32(state()["Mode"]);
                    if (mode == 1)
                    {
                        send(new { Command = "debug_global", Project = projectPath, ExpectedMode = 1, Action = "reset" });
                        WaitMode(state, 2);
                        canClose = true;
                    }
                    else canClose = mode == 2;
                }
                if (canClose)
                {
                    host.Dispose();
                    evidence.Shutdown = "CloseQuitAndExitCodeZero";
                }
                else Assert.Fail("Owned Excel was left open because design mode was not verified. PID: " + host.ProcessId);
            }, () => TestContext.AddResultFile(report));
        }

        private static void WaitMode(Func<IDictionary<string, object>> state, int expected)
        {
            var deadline = Stopwatch.StartNew();
            while (Convert.ToInt32(state()["Mode"]) != expected)
            {
                if (deadline.ElapsedMilliseconds > 5000) Assert.Fail("The expected native mode was not observed: " + expected);
                Thread.Sleep(50); // Read-only observation, never a replay of a native command.
            }
        }
    }
}
