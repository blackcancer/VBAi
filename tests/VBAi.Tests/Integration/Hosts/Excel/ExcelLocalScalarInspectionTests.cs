using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelLocalScalarInspectionTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeSkipsUnsupportedScalarPageWithoutQuickWatch()
        {
            RunDiagnosticPage(3, 3);
        }

        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsOneLongScalarWithNativePhaseEvidence()
        {
            RunDiagnosticPage(0, 1);
        }

        /// <summary>Inspects the complete declared page once, retaining native phase and shutdown evidence.</summary>
        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsFullScalarPageWithThreeNativeObservers()
        {
            RunDiagnosticPage(0, 6);
        }

        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsEverySupportedDeclaredScalarTypeWithNativePhaseEvidence()
        {
            RunDiagnosticPage(0, 14, true);
        }

        // Each page stays below the installed trace's 128-event bound. The full-page
        // diagnostic above remains available and must not pass without terminal evidence.
        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsSupportedScalarFirstPageWithNativePhaseEvidence()
        {
            RunDiagnosticPage(0, 4, true);
        }

        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsSupportedScalarSecondPageWithNativePhaseEvidence()
        {
            RunDiagnosticPage(4, 4, true);
        }

        [STATestMethod, TestCategory("ExcelScalarDiagnostics")]
        public void InstalledBridgeReadsSupportedScalarLastPageWithNativePhaseEvidence()
        {
            RunDiagnosticPage(8, 6, true);
        }

        /// <summary>Uses one owned host and one inspection request; uncertain native work retains the host.</summary>
        private void RunDiagnosticPage(int offset, int limit, bool allTypes = false)
        {
            string tracePath = Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName);
            if (string.IsNullOrWhiteSpace(tracePath))
                Assert.Inconclusive("Scalar diagnostics require VBAi_VBE_INSPECTION_TRACE in the host environment before launch.");
            Assert.IsTrue(Path.IsPathRooted(tracePath) && tracePath.Length > 2 && tracePath[1] == ':' &&
                (tracePath[2] == '\\' || tracePath[2] == '/'), "An absolute local phase-evidence path is required.");
            var phaseStartedUtc = DateTime.UtcNow;
            var host = ExcelVbeFixture.StartOwnedWithTrace(tracePath);
            string projectPath = host.File("ScalarPage" + offset + ".xlsm");
            string reportDirectory = string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"))
                ? TestContext.TestResultsDirectory : host.Root;
            string report = Path.Combine(reportDirectory, "excel-scalar-page-" + offset + "-" + host.ProcessId + "-" + Path.GetFileName(host.Root) + ".json");
            var evidence = new ExcelScalarQualificationEvidence(report, host.ProcessId, host.Root) { PhaseTracePath = tracePath };
            const string module = "ScalarPageAudit", procedure = "AuditPage";
            string source = "Option Explicit\r\nPublic Sub AuditPage()\r\n" +
                "    Dim auditCount As Long\r\n    Dim auditText As String\r\n    Dim auditFlag As Boolean\r\n" +
                "    Dim auditValues(1 To 2) As Long\r\n    Dim auditUnknown As Variant\r\n    Dim auditObject As Object\r\n" +
                "    auditCount = 42\r\n    auditText = \"VBAi scalar page probe\"\r\n    auditFlag = True\r\n" +
                "    auditValues(1) = 7\r\n    auditUnknown = Empty\r\n    Stop\r\nEnd Sub";
            string[] names = { "auditCount", "auditText", "auditFlag", "auditValues", "auditUnknown", "auditObject" };
            string[] types = { "Long", "String", "Boolean", "Long", "Variant", "Object" };
            string[] skipReasons = { null, null, null, "ArrayDeclaration", "NonScalarOrVariantType", "NonScalarOrVariantType" };
            int eligible = 3;
            if (allTypes)
            {
                names = new[] { "auditCount", "auditText", "auditFlag", "auditByte", "auditInteger", "auditLongLong", "auditLongPtr", "auditSingle", "auditDouble", "auditCurrency", "auditDate", "auditValues", "auditUnknown", "auditObject" };
                types = new[] { "Long", "String", "Boolean", "Byte", "Integer", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "Long", "Variant", "Object" };
                eligible = 11;
                skipReasons = Enumerable.Repeat<string>(null, eligible).Concat(new[] { "ArrayDeclaration", "NonScalarOrVariantType", "NonScalarOrVariantType" }).ToArray();
                source = "Option Explicit\r\nPublic Sub AuditPage()\r\n" + string.Join("\r\n", names.Select((name, index) =>
                    "    Dim " + name + (index == eligible ? "(1 To 2)" : "") + " As " + types[index])) +
                    "\r\n    auditCount = 42\r\n    auditText = \"VBAi scalar page probe\"\r\n    auditFlag = True" +
                    "\r\n    auditByte = 200\r\n    auditInteger = -1234\r\n    auditLongLong = 4294967296^\r\n    auditLongPtr = 4294967296^" +
                    "\r\n    auditSingle = 1.5\r\n    auditDouble = 1.25\r\n    auditCurrency = 12.5\r\n    auditDate = DateSerial(2026, 11, 23)" +
                    "\r\n    auditValues(1) = 7\r\n    auditUnknown = Empty\r\n    Stop\r\nEnd Sub";
            }
            var json = new JavaScriptSerializer();
            bool bridgeAvailable = true, inspectionPending = false, executionPending = false, canClose = true;
            Func<object, IDictionary<string, object>> send = request =>
            {
                try
                {
                    var response = evidence.Send(request, () => VbeBridgeClient.Read("VBAi." + host.ProcessId, request, 20000), value =>
                    {
                        Assert.IsNotNull(value, "The owned bridge did not return; preserve host and evidence.");
                        Assert.AreEqual(true, value["Ok"], json.Serialize(value));
                    });
                    bridgeAvailable = true;
                    return response;
                }
                catch { bridgeAvailable = false; throw; }
            };
            Func<IDictionary<string, object>> state = () => VbeBridgeClient.Object(send(new { Command = "debug_state", Project = projectPath })["Data"]);
            evidence.Run(() =>
            {
                ((dynamic)UiInvoke.Field<object>(host, "workbook")).SaveAs(projectPath, 52);
                var status = VbeBridgeClient.Object(send(new { Command = "status" })["Data"]);
                Assert.AreEqual(host.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                var projects = (object[])send(new { Command = "list_projects" })["Data"];
                var target = projects.Select(VbeBridgeClient.Object).Single(p => string.Equals(Convert.ToString(p["FileName"]), projectPath, StringComparison.OrdinalIgnoreCase));
                string projectName = (string)target["Name"];
                Assert.AreEqual(2, Convert.ToInt32(target["Mode"]));
                send(new { Command = "create_module", Project = projectPath, Module = module, ExpectedMode = 2 });
                var empty = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                send(new { Command = "replace_lines", Project = projectPath, Module = module, ExpectedSha256 = empty["Sha256"], StartLine = 1, Count = 0, Text = source });
                var original = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                send(new { Command = "select_code", Project = projectPath, Module = module, StartLine = 2, ExpectedSha256 = original["Sha256"] });
                var navigation = Stopwatch.StartNew();
                while (!string.Equals(Convert.ToString(state()["ActiveModule"]), module, StringComparison.Ordinal))
                {
                    if (navigation.ElapsedMilliseconds > 5000) Assert.Fail("Owned native module navigation did not settle.");
                    Thread.Sleep(50);
                }
                canClose = false;
                executionPending = true;
                send(new { Command = "run_sub", Project = projectPath, Module = module, Procedure = procedure, ExpectedMode = 2, ExpectedSha256 = original["Sha256"] });
                WaitMode(state, 1);
                executionPending = false;
                send(new { Command = "open_debug_pane", Action = "locals" });
                var before = state();
                Assert.AreEqual(1, Convert.ToInt32(before["Mode"]));
                Assert.AreEqual(projectPath, before["SelectedProjectPath"]);
                Assert.AreEqual(module, before["ActiveModule"]);
                inspectionPending = true;
                // Exactly one attempt. Even a client timeout never authorizes a second inspection.
                var inspected = VbeBridgeClient.Object(send(new
                {
                    Command = "inspect_local_scalars",
                    Project = projectPath,
                    Module = module,
                    Procedure = procedure,
                    ExpectedMode = 1,
                    ExpectedSha256 = original["Sha256"],
                    Offset = offset,
                    Limit = limit
                })["Data"]);
                Assert.IsNotNull(inspected);
                Assert.IsTrue(inspected.ContainsKey("Aborted"), "A malformed response is not terminal native evidence.");
                inspectionPending = false;
                Assert.AreEqual(false, inspected["Aborted"], json.Serialize(inspected));
                Assert.AreEqual(projectName, inspected["Project"]);
                Assert.AreEqual(module, inspected["Module"]);
                Assert.AreEqual(procedure, inspected["Procedure"]);
                Assert.AreEqual(projectName + "." + module + "." + procedure, inspected["Context"]);
                Assert.AreEqual(original["Sha256"], inspected["Sha256"]);
                Assert.AreEqual("DeclaredScalarCandidatesOnly", inspected["Coverage"]);
                Assert.IsNull(inspected["Error"], "No partial scalar failure may be accepted as a completed page.");
                Assert.AreEqual(false, inspected["RuntimeInventoryComplete"]);
                Assert.AreEqual(names.Length, Convert.ToInt32(inspected["TotalCandidates"]));
                Assert.AreEqual(eligible, Convert.ToInt32(inspected["EligibleCandidates"]));
                Assert.AreEqual(offset, Convert.ToInt32(inspected["Offset"]));
                Assert.AreEqual(limit, Convert.ToInt32(inspected["Items"] == null ? 0 : ((object[])inspected["Items"]).Length));
                if (offset + limit >= names.Length) Assert.IsNull(inspected["NextOffset"]);
                else Assert.AreEqual(offset + limit, Convert.ToInt32(inspected["NextOffset"]));
                var rows = ((object[])inspected["Items"]).Select(VbeBridgeClient.Object).ToArray();
                CollectionAssert.AreEqual(names.Skip(offset).Take(limit).ToArray(), rows.Select(row => (string)row["Name"]).ToArray());
                for (int index = 0; index < rows.Length; index++)
                {
                    var row = rows[index]; int candidate = offset + index;
                    Assert.AreEqual(names[candidate], row["Expression"]);
                    Assert.AreEqual("Variable", row["Kind"]);
                    Assert.AreEqual(types[candidate], row["TypeName"]);
                    Assert.AreEqual(candidate + 3, Convert.ToInt32(row["Line"]));
                    Assert.AreEqual(9, Convert.ToInt32(row["Column"]));
                    Assert.IsNull(row["Error"], names[candidate]);
                    Assert.AreEqual(skipReasons[candidate], row["SkipReason"], names[candidate]);
                    if (candidate >= eligible)
                    {
                        Assert.AreEqual("Skipped", row["Status"]);
                        Assert.IsNull(row["Value"]);
                    }
                    else
                    {
                        Assert.AreEqual("Read", row["Status"]);
                        if (candidate == 0) Assert.AreEqual("42", row["Value"]);
                        if (candidate == 1) Assert.AreEqual("\"VBAi scalar page probe\"", row["Value"]);
                        // The installed French VBE can display True as Vrai; neither False nor a missing value passes.
                        if (candidate == 2) Assert.IsTrue(new[] { "True", "Vrai" }.Contains(Convert.ToString(row["Value"])), json.Serialize(row));
                        if (candidate >= 3 && candidate <= 9)
                        {
                            decimal[] expected = { 200m, -1234m, 4294967296m, 4294967296m, 1.5m, 1.25m, 12.5m };
                            string displayedNumber = Convert.ToString(row["Value"]);
                            if (types[candidate] == "LongLong" || types[candidate] == "LongPtr") displayedNumber = displayedNumber.TrimEnd('^');
                            Assert.AreEqual(expected[candidate - 3], decimal.Parse(displayedNumber.Replace(',', '.'),
                                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture), json.Serialize(row));
                        }
                        if (candidate == 10)
                        {
                            DateTime value;
                            string displayed = Convert.ToString(row["Value"]).Trim('#');
                            bool parsed = DateTime.TryParse(displayed, System.Globalization.CultureInfo.GetCultureInfo("fr-FR"), System.Globalization.DateTimeStyles.None, out value) ||
                                DateTime.TryParse(displayed, System.Globalization.CultureInfo.GetCultureInfo("en-US"), System.Globalization.DateTimeStyles.None, out value);
                            Assert.IsTrue(parsed, json.Serialize(row));
                            Assert.AreEqual(new DateTime(2026, 11, 23), value.Date);
                        }
                    }
                }
                Assert.AreEqual(true, inspected["SelectionRestored"]);
                Assert.AreEqual(true, inspected["FocusRestored"]);
                var after = state();
                Assert.AreEqual(1, Convert.ToInt32(after["Mode"]));
                Assert.AreEqual(projectPath, after["SelectedProjectPath"]);
                Assert.AreEqual(module, after["ActiveModule"]);
                Assert.AreEqual(json.Serialize(before["Selection"]), json.Serialize(after["Selection"]));
                var unchanged = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                Assert.AreEqual(original["Sha256"], unchanged["Sha256"]);
                var phases = ReadTerminalPhases(tracePath, host.ProcessId, phaseStartedUtc);
                evidence.PhaseEvidence = phases;
                AssertScalarPagePhases(phases, Math.Max(0, Math.Min(offset + limit, eligible) - offset));
            }, () =>
            {
                // File reads cannot replay native work, including when the bridge is unavailable.
                if (inspectionPending || executionPending || !bridgeAvailable)
                    evidence.Shutdown = "Retained; native outcome uncertain; no Reset, Close or Quit was emitted";
                evidence.PhaseEvidence = ReadPhaseRows(tracePath, host.ProcessId, phaseStartedUtc);
                if (inspectionPending || executionPending || !bridgeAvailable)
                {
                    evidence.Shutdown = "Retained; native outcome uncertain; no Reset, Close or Quit was emitted";
                    Assert.Fail("Owned Excel retained for an uncertain native operation. PID: " + host.ProcessId + "; fixture: " + host.Root);
                }
                if (!canClose)
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
                Assert.IsTrue(canClose, "Retain owned Excel because design mode was not verified.");
                host.Dispose();
                evidence.Shutdown = "CloseQuitAndExitCodeZero";
            }, () => TestContext.AddResultFile(report));
        }

        private static IDictionary<string, object>[] ReadPhaseRows(string path, int processId, DateTime notBeforeUtc)
        {
            if (!File.Exists(path)) return new IDictionary<string, object>[0];
            string text;
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                Assert.IsTrue(file.Length <= VbeInspectionTrace.MaximumFileBytes, "Phase file exceeds its documented bound.");
                using (var reader = new StreamReader(file)) text = reader.ReadToEnd();
            }
            var lines = text.Split('\n');
            var json = new JavaScriptSerializer();
            // A still-appending final line is incomplete evidence until its newline is durable.
            return lines.Take(lines.Length - 1).Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => (IDictionary<string, object>)json.DeserializeObject(line))
                .Where(row => Convert.ToInt32(row["HostProcessId"]) == processId &&
                    DateTime.Parse((string)row["Utc"], System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.RoundtripKind).ToUniversalTime() >= notBeforeUtc).ToArray();
        }

        private static IDictionary<string, object>[] ReadTerminalPhases(string path, int processId, DateTime notBeforeUtc)
        {
            var deadline = Stopwatch.StartNew();
            IDictionary<string, object>[] rows;
            do
            {
                rows = ReadPhaseRows(path, processId, notBeforeUtc);
                if (rows.Any(row => (string)row["Phase"] == "Terminal")) return rows;
                Thread.Sleep(25);
            } while (deadline.ElapsedMilliseconds < 5000);
            Assert.Fail("No terminal phase evidence was observed for the owned host; an empty log proves no native behavior.");
            return rows;
        }

        private static void AssertScalarPagePhases(IDictionary<string, object>[] rows, int eligibleScalars)
        {
            Assert.IsTrue(rows.Length > 0);
            Assert.AreEqual(1, rows.Select(row => (string)row["Correlation"]).Distinct().Count(), "The owned host must have exactly one inspection correlation.");
            var phases = rows.Select(row => (string)row["Phase"]).ToArray();
            foreach (string required in new[] { "Enqueue", "CallbackEntered", "OwnerSta", "CoreEntered", "ContextValidated", "CoreTerminal", "Terminal" })
                Assert.IsTrue(phases.Contains(required), "Missing phase: " + required);
            foreach (string single in new[] { "Enqueue", "CallbackEntered", "OwnerSta", "CoreEntered", "CoreTerminal", "Terminal" })
                Assert.AreEqual(1, phases.Count(phase => phase == single), "Exactly one inspection request is required: " + single);
            Assert.IsTrue(rows.All(row => row["ErrorType"] == null), "Successful inspection must not conceal a native phase error.");
            var callback = rows.Single(row => (string)row["Phase"] == "CallbackEntered");
            foreach (var row in rows.Where(row => new[] { "OwnerSta", "ContextValidated", "Command229Before", "Command229Returned", "CoreTerminal" }.Contains((string)row["Phase"])))
            {
                Assert.AreEqual(callback["ThreadId"], row["ThreadId"]);
                Assert.AreEqual("STA", row["Apartment"]);
            }
            if (eligibleScalars == 0)
                Assert.IsFalse(phases.Any(phase => phase.StartsWith("Command229", StringComparison.Ordinal) || phase.StartsWith("Observer", StringComparison.Ordinal)), "Skipped declarations must never open Quick Watch.");
            else
                foreach (string required in new[] { "Command229Before", "Command229Returned", "ObserverEntered", "ObserverDialogFound", "ObserverReadComplete", "ObserverTerminal" })
                    Assert.AreEqual(eligibleScalars, phases.Count(phase => phase == required), "Exactly one native observer cycle per eligible scalar is required: " + required);
        }

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
            Func<object, IDictionary<string, object>> send = request =>
            {
                try
                {
                    var result = evidence.Send(request, () => VbeBridgeClient.Read("VBAi." + host.ProcessId, request, 20000), response =>
                    {
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
                send(new
                {
                    Command = "replace_lines",
                    Project = projectPath,
                    Module = module,
                    ExpectedSha256 = empty["Sha256"],
                    StartLine = 1,
                    Count = 0,
                    Text = source
                });
                var original = VbeBridgeClient.Object(send(new { Command = "read_module", Project = projectPath, Module = module })["Data"]);
                // Let native navigation settle before the single execution request.
                // A refused cold-start Run Sub remains a separate qualification gap.
                send(new
                {
                    Command = "select_code",
                    Project = projectPath,
                    Module = module,
                    StartLine = 2,
                    ExpectedSha256 = original["Sha256"]
                });
                var navigation = Stopwatch.StartNew();
                while (!string.Equals(Convert.ToString(state()["ActiveModule"]), module, StringComparison.Ordinal))
                {
                    if (navigation.ElapsedMilliseconds > 5000)
                        Assert.Fail("The native fixture code pane did not become active before execution.");
                    Thread.Sleep(50);
                }
                canClose = false;
                send(new
                {
                    Command = "run_sub",
                    Project = projectPath,
                    Module = module,
                    Procedure = "AuditLocals",
                    ExpectedMode = 2,
                    ExpectedSha256 = original["Sha256"]
                });
                WaitMode(state, 1);
                send(new { Command = "open_debug_pane", Action = "locals" });
                var inspected = VbeBridgeClient.Object(send(new
                {
                    Command = "inspect_local_scalars",
                    Project = projectPath,
                    Module = module,
                    Procedure = "AuditLocals",
                    ExpectedMode = 1,
                    ExpectedSha256 = original["Sha256"],
                    Limit = 8
                })["Data"]);
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
