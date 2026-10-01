using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies the installed editor UI inside owned Excel, with a retained inactive closed-project tab.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("NativeMonacoClosedProject"), DoNotParallelize]
    public sealed class ExcelNativeMonacoClosedProjectTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void InstalledEmbeddedEditorKeepsLiveStatusAfterOtherOwnedProjectCloses()
        {
            if (Environment.GetEnvironmentVariable("VBAi_NATIVE_MONACO_CLOSED_SCOPE_TEST") != "1")
                Assert.Inconclusive("Explicit VBAi_NATIVE_MONACO_CLOSED_SCOPE_TEST=1 is required; this test opens disposable native workbooks.");
            Guid expected;
            Assert.IsTrue(Guid.TryParse(Environment.GetEnvironmentVariable("VBAi_MONACO_EXPECTED_MVID"), out expected) && expected != Guid.Empty,
                "Set VBAi_MONACO_EXPECTED_MVID to the full frozen installed payload identity before any launch.");
            Assert.AreEqual(expected, typeof(VbeSession).Module.ModuleVersionId, "The isolated tests must reference the exact installed payload.");
            string output = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"));
            Directory.CreateDirectory(output);
            string trace = Path.Combine(output, "monaco-bootstrap-" + Guid.NewGuid().ToString("N") + ".jsonl");
            var host = ExcelVbeFixture.StartOwnedWithTrace(trace);
            string closedPath = host.File("ClosedScope.xlsm"), livePath = host.File("LiveScope.xlsm");
            const string closedModule = "ModuleClosedScope", liveModule = "ModuleLiveScope";
            const string closedCode = "Option Explicit\r\n' VBAi owned closed scope probe\r\nPublic Function ClosedValue() As Long\r\n    ClosedValue = 17\r\nEnd Function";
            const string liveCode = "Option Explicit\r\n' VBAi owned live scope probe\r\nPublic Function LiveValue() As Long\r\n    LiveValue = 42\r\nEnd Function";
            var records = new List<object>(); var json = new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024 };
            var containment = new OfficeCommandContainment();
            string report = host.File("native-monaco-closed-project.json"), capture = host.File("native-monaco-live.png");
            string result = "RUNNING"; Exception failure = null;
            Action persist = () => File.WriteAllText(report, json.Serialize(new { Result = result, HostProcessId = host.ProcessId,
                ExpectedMvid = expected.ToString("D"), Source = "Actual installed embedded editor; detached forms do not qualify",
                ClosedPath = closedPath, LivePath = livePath, CommandPending = containment.Pending, DeliveryUncertain = containment.Uncertain,
                Records = records, Shutdown = host.ShutdownDiagnostics }), new UTF8Encoding(false));
            Action<string, object> phase = (name, value) => { records.Add(new { Phase = name, Utc = DateTime.UtcNow.ToString("O"), Data = value }); persist(); };
            Func<string, object, IDictionary<string, object>> send = (command, request) => {
                var response = containment.Send(command, request, records.Add, persist,
                    () => VbeBridgeClient.Read("VBAi." + host.ProcessId, request, 20000), host.PreserveMonacoNativeOutcome);
                Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                return response;
            };
            Func<string, string, IDictionary<string, object>> read = (path, module) => VbeBridgeClient.Object(send("read_module",
                new { Command = "read_module", Project = path, Module = module })["Data"]);
            Func<string, IDictionary<string, object>> state = path => VbeBridgeClient.Object(send("debug_state", new { Command = "debug_state", Project = path })["Data"]);
            Action<string, string, object> select = (path, module, sha) => {
                send("select_code", new { Command = "select_code", Project = path, Module = module, StartLine = 1,
                    StartColumn = 1, ExpectedSha256 = sha, ExpectedMode = 2 });
                var actual = state(path);
                RequireSelectedScope(actual, path, module); phase("NativeSelection", actual);
            };
            IDictionary<string, object> closedSource = null, liveSource = null;
            try
            {
                persist(); phase("CreateOwnedBooksIntent", new { closedPath, livePath });
                containment.RequireTerminal(); host.PrepareMonacoBooks(closedPath, livePath);
                var status = VbeBridgeClient.Object(send("status", new { Command = "status" })["Data"]);
                ExcelVbeFixture.RequireMonacoCandidate(expected, typeof(VbeSession).Module.ModuleVersionId, host.ProcessId, status);
                var projects = ((object[])send("list_projects", new { Command = "list_projects" })["Data"]).Select(VbeBridgeClient.Object).ToArray();
                foreach (string path in new[] { closedPath, livePath })
                {
                    var project = projects.Single(item => string.Equals(Convert.ToString(item["FileName"]), path, StringComparison.OrdinalIgnoreCase));
                    Assert.AreEqual(2, Convert.ToInt32(project["Mode"]));
                    string module = path == closedPath ? closedModule : liveModule;
                    send("create_module", new { Command = "create_module", Project = path, Module = module, ExpectedMode = 2 });
                    var empty = read(path, module);
                    send("replace_lines", new { Command = "replace_lines", Project = path, Module = module, StartLine = 1, Count = 0,
                        Text = path == closedPath ? closedCode : liveCode, ExpectedSha256 = empty["Sha256"], ExpectedMode = 2 });
                }
                closedSource = read(closedPath, closedModule); liveSource = read(livePath, liveModule);
                phase("ExactSourceBeforeClose", new { Closed = closedSource, Live = liveSource, Projects = projects });
                File.WriteAllText(host.File("closed-source-before-close.vba.txt"), Convert.ToString(closedSource["Code"]), new UTF8Encoding(false));
                File.WriteAllText(host.File("live-source-before-close.vba.txt"), Convert.ToString(liveSource["Code"]), new UTF8Encoding(false));
                containment.RequireTerminal(); host.SaveMonacoBaselines(closedPath, livePath);
                string closedDiskHash = ReadNativeBaselineHash(closedPath);
                phase("PreservedNativeBaselines", new { Closed = new { Path = closedPath, Sha256 = closedDiskHash },
                    Live = new { Path = livePath, Sha256 = ReadNativeBaselineHash(livePath) },
                    ClosedSourceSnapshot = host.File("closed-source-before-close.vba.txt"), LiveSourceSnapshot = host.File("live-source-before-close.vba.txt") });
                CultureInfo culture = host.ReadMonacoHostCulture();
                string synchronized = ExcelVbeFixture.MonacoLocalized(ExcelVbeFixture.MonacoSynchronizedStatus, culture);
                phase("HostLocalization", new { Culture = culture.Name, Synchronized = synchronized,
                    ClosedWarning = ExcelVbeFixture.MonacoLocalized(ExcelVbeFixture.MonacoClosedStatus, culture) });
                select(closedPath, closedModule, closedSource["Sha256"]);
                WaitForTab(host, culture, closedModule, phase);
                select(livePath, liveModule, liveSource["Sha256"]);
                WaitForTab(host, culture, liveModule, phase);
                // Close only the old registered owned workbook once, after exact source snapshots and native baselines exist.
                phase("OldOwnedCloseWithoutSaveIntent", new { Path = closedPath, Code = closedSource, Save = false, Attempt = 1 });
                containment.RequireTerminal(); host.CloseMonacoOldScopeWithoutSave(closedPath);
                phase("OldOwnedCloseReturned", new { Path = closedPath, Save = false });
                Assert.AreEqual(closedDiskHash, ReadNativeBaselineHash(closedPath), "Close(false) changed the retained old native baseline.");
                select(livePath, liveModule, liveSource["Sha256"]);
                var selectedBefore = state(livePath); RequireSelectedScope(selectedBefore, livePath, liveModule);
                var clock = Stopwatch.StartNew(); long? stableSince = null; string lastFailure = null;
                ExcelVbeFixture.MonacoNativeObservation actualUi = null;
                do
                {
                    phase("ActualHostUiObservationStarting", new { ElapsedMilliseconds = clock.ElapsedMilliseconds });
                    actualUi = host.ReadInstalledMonaco(culture); phase("ActualHostUiObservation", actualUi);
                    try
                    {
                        ExcelVbeFixture.RequireInstalledMonacoObservation(actualUi, host.ProcessId, closedModule, liveModule, synchronized);
                        if (!stableSince.HasValue) stableSince = clock.ElapsedMilliseconds;
                        if (clock.ElapsedMilliseconds - stableSince.Value >= 3000) break; // More than three reconciliation timer periods.
                    }
                    catch (AssertFailedException error) { stableSince = null; lastFailure = error.Message; }
                    Thread.Sleep(250); // Read-only settlement; select/close/source mutations are never replayed.
                } while (clock.ElapsedMilliseconds < 20000);
                phase("CaptureActualInstalledWindowIntent", new { Capture = capture });
                host.CaptureInstalledMonaco(actualUi, capture);
                Assert.IsTrue(stableSince.HasValue && clock.ElapsedMilliseconds - stableSince.Value >= 3000,
                    "Installed embedded Monaco never settled with visible live code and a scoped healthy status: " + lastFailure);
                var afterSource = read(livePath, liveModule); var afterState = state(livePath);
                RequireSelectedScope(afterState, livePath, liveModule);
                Assert.AreEqual(liveSource["Sha256"], afterSource["Sha256"]); Assert.AreEqual(liveSource["Code"], afterSource["Code"]);
                Assert.AreEqual(json.Serialize(selectedBefore["Selection"]), json.Serialize(afterState["Selection"]));
                var remaining = ((object[])send("list_projects", new { Command = "list_projects" })["Data"]).Select(VbeBridgeClient.Object).ToArray();
                Assert.IsFalse(remaining.Any(project => string.Equals(Convert.ToString(project["FileName"]), closedPath, StringComparison.OrdinalIgnoreCase)));
                Assert.AreEqual(1, remaining.Count(project => string.Equals(Convert.ToString(project["FileName"]), livePath, StringComparison.OrdinalIgnoreCase)));
                phase("LiveScopeUnchanged", new { Source = afterSource, State = afterState, Projects = remaining });
                result = "ASSERTIONS_PASS_CLEANUP_PENDING";
            }
            catch (Exception error) { failure = error; result = "FAILED"; phase("OriginalFailure", error.ToString()); }
            try
            {
                containment.RequireTerminal();
                Assert.IsFalse(host.PreserveForDiagnosticRecovery, "Native close or bridge outcome uncertain; retain exact owned host without cleanup.");
                phase("OwnedCleanupIntent", new { ForcedTermination = false });
                host.CloseMonacoLiveScopeWithoutSave(livePath); host.Dispose();
                Assert.AreEqual(true, host.ShutdownDiagnostics["Exited"]); Assert.AreEqual("0x00000000", host.ShutdownDiagnostics["ExitCodeHex"]);
                phase("OwnedNormalExitVerified", host.ShutdownDiagnostics);
                if (failure == null) result = "PASS";
            }
            catch (Exception cleanup)
            {
                host.PreserveMonacoNativeOutcome(); result = "FAILED"; phase("CleanupFailureNoReplay", cleanup.ToString());
                failure = failure == null ? cleanup : new AggregateException("Native Monaco scenario and cleanup both failed; original evidence retained.", failure, cleanup);
            }
            finally
            {
                persist(); TestContext.AddResultFile(report);
                if (File.Exists(capture)) TestContext.AddResultFile(capture);
                if (File.Exists(capture + ".json")) TestContext.AddResultFile(capture + ".json");
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private static void RequireSelectedScope(IDictionary<string, object> state, string path, string module)
        {
            Assert.AreEqual(2, Convert.ToInt32(state["Mode"])); Assert.AreEqual(path, Convert.ToString(state["SelectedProjectPath"]), true);
            Assert.AreEqual(module, Convert.ToString(state["ActiveModule"]));
        }

        private static string ReadNativeBaselineHash(string path)
        {
            using (var bytes = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sha = System.Security.Cryptography.SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }

        private static void WaitForTab(ExcelVbeFixture host, CultureInfo culture, string module, Action<string, object> phase)
        {
            var clock = Stopwatch.StartNew();
            do
            {
                var ui = host.ReadInstalledMonaco(culture); phase("RealEditorTabCreationObservation", ui);
                if (ui.EditorProcessId == host.ProcessId && ui.Tabs.Contains(module) && ui.SelectedTabs.Contains(module)) return;
                Thread.Sleep(250);
            } while (clock.ElapsedMilliseconds < 15000);
            Assert.Fail("Native select_code did not create/select the real installed Monaco tab " + module + "; no detached window or command replay is substituted.");
        }
    }
}
