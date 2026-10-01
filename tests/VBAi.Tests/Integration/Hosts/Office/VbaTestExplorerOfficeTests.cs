using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies the registered add-in in owned, disposable Office processes without changing trust or registration.</summary>
    [TestClass, TestCategory("VbaTestOffice"), DoNotParallelize]
    public sealed class VbaTestExplorerOfficeTests
    {
        private const string ModuleName = "VBAiOfficeModule";
        private const string ProductionModuleName = "VBAiOfficeProduction";
        private const string RuntimeClsid = "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out WindowRect rectangle);
        [StructLayout(LayoutKind.Sequential)] private struct WindowRect { public int Left, Top, Right, Bottom; }

        /// <summary>Uses only explicit discovered IDs and the installed native callback, never a substitute execution adapter.</summary>
        [STATestMethod]
        [DataRow("Word")]
        [DataRow("PowerPoint")]
        [DataRow("Access")]
        [DataRow("Publisher")]
        public void RegisteredOfficeRunsBatchAndSingleWithVerifiedResults(string host)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Native Office tests require VBAi_RUN_OFFICE_TESTS=1.");
            // Capture provenance before COM activation; an unavailable source identity must not become a fabricated result.
            string sourceRevision = Git("rev-parse HEAD").Trim();
            string sourceStatus = Git("status --porcelain");
            using (var fixture = OfficeVbeFixture.Start(host, host == "Access" ? "Access.Application.16" : null,
                allowExistingHost: true, allowForcedTermination: false))
            {
                ExecuteRegisteredOfficeQualification(fixture, host, sourceRevision, sourceStatus);
                // The successful non-inlined scope has returned before the optional testhost-only collection.
                fixture.CollectSettledWordScopeDiagnostic();
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ExecuteRegisteredOfficeQualification(OfficeVbeFixture fixture, string host, string sourceRevision, string sourceStatus)
        {
            try
            {
            Console.WriteLine("Registered Office qualification=" + fixture.Root);
            var status = fixture.Data("status");
            Save(fixture, "identity.json", new { CurrentSourceRevision = sourceRevision, SourceStatus = sourceStatus,
                Host = host, fixture.ProcessId, fixture.DocumentPath, fixture.Project, LoadedAssembly = status,
                ExecutionBoundary = "Registered in-process add-in with registered VBAi.TestRuntime callback" });
            Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
            Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(status["HostProcessId"]));
            Save(fixture, "callback-registration.json", RequireCallback((string)status["AssemblyPath"]));

            var publisherSources = host == "Publisher" ? fixture.Items("list_modules").ToDictionary(
                item => (string)item["Name"], item => (string)fixture.Data("read_module", "Module", item["Name"])["Code"],
                StringComparer.OrdinalIgnoreCase) : null;
            fixture.Data("create_module", "Module", ModuleName, "ExpectedMode", 2);
            bool documentCoverage = host == "Word" || host == "PowerPoint";
            if (documentCoverage)
            {
                fixture.Data("create_module", "Module", ProductionModuleName, "ExpectedMode", 2);
                ReplaceSource(fixture, ProductionSource, ProductionModuleName);
            }
            ReplaceSource(fixture, documentCoverage ? DocumentCoverageSource : Source);
            var original = fixture.Data("read_module", "Module", ModuleName);
            File.WriteAllText(Path.Combine(fixture.Root, "synthetic-tests.bas"), (string)original["Code"], Utf8);
            var before = fixture.Data("discover_vba_tests");
            Save(fixture, "discovery-before.json", before);
            AssertCatalogue(before);

            var preview = fixture.Data("preview_vba_test_support");
            string support = (string)preview["Text"];
            // Review the complete generated source against the source generator in this exact candidate build.
            // Installation still requires both the exact preview text and its exact project revision.
            Assert.AreEqual(VbaTestRuntimeSource.Generate(ReviewCatalogue(before)), support,
                "The installed candidate did not generate the exact reviewed support source.");
            StringAssert.Contains(support, "CreateObject(\"VBAi.TestRuntime\")");
            StringAssert.Contains(support, "Public Sub VBAiExecutePendingTest()");
            Assert.IsFalse(support.Contains("Shell("), "Unexpected generated execution outside VBA.");
            File.WriteAllText(Path.Combine(fixture.Root, "reviewed-support.bas"), support, Utf8);
            Save(fixture, "support-preview.json", preview);
            var installed = fixture.Data("install_vba_test_support", "ExpectedProjectVersion", preview["ExpectedProjectVersion"],
                "ExpectedMode", 2, "Text", support);
            Save(fixture, "support-install.json", installed);
            Assert.AreEqual(true, installed["Applied"]);
            fixture.SaveNative();
            var catalog = fixture.Data("discover_vba_tests");
            Save(fixture, "discovery-ready.json", catalog);
            var tests = AssertCatalogue(catalog);
            Assert.IsTrue(string.IsNullOrEmpty(catalog["ExecutionUnavailableReason"] as string),
                "Registered execution is unavailable: " + catalog["ExecutionUnavailableReason"]);
            string revision = (string)catalog["ExpectedProjectVersion"];
            string[] ids = tests.OrderBy(test => (string)test["Procedure"], StringComparer.Ordinal).Select(test => (string)test["Id"]).ToArray();
            Assert.AreEqual(4, ids.Distinct(StringComparer.Ordinal).Count());

            var batch = Run(fixture, revision, ids, "batch");
            AssertReport(batch, revision, new[] { "ABooleanPass", "BBooleanFail", "CSwallowedAssertion", "DRuntimeError" });
            AssertSource(fixture, original);
            var passed = tests.Single(test => (string)test["Procedure"] == "ABooleanPass");
            var single = Run(fixture, revision, new[] { (string)passed["Id"] }, "single");
            AssertReport(single, revision, new[] { "ABooleanPass" });
            AssertSource(fixture, original);
            if (documentCoverage) QualifyOfficeDocumentCoverage(fixture, catalog, revision, passed);

            var shown = fixture.Data("show_vba_test_explorer");
            Save(fixture, "explorer.json", shown);
            Assert.AreEqual(true, shown["Opened"]);
            Assert.AreEqual(true, shown["Docked"]);
            var window = new IntPtr(Convert.ToInt64(shown["Hwnd"]));
            Assert.IsTrue(IsWindow(window), "Explorer handle is not a live window.");
            uint owner; Assert.AreNotEqual(0u, GetWindowThreadProcessId(window, out owner));
            Assert.AreEqual((uint)fixture.ProcessId, owner, "Explorer belongs to a different process.");
            WindowRect bounds; Assert.IsTrue(GetWindowRect(window, out bounds));
            Assert.IsTrue(bounds.Right > bounds.Left && bounds.Bottom > bounds.Top, "Explorer has no visible-sized bounds.");
            Save(fixture, "explorer-native-window.json", new { Hwnd = window.ToInt64(), ProcessId = owner,
                bounds.Left, bounds.Top, bounds.Right, bounds.Bottom });

            if (host == "Access")
            {
                // The final stale-revision edit must remain unsaved; discard owned Access objects on cleanup.
                fixture.StopAccessSaveDialogHandler();
                fixture.RequireAdapterOnlyCleanup();
            }
            ReplaceSource(fixture, ((string)original["Code"]).TrimEnd() + "\r\n' changed after verified native runs\r\n");
            var changed = fixture.Data("read_module", "Module", ModuleName);
            var refused = fixture.Response("run_vba_tests", "ExpectedProjectVersion", revision, "ExpectedMode", 2,
                "Items", new[] { (string)passed["Id"] });
            Save(fixture, "stale-run-refusal.json", refused);
            Assert.AreEqual(false, refused["Ok"], "A stale source revision was executed.");
            StringAssert.Contains(Convert.ToString(refused["Error"]), "ExpectedProjectVersion");
            AssertSource(fixture, changed);
            var historical = fixture.Data("vba_test_run_status", "Query", single["run"], "Action", "compact");
            Save(fixture, "historical-stale-report.json", historical);
            Assert.AreEqual(true, historical["Stale"]);
            Assert.AreEqual(false, historical["Pending"]);
            Assert.AreEqual("Completed", historical["State"]);
            // Publisher requires the disposable VBA project to be saved before its guarded Quit.
            if (host == "Publisher")
            {
                publisherSources.Add(ModuleName, (string)changed["Code"]);
                publisherSources.Add(VbaTestRuntimeSource.ModuleName, support);
                fixture.SaveReviewedPublisherProject(ModuleName, publisherSources);
            }
            }
            catch (Exception error)
            {
                Save(fixture, "qualification-body-error.json", new { Exception = error.ToString(), error.HResult, fixture.NativeExecutionUnsettled });
                throw;
            }
        }

        private static IDictionary<string, object>[] AssertCatalogue(IDictionary<string, object> catalog)
        {
            var modules = ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object).ToArray();
            var module = modules.Single(item => (string)item["Name"] == ModuleName);
            Assert.IsTrue(string.IsNullOrEmpty(module["Diagnostic"] as string), Convert.ToString(module["Diagnostic"]));
            foreach (string fixture in new[] { "ModuleInitialize", "ModuleCleanup", "TestInitialize", "TestCleanup" })
                Assert.IsNotNull(module[fixture], "Missing discovered fixture: " + fixture);
            var tests = ((object[])module["Tests"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(4, tests.Length);
            Assert.AreEqual(4, modules.Sum(item => ((object[])item["Tests"]).Length), "Unexpected tests outside the synthetic module.");
            foreach (var test in tests) Assert.IsTrue(string.IsNullOrEmpty(test["Diagnostic"] as string), Convert.ToString(test["Diagnostic"]));
            return tests;
        }

        private static IDictionary<string, object> Run(OfficeVbeFixture fixture, string revision, string[] ids, string name, bool coverage = false)
        {
            // Even a missing start response can conceal an already dispatched native attempt.
            fixture.NativeExecutionUnsettled = true;
            // Original and copy can share VBProject.Name while a coverage run is pending.
            string selector = coverage && fixture.Kind == "Word" ? fixture.DocumentPath : fixture.Project;
            var started = coverage
                ? fixture.Data("run_vba_tests", "Project", selector, "ExpectedProjectVersion", revision, "ExpectedMode", 2, "Items", ids, "Action", "coverage")
                : fixture.Data("run_vba_tests", "Project", selector, "ExpectedProjectVersion", revision, "ExpectedMode", 2, "Items", ids);
            Save(fixture, name + "-start.json", started);
            string query = (string)started["Query"];
            Assert.IsFalse(string.IsNullOrEmpty(query));
            var deadline = DateTime.UtcNow.AddSeconds(60);
            IDictionary<string, object> state;
            do
            {
                state = fixture.Data("vba_test_run_status", "Project", selector, "Query", query, "Action", "compact");
                object pending, reportValue, uncertain;
                if (state.TryGetValue("Pending", out pending) && Equals(pending, false)
                    && state.TryGetValue("Report", out reportValue) && reportValue != null)
                {
                    var observed = VbeBridgeClient.Object(reportValue);
                    if (observed.TryGetValue("uncertain", out uncertain) && Equals(uncertain, false))
                        fixture.NativeExecutionUnsettled = false;
                }
                if (!Convert.ToBoolean(state["Pending"])) break;
                Thread.Sleep(100);
            } while (DateTime.UtcNow < deadline);
            Save(fixture, name + "-status.json", state);
            Assert.AreEqual(false, state["Pending"], "Native run observation timed out. No retry or reset was attempted.");
            Assert.AreEqual("Completed", state["State"]);
            Assert.AreEqual(false, state["Stale"]);
            var report = VbeBridgeClient.Object(state["Report"]);
            Save(fixture, name + "-compact.json", report);
            var human = fixture.Data("vba_test_run_status", "Project", selector, "Query", query, "Action", "human");
            File.WriteAllText(Path.Combine(fixture.Root, name + "-human.txt"), (string)human["Report"], Utf8);
            Assert.AreEqual(query, report["run"]);
            return report;
        }

        private static void QualifyOfficeDocumentCoverage(OfficeVbeFixture fixture, IDictionary<string, object> catalog,
            string revision, IDictionary<string, object> selected)
        {
            int? originalWordSecurity = fixture.Kind == "Word" ? SetOwnedWordCoveragePolicy(fixture, 2, "prepare") : (int?)null;
            try { QualifyOfficeDocumentCoverageCore(fixture, catalog, revision, selected); }
            finally
            {
                // Do not mutate a retained host after an uncertain native outcome. Owned cleanup remains a separate action.
                if (originalWordSecurity.HasValue && !fixture.NativeExecutionUnsettled)
                    SetOwnedWordCoveragePolicy(fixture, originalWordSecurity.Value, "restore");
            }
        }

        private static void QualifyOfficeDocumentCoverageCore(OfficeVbeFixture fixture, IDictionary<string, object> catalog,
            string revision, IDictionary<string, object> selected)
        {
            // Exercise the installed service and real host copy adapter; the fixture does not inject a clone.
            if (fixture.Kind == "PowerPoint") ConfigureOwnedPowerPointCoveragePolicy(fixture);
            var preview = fixture.Data("vba_test_coverage");
            Save(fixture, "coverage-preview.json", preview);
            Assert.AreEqual(false, preview["Available"], "A preview must not claim measured coverage.");
            Assert.AreEqual(true, preview["Supported"], Convert.ToString(preview["ExecutionUnavailableReason"]));
            Assert.AreEqual(true, preview["DenominatorKnown"]);
            Assert.AreEqual(2, Convert.ToInt32(preview["EligibleProcedureCount"]));
            var probes = ((object[])preview["Probes"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(2, probes.Length);
            CollectionAssert.AreEquivalent(new[] { "EnteredFunction", "UnenteredFunction" },
                probes.Select(probe => (string)probe["Procedure"]).ToArray());
            foreach (var probe in probes)
            {
                Assert.AreEqual(ProductionModuleName, probe["Module"]);
                Assert.AreEqual("Function", probe["Kind"]);
                Assert.IsTrue(Convert.ToInt32(probe["OriginalLine"]) > 0);
            }

            var originalModules = fixture.Items("list_modules").Select(module => (string)module["Name"])
                .ToDictionary(name => name, name => fixture.Data("read_module", "Module", name), StringComparer.Ordinal);
            byte[] originalDisk = ReadOwnedDocumentBytes(fixture.DocumentPath);
            string[] openBefore;
            string countersBefore = ReadOfficeCounters(fixture, out openBefore);
            Assert.AreEqual("2,0", countersBefore, "Only the entered procedure should have run in the original batch and single test.");
            CollectionAssert.AreEquivalent(new[] { fixture.DocumentPath }, openBefore);
            string coverageRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "CoverageRuns");
            var foldersBefore = new HashSet<string>(Directory.Exists(coverageRoot) ? Directory.GetDirectories(coverageRoot) : new string[0], StringComparer.OrdinalIgnoreCase);

            string testId = (string)selected["Id"];
            var report = Run(fixture, revision, new[] { testId }, "coverage", true);
            AssertReport(report, revision, new[] { "ABooleanPass" });
            var results = ((object[])report["tests"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(testId, results.Single()["id"], "A clone-specific test ID escaped into the original run.");
            var compact = VbeBridgeClient.Object(report["coverage"]);
            AssertMeasuredCoverage(compact, "available", "complete", "revision", "metric", "eligible", "hit", "percent", revision);
            Assert.AreEqual(preview["Project"], compact["original"]);
            Assert.AreEqual(false, compact["statementCoverageAvailable"]);
            Assert.AreEqual(false, compact["branchCoverageAvailable"]);
            var hits = ((object[])compact["probes"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(2, hits.Length);
            foreach (var hit in hits)
            {
                var probe = probes.Single(item => Equals(item["Id"], hit["id"]));
                Assert.AreEqual(probe["Module"], hit["module"]);
                Assert.AreEqual(probe["Procedure"], hit["procedure"]);
                Assert.AreEqual(probe["OriginalLine"], hit["line"]);
                Assert.AreEqual(probe["OriginalColumn"], hit["column"]);
                Assert.AreEqual(Equals(probe["Procedure"], "EnteredFunction"), hit["entered"]);
            }
            var measured = fixture.Data("vba_test_coverage", "Query", report["run"]);
            Save(fixture, "coverage-measured.json", measured);
            Assert.AreEqual("Completed", measured["State"]);
            Assert.AreEqual(false, measured["Stale"]);
            Assert.AreEqual(false, measured["Uncertain"]);
            Assert.AreEqual(true, measured["Available"]);
            AssertMeasuredCoverage(VbeBridgeClient.Object(measured["Report"]),
                "Available", "Complete", "Revision", "Metric", "Eligible", "Hit", "Percent", revision);
            string human = File.ReadAllText(Path.Combine(fixture.Root, "coverage-human.txt"));
            StringAssert.Contains(human, "50%");
            StringAssert.Contains(human, "(1/2)");

            var refreshed = fixture.Data("discover_vba_tests");
            Assert.AreEqual(revision, refreshed["ExpectedProjectVersion"], "Coverage changed the original source/reference revision.");
            CollectionAssert.AreEquivalent(AssertCatalogue(catalog).Select(test => (string)test["Id"]).ToArray(),
                AssertCatalogue(refreshed).Select(test => (string)test["Id"]).ToArray(), "Coverage changed the original discovered test IDs.");
            CollectionAssert.AreEquivalent(originalModules.Keys.ToArray(),
                fixture.Items("list_modules").Select(module => (string)module["Name"]).ToArray(),
                "Instrumentation leaked into the original project.");
            foreach (var original in originalModules) AssertSource(fixture, original.Value, original.Key);
            CollectionAssert.AreEqual(originalDisk, ReadOwnedDocumentBytes(fixture.DocumentPath), "Coverage saved or changed the original document file.");
            string[] openAfter;
            string countersAfter = ReadOfficeCounters(fixture, out openAfter);
            Assert.AreEqual(countersBefore, countersAfter, "The coverage selection executed in the original project.");
            CollectionAssert.AreEquivalent(openBefore, openAfter, "The measured copy remained open after a completed run.");
            var openProjects = fixture.Items("list_projects");
            string copyName = fixture.Kind == "PowerPoint" ? "coverage.pptm" : "coverage" + Path.GetExtension(fixture.DocumentPath);
            Assert.IsFalse(openProjects.Any(project => string.Equals(Path.GetFileName(project["FileName"] as string), copyName, StringComparison.OrdinalIgnoreCase)),
                "A coverage copy project remained in the live VBE.");

            var retained = Directory.GetDirectories(coverageRoot).Where(folder => !foldersBefore.Contains(folder)).Where(folder => {
                string planPath = Path.Combine(folder, "coverage-plan.json");
                if (!File.Exists(planPath)) return false;
                var plan = VbeBridgeClient.Object(new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 }.DeserializeObject(File.ReadAllText(planPath)));
                return Equals(plan["Original"], preview["Project"]) && Equals(plan["Revision"], revision);
            }).ToArray();
            Assert.AreEqual(1, retained.Length, "No unique retained coverage plan belongs to this original revision.");
            Assert.IsTrue(File.Exists(Path.Combine(retained[0], copyName)), "The owned measurement copy was not retained for inspection.");
            Save(fixture, "coverage-original-preservation.json", new { OriginalRevision = revision, TestId = testId,
                CountersBefore = countersBefore, CountersAfter = countersAfter, OpenBefore = openBefore, OpenAfter = openAfter,
                CopyClosed = true, RetainedFolder = retained[0], ExpectedEligible = 2, ExpectedHit = 1, ExpectedPercent = 50 });
        }

        private static void ConfigureOwnedPowerPointCoveragePolicy(OfficeVbeFixture fixture)
        {
            // The fixture's initial ForceDisable applies to files opened later. ByUI respects the existing Trust Center policy.
            // Change only this verified owned application; do not enable macros unconditionally or touch retained runs.
            Assert.IsFalse(fixture.NativeExecutionUnsettled, "Coverage policy setup requires a confirmed settled owned fixture.");
            object application = Marshal.GetActiveObject("PowerPoint.Application");
            try
            {
                IntPtr window = PowerPointWindow.Read(application);
                Assert.AreNotEqual(IntPtr.Zero, window);
                uint owner; Assert.AreNotEqual(0u, GetWindowThreadProcessId(window, out owner));
                Assert.AreEqual((uint)fixture.ProcessId, owner, "Coverage policy setup resolved another PowerPoint process.");
                int before = Convert.ToInt32(((dynamic)application).AutomationSecurity);
                fixture.NativeExecutionUnsettled = true;
                ((dynamic)application).AutomationSecurity = 2; // msoAutomationSecurityByUI; never msoAutomationSecurityLow.
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)application).AutomationSecurity));
                fixture.NativeExecutionUnsettled = false;
                Save(fixture, "powerpoint-coverage-security.json", new { ProcessId = owner, Hwnd = window.ToInt64(),
                    AutomationSecurityBefore = before, AutomationSecurityAfter = 2, Mode = "ByUI", TrustSettingsChanged = false });
            }
            finally { if (Marshal.IsComObject(application)) Marshal.ReleaseComObject(application); }
        }

        private static int SetOwnedWordCoveragePolicy(OfficeVbeFixture fixture, int security, string phase)
        {
            // ByUI respects the existing Trust Center policy; only the disposable owned application is changed.
            Assert.IsFalse(fixture.NativeExecutionUnsettled, "Coverage policy setup requires a confirmed settled owned fixture.");
            object application = Marshal.GetActiveObject("Word.Application");
            try
            {
                IntPtr window = VbaTestWordValuesHost.ReadApplicationWindow(application);
                uint owner; Assert.AreNotEqual(0u, GetWindowThreadProcessId(window, out owner));
                Assert.AreEqual((uint)fixture.ProcessId, owner, "Coverage policy setup resolved another Word process.");
                int before = Convert.ToInt32(((dynamic)application).AutomationSecurity);
                fixture.NativeExecutionUnsettled = true;
                ((dynamic)application).AutomationSecurity = security;
                Assert.AreEqual(security, Convert.ToInt32(((dynamic)application).AutomationSecurity));
                fixture.NativeExecutionUnsettled = false;
                Save(fixture, "word-coverage-security-" + phase + ".json", new { ProcessId = owner, Hwnd = window.ToInt64(),
                    AutomationSecurityBefore = before, AutomationSecurityAfter = security, TrustSettingsChanged = false });
                return before;
            }
            finally { if (Marshal.IsComObject(application)) Marshal.ReleaseComObject(application); }
        }

        private static void AssertMeasuredCoverage(IDictionary<string, object> coverage, string available, string complete,
            string revisionKey, string metric, string eligible, string hit, string percent, string revision)
        {
            Assert.AreEqual(true, coverage[available]); Assert.AreEqual(true, coverage[complete]);
            Assert.AreEqual(revision, coverage[revisionKey]); Assert.AreEqual("Procedure", coverage[metric]);
            Assert.AreEqual(2, Convert.ToInt32(coverage[eligible])); Assert.AreEqual(1, Convert.ToInt32(coverage[hit]));
            Assert.IsNotNull(coverage[percent]); Assert.AreEqual(50d, Convert.ToDouble(coverage[percent]));
        }

        private static byte[] ReadOwnedDocumentBytes(string path)
        {
            // Word keeps the owned saved document open for writing even while Saved is true.
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var bytes = new MemoryStream())
            { input.CopyTo(bytes); return bytes.ToArray(); }
        }
        private static readonly List<object[]> retainedCounterContexts = new List<object[]>();
        private static string ReadOfficeCounters(OfficeVbeFixture fixture, out string[] openPaths)
        {
            // A mismatched registered application is refused before inspecting or activating any document.
            // Read only the synthetic getter, with the host's actual invocation shape and exact owned context.
            object application = Marshal.GetActiveObject(fixture.Kind + ".Application");
            object documents = null, sourceDocument = null, sourceProject = null;
            try
            {
                // PowerPoint HWND is a restricted vtable member; the shared reader uses the published PIA layout.
                IntPtr handle = fixture.Kind == "PowerPoint" ? PowerPointWindow.Read(application)
                    : VbaTestWordValuesHost.ReadApplicationWindow(application);
                Assert.AreNotEqual(IntPtr.Zero, handle, "Counter inspection requires a verifiable owned application window.");
                uint owner; Assert.AreNotEqual(0u, GetWindowThreadProcessId(handle, out owner));
                Assert.AreEqual((uint)fixture.ProcessId, owner, "Counter inspection resolved another " + fixture.Kind + " process.");
                documents = fixture.Kind == "PowerPoint" ? ((dynamic)application).Presentations : ((dynamic)application).Documents;
                var paths = new List<string>();
                int documentCount = Convert.ToInt32(((dynamic)documents).Count);
                Assert.IsTrue(documentCount > 0 && documentCount <= 1000, "The owned document inventory must be bounded.");
                for (int documentIndex = 1; documentIndex <= documentCount; documentIndex++)
                {
                    object document = ((dynamic)documents)[documentIndex];
                    bool retained = false;
                    try
                    {
                        string path = Path.GetFullPath((string)((dynamic)document).FullName);
                        paths.Add(path);
                        if (string.Equals(path, fixture.DocumentPath, StringComparison.OrdinalIgnoreCase))
                        {
                            Assert.IsNull(sourceDocument, "Multiple open documents have the owned path.");
                            sourceDocument = document; retained = true;
                        }
                    }
                    finally { if (!retained && Marshal.IsComObject(document)) Marshal.ReleaseComObject(document); }
                }
                openPaths = paths.ToArray();
                Assert.AreEqual(1, paths.Count(path => string.Equals(path, fixture.DocumentPath, StringComparison.OrdinalIgnoreCase)));
                string fileName = Path.GetFileName(fixture.DocumentPath);
                Assert.AreEqual(1, paths.Count(path => string.Equals(Path.GetFileName(path), fileName, StringComparison.OrdinalIgnoreCase)), "The getter's document filename is ambiguous.");
                if (fixture.Kind == "PowerPoint")
                {
                    Assert.IsNotNull(sourceDocument);
                    var transport = new VbaTestPowerPointValuesHost {
                        ReadProcessName = () => "POWERPNT", ReadProcessId = () => fixture.ProcessId,
                        ReadActiveApplication = unused => application };
                    sourceProject = ((dynamic)sourceDocument).VBProject;
                    var target = transport.ResolveTarget(sourceProject, fixture.DocumentPath);
                    return Convert.ToString(transport.Invoke(target, ModuleName, "ReadCoverageCounters", new object[0]));
                }
                Assert.AreEqual("Word", fixture.Kind, "Counter inspection is not implemented for this host.");
                Assert.IsNotNull(sourceDocument);
                var wordTransport = new VbaTestWordValuesHost {
                    ReadProcessName = () => "WINWORD", ReadProcessId = () => fixture.ProcessId,
                    ReadActiveApplication = unused => application };
                sourceProject = ((dynamic)sourceDocument).VBProject;
                using (var wordTarget = (VbaTestWordValuesHost.OwnedTarget)wordTransport.ResolveTarget(sourceProject, fixture.DocumentPath))
                {
                    fixture.NativeExecutionUnsettled = true;
                    string counters = Convert.ToString(wordTransport.Invoke(wordTarget, ModuleName, "ReadCoverageCounters", new object[] { false, false }));
                    fixture.NativeExecutionUnsettled = false;
                    return counters;
                }
            }
            catch (Exception error)
            {
                Save(fixture, "counter-read-error.json", new { fixture.ProcessId, fixture.DocumentPath,
                    fixture.NativeExecutionUnsettled, Error = error.ToString(), HResult = error.HResult });
                throw;
            }
            finally
            {
                if (fixture.NativeExecutionUnsettled)
                {
                    // The uncertain target borrows this context; retain its independent acquisitions.
                    lock (retainedCounterContexts) retainedCounterContexts.Add(new[] { sourceProject, sourceDocument, documents, application });
                }
                else
                {
                    if (sourceProject != null && Marshal.IsComObject(sourceProject)) Marshal.ReleaseComObject(sourceProject);
                    if (sourceDocument != null && Marshal.IsComObject(sourceDocument)) Marshal.ReleaseComObject(sourceDocument);
                    if (documents != null && Marshal.IsComObject(documents)) Marshal.ReleaseComObject(documents);
                    if (Marshal.IsComObject(application)) Marshal.ReleaseComObject(application);
                }
            }
        }

        private static void AssertReport(IDictionary<string, object> report, string revision, string[] names)
        {
            Assert.AreEqual(revision, report["revision"]);
            Assert.AreEqual(false, report["uncertain"]);
            Assert.IsTrue(string.IsNullOrEmpty(report["error"] as string), Convert.ToString(report["error"]));
            var tests = ((object[])report["tests"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(names.Length, tests.Length, "Missing or extra verified native results.");
            foreach (string name in names)
            {
                var result = tests.Single(test => (string)test["name"] == ModuleName + "." + name);
                string expected = name == "ABooleanPass" ? "Passed" : name == "DRuntimeError" ? "Error" : "Failed";
                Assert.AreEqual(expected, result["outcome"], name);
                Assert.AreEqual("Test", result["phase"], "Fixture failure contaminated " + name);
                if (name == "DRuntimeError") Assert.AreEqual(5, Convert.ToInt32(result["error"]));
                if (name == "CSwallowedAssertion") StringAssert.Contains((string)result["message"], "Deliberate swallowed assertion");
            }
        }

        private static void ReplaceSource(OfficeVbeFixture fixture, string source, string module = ModuleName)
        {
            var before = fixture.Data("read_module", "Module", module);
            int lines = Convert.ToInt32(fixture.Data("component_properties", "Module", module)["CodeLines"]);
            fixture.Data("replace_lines", "Module", module, "StartLine", 1, "Count", lines,
                "ExpectedSha256", before["Sha256"], "Text", source);
            Assert.AreEqual(source.Replace("\r\n", "\n").TrimEnd('\n'),
                ((string)fixture.Data("read_module", "Module", module)["Code"]).Replace("\r\n", "\n").TrimEnd('\n'),
                "Synthetic test source was not written exactly.");
        }

        private static VbaTestCatalog ReviewCatalogue(IDictionary<string, object> catalog)
        {
            var review = new VbaTestCatalog { Project = new VbaTestProjectSnapshot() };
            foreach (var item in ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object))
                review.Modules.Add(new VbaTestModule { Name = (string)item["Name"], Diagnostic = item["Diagnostic"] as string,
                    Tests = ((object[])item["Tests"]).Select(Descriptor).ToList(),
                    ModuleInitialize = Descriptor(item["ModuleInitialize"]), ModuleCleanup = Descriptor(item["ModuleCleanup"]),
                    TestInitialize = Descriptor(item["TestInitialize"]), TestCleanup = Descriptor(item["TestCleanup"]) });
            return review;
        }

        private static VbaTestDescriptor Descriptor(object value)
        {
            if (value == null) return null;
            var item = VbeBridgeClient.Object(value);
            return new VbaTestDescriptor { Id = (string)item["Id"], Module = (string)item["Module"],
                Procedure = (string)item["Procedure"], Kind = (string)item["Kind"], Line = Convert.ToInt32(item["Line"]),
                Diagnostic = item["Diagnostic"] as string, IgnoreReason = item["IgnoreReason"] as string };
        }

        private static void AssertSource(OfficeVbeFixture fixture, IDictionary<string, object> expected, string module = ModuleName)
        {
            var current = fixture.Data("read_module", "Module", module);
            Assert.AreEqual(expected["Sha256"], current["Sha256"], "Test source changed during execution.");
            Assert.AreEqual(expected["Code"], current["Code"]);
        }

        private static object RequireCallback(string loadedAssembly)
        {
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var server = classes.OpenSubKey(@"CLSID\" + RuntimeClsid + @"\InprocServer32"))
            using (var progId = classes.OpenSubKey(@"VBAi.TestRuntime\CLSID"))
            {
                Assert.IsNotNull(server, "VBAi.TestRuntime is not registered; callback qualification cannot use a fallback.");
                Assert.IsNotNull(progId, "VBAi.TestRuntime ProgID is not registered.");
                Assert.AreEqual(RuntimeClsid, Convert.ToString(progId.GetValue("")), true);
                string codeBase = Convert.ToString(server.GetValue("CodeBase"));
                Uri uri; Assert.IsTrue(Uri.TryCreate(codeBase, UriKind.Absolute, out uri) && uri.IsFile);
                Assert.AreEqual(Path.GetFullPath(loadedAssembly), Path.GetFullPath(uri.LocalPath), true,
                    "Callback registration does not point at the exact loaded candidate.");
                Assert.AreEqual(typeof(VbaTestRuntime).FullName, server.GetValue("Class"));
                Assert.AreEqual(typeof(VbaTestRuntime).Assembly.FullName, server.GetValue("Assembly"));
                Assert.AreEqual("v4.0.30319", server.GetValue("RuntimeVersion"));
                Assert.AreEqual("Both", Convert.ToString(server.GetValue("ThreadingModel")), true);
                Assert.AreEqual("mscoree.dll", Convert.ToString(server.GetValue("")), true);
                return new { CodeBase = codeBase, Class = server.GetValue("Class"), Assembly = server.GetValue("Assembly"),
                    ProgId = "VBAi.TestRuntime", Clsid = RuntimeClsid, VerifiedLoadedAssembly = loadedAssembly };
            }
        }

        private static void Save(OfficeVbeFixture fixture, string name, object value)
        {
            File.WriteAllText(Path.Combine(fixture.Root, name), new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value), Utf8);
        }

        private static string Git(string arguments)
        {
            var start = new ProcessStartInfo("git", arguments) { WorkingDirectory = Environment.CurrentDirectory,
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using (var process = Process.Start(start))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                Assert.IsTrue(process.WaitForExit(10000), "Source revision query timed out.");
                Assert.AreEqual(0, process.ExitCode, "Source revision unavailable: " + error);
                return output;
            }
        }

        private static readonly string DocumentCoverageSource = Source.Replace(
            "ABooleanPass = mReady And (mInitialized = 1)",
            "ABooleanPass = mReady And (mInitialized = 1) And (VBAiOfficeProduction.EnteredFunction() = 5)") + @"
Public Function ReadCoverageCounters(Optional ByVal first As Variant, Optional ByVal second As Variant) As String
    ReadCoverageCounters = CStr(VBAiOfficeProduction.EnteredCalls) & "","" & CStr(VBAiOfficeProduction.UnenteredCalls)
End Function
";

        private const string ProductionSource = @"Option Explicit
Public EnteredCalls As Long
Public UnenteredCalls As Long
Public Function EnteredFunction() As Long
    EnteredCalls = EnteredCalls + 1
    EnteredFunction = 5
End Function
Public Function UnenteredFunction() As Long
    UnenteredCalls = UnenteredCalls + 1
    UnenteredFunction = 9
End Function
";

        private const string Source = @"Option Explicit
'@TestModule
Private mInitialized As Long
Private mCleanup As Long
Private mReady As Boolean
'@ModuleInitialize
Public Sub InitializeModule()
    mInitialized = 0
    mCleanup = 0
    mReady = False
End Sub
'@TestInitialize
Public Sub InitializeTest()
    VBAiTestSupport.AreEqual mInitialized, mCleanup, ""Previous test cleanup must have completed.""
    mInitialized = mInitialized + 1
    mReady = True
End Sub
'@TestCleanup
Public Sub CleanupTest()
    mCleanup = mCleanup + 1
    mReady = False
    VBAiTestSupport.AreEqual mInitialized, mCleanup, ""Each entered test must be cleaned up exactly once.""
End Sub
'@ModuleCleanup
Public Sub CleanupModule()
    VBAiTestSupport.AreEqual mInitialized, mCleanup, ""All entered tests must be cleaned up.""
    VBAiTestSupport.IsFalse mReady, ""Test cleanup must reset the fixture.""
End Sub
'@TestMethod
Public Function ABooleanPass() As Boolean
    ABooleanPass = mReady And (mInitialized = 1)
End Function
'@TestMethod
Public Function BBooleanFail() As Boolean
    BBooleanFail = False
End Function
'@TestMethod
Public Sub CSwallowedAssertion()
    On Error Resume Next
    VBAiTestSupport.Fail ""Deliberate swallowed assertion must remain failed.""
End Sub
'@TestMethod
Public Sub DRuntimeError()
    Err.Raise 5, ""OfficeExplorerFixture"", ""Deliberate runtime error.""
End Sub
";
    }
}
