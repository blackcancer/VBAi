using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    /// <summary>Qualifies the installed Excel test service and its real registered result callback on one disposable workbook.</summary>
    [TestClass, TestCategory("VbaTestExcelRegistered"), DoNotParallelize]
    public sealed class ExcelRegisteredTestExplorerTests
    {
        private const string TestsModule = "RegisteredExcelTests", ProductionModule = "RegisteredProduction";
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        /// <summary>Checks native callback outcomes, fixture counts, copy coverage, stale refusal, tool-window ownership and normal shutdown.</summary>
        [STATestMethod]
        public void InstalledBatchSingleCopyCoverageAndExplorerPreserveOriginalWorkbook()
            => ExecuteRegisteredQualification(1);

        /// <summary>Runs three independent coverage copies in one owned Excel session, checking detachment after each.</summary>
        [STATestMethod]
        public void InstalledRepeatedCopyCoveragePreservesOriginalAndDetachesEveryProject()
            => ExecuteRegisteredQualification(3);

        private static void ExecuteRegisteredQualification(int coverageCycles)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Registered Excel qualification requires VBAi_RUN_EXCEL_TESTS=1.");
            ExcelVbeFixture.Run(host =>
            {
                var status = Data(host, new { Command = "status" });
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                Save(host, "callback-registration.json", VbaTestExplorerOfficeTests.RequireCallback((string)status["AssemblyPath"]));
                var projects = (object[])DataValue(host, new { Command = "list_projects" });
                Assert.AreEqual(1, projects.Length, "Only the disposable workbook may be present.");
                string selector = Convert.ToString(VbeBridgeClient.Object(projects[0])["Name"]);
                Create(host, selector, TestsModule, TestSource);
                Create(host, selector, ProductionModule, ProductionSource);
                var catalog = Data(host, new { Command = "discover_vba_tests", Project = selector });
                var preview = Data(host, new { Command = "preview_vba_test_support", Project = selector });
                string reviewed = VbaTestRuntimeSource.Generate(Review(catalog));
                Assert.AreEqual(reviewed, preview["Text"], "Support must match the complete exact-candidate generator output.");
                StringAssert.Contains(reviewed, "CreateObject(\"VBAi.TestRuntime\")");
                Data(host, new { Command = "install_vba_test_support", Project = selector,
                    ExpectedProjectVersion = preview["ExpectedProjectVersion"], ExpectedMode = 2, Text = reviewed }, true);
                string path = host.File("RegisteredExplorer.xlsm");
                var properties = Data(host, new { Command = "project_properties", Project = selector });
                Data(host, new { Command = "save_host_document_as", Project = selector, Path = path,
                    ExpectedProjectVersion = properties["Version"] }, true);
                selector = path;
                catalog = Data(host, new { Command = "discover_vba_tests", Project = selector });
                Assert.IsTrue(string.IsNullOrEmpty(catalog["ExecutionUnavailableReason"] as string),
                    Convert.ToString(catalog["ExecutionUnavailableReason"]));
                var tests = Review(catalog).Tests.OrderBy(test => test.Procedure, StringComparer.Ordinal).ToArray();
                Assert.AreEqual(4, tests.Length);
                string revision = Convert.ToString(catalog["ExpectedProjectVersion"]);
                var originals = ((object[])DataValue(host, new { Command = "list_modules", Project = selector }))
                    .Select(VbeBridgeClient.Object).ToDictionary(item => (string)item["Name"],
                        item => Data(host, new { Command = "read_module", Project = selector, Module = item["Name"] }), StringComparer.Ordinal);
                byte[] disk = ReadBytes(path);
                var batch = Run(host, selector, revision, tests.Select(test => test.Id).ToArray(), "batch");
                AssertResults(batch, revision, new[] { "Passed", "Failed", "Failed", "Error" });
                var passed = tests.Single(test => test.Procedure == "AExpectedValue");
                AssertResults(Run(host, selector, revision, new[] { passed.Id }, "single"), revision, new[] { "Passed" });
                Assert.AreEqual(2d, Convert.ToDouble(host.ReadCell("A1")), "Only the selected pass test calls production.");
                Assert.AreEqual(5d, Convert.ToDouble(host.ReadCell("C1")), "Each of the five test invocations must initialize once.");
                Assert.AreEqual(5d, Convert.ToDouble(host.ReadCell("D1")), "False results, swallowed assertions and runtime errors must still clean up.");
                var coveragePreview = Data(host, new { Command = "vba_test_coverage", Project = selector });
                Assert.AreEqual(true, coveragePreview["Supported"]);
                Assert.AreEqual(2, Convert.ToInt32(coveragePreview["EligibleProcedureCount"]));
                IDictionary<string, object> covered = null;
                var queries = new HashSet<string>(StringComparer.Ordinal);
                for (int cycle = 0; cycle < coverageCycles; cycle++)
                {
                    string phase = coverageCycles == 1 ? "coverage" : "coverage-cycle-" + (cycle + 1);
                    covered = Run(host, selector, revision, new[] { passed.Id }, phase, true);
                    Assert.IsTrue(queries.Add(Convert.ToString(covered["run"])), "A new measurement reused an earlier run identity.");
                    AssertResults(covered, revision, new[] { "Passed" });
                    var coverage = VbeBridgeClient.Object(covered["coverage"]);
                    Assert.AreEqual(true, coverage["available"]); Assert.AreEqual(true, coverage["complete"]);
                    Assert.AreEqual(2, Convert.ToInt32(coverage["eligible"])); Assert.AreEqual(1, Convert.ToInt32(coverage["hit"]));
                    Assert.AreEqual(50d, Convert.ToDouble(coverage["percent"]));
                    Assert.AreEqual(false, coverage["statementCoverageAvailable"]); Assert.AreEqual(false, coverage["branchCoverageAvailable"]);
                    Assert.AreEqual(2d, Convert.ToDouble(host.ReadCell("A1")), "Coverage must execute in its copy, never the original.");
                    Assert.AreEqual(5d, Convert.ToDouble(host.ReadCell("C1"))); Assert.AreEqual(5d, Convert.ToDouble(host.ReadCell("D1")));
                    CollectionAssert.AreEqual(disk, ReadBytes(path), "The original saved bytes must remain unchanged.");
                    Assert.AreEqual(revision, Data(host, new { Command = "discover_vba_tests", Project = selector })["ExpectedProjectVersion"]);
                    foreach (var original in originals)
                        Assert.AreEqual(original.Value["Sha256"], Data(host, new { Command = "read_module", Project = selector, Module = original.Key })["Sha256"]);
                    object books=null, book=null;
                    var liveProjects=(object[])DataValue(host, new { Command = "list_projects" });
                    try
                    {
                        object application=VBAi.Tests.Infrastructure.UiInvoke.Field<object>(host,"application");
                        books=((dynamic)application).Workbooks;
                        int count=Convert.ToInt32(((dynamic)books).Count);
                        Assert.AreEqual(1,count,"The copied workbook must be absent from Excel's document collection.");
                        book=((dynamic)books)[1];
                        Assert.AreEqual(path,Convert.ToString(((dynamic)book).FullName),true);
                        Save(host,phase + "-document-absence.json",new { WorkbookCount=count,OriginalPath=path,VbeProjects=liveProjects });
                    }
                    finally
                    {
                        if(book!=null && Marshal.IsComObject(book)) Marshal.ReleaseComObject(book);
                        if(books!=null && Marshal.IsComObject(books)) Marshal.ReleaseComObject(books);
                    }
                    Assert.AreEqual(1, liveProjects.Length, "The completed copy must also leave the VBE project collection.");
                    Save(host, phase + "-preservation.json", new { Run = covered["run"], OriginalSourceAndBytesPreserved = true,
                        EveryCopyDetached = true, OriginalProductionCalls = 2, OriginalFixtureInvocations = 5 });
                }
                Assert.AreEqual(coverageCycles, queries.Count);
                Save(host, "coverage-cycles-complete.json", new { Cycles = coverageCycles, Queries = queries.ToArray(),
                    OriginalSourceAndBytesPreserved = true, EveryCopyDetached = true, NativeActionsReplayed = false });
                var shown = Data(host, new { Command = "show_vba_test_explorer", Project = selector });
                Assert.AreEqual(true, shown["Opened"]); Assert.AreEqual(true, shown["Docked"]);
                var window = new IntPtr(Convert.ToInt64(shown["Hwnd"]));
                Assert.IsTrue(IsWindow(window)); uint owner;
                Assert.AreNotEqual(0u, GetWindowThreadProcessId(window, out owner)); Assert.AreEqual((uint)host.ProcessId, owner);
                var current = Data(host, new { Command = "read_module", Project = selector, Module = TestsModule });
                Data(host, new { Command = "replace_lines", Project = selector, Module = TestsModule, StartLine = 1,
                    Count = ((string)current["Code"]).Split('\n').Length - 1, ExpectedSha256 = current["Sha256"],
                    Text = ((string)current["Code"]).TrimEnd() + "\r\n' deliberate stale catalogue\r\n" }, true);
                var refusal = host.Command(new { Command = "run_vba_tests", Project = selector,
                    ExpectedProjectVersion = revision, ExpectedMode = 2, Items = new[] { passed.Id } });
                Assert.AreEqual(false, refusal["Ok"]); StringAssert.Contains(Convert.ToString(refusal["Error"]), "ExpectedProjectVersion");
                Assert.AreEqual(2d, Convert.ToDouble(host.ReadCell("A1")), "A stale request must not execute a test.");
                Save(host, "registered-explorer.json", new { Status=status, Batch=batch, Coverage=covered,
                    Explorer=shown, StaleRefusal=refusal, CoverageCycles=coverageCycles, SourcePreserved=true, SavedBytesPreserved=true, OriginalProductionCalls=2 });
            });
        }

        private static IDictionary<string, object> Run(ExcelVbeFixture host, string project, string revision, string[] ids, string name, bool coverage = false)
        {
            host.PreserveForDiagnosticRecovery = true;
            var start = Data(host, new { Command = "run_vba_tests", Project = project,
                ExpectedProjectVersion = revision, ExpectedMode = 2, Items = ids, Action = coverage ? "coverage" : null });
            Save(host, name + "-start.json", start);
            string query = Convert.ToString(start["Query"]); Assert.IsFalse(string.IsNullOrEmpty(query));
            DateTime deadline = DateTime.UtcNow.AddSeconds(60);
            IDictionary<string, object> state;
            do
            {
                state = Data(host, new { Command = "vba_test_run_status", Project = project, Query = query, Action = "compact" });
                if (Equals(state["Pending"], false)) break;
                Thread.Sleep(100);
            } while (DateTime.UtcNow < deadline);
            Save(host, name + "-status.json", state);
            Assert.AreEqual(false, state["Pending"], "Uncertain or timed-out execution is retained without cleanup or retry.");
            var report = VbeBridgeClient.Object(state["Report"]);
            if (Equals(report["uncertain"], false)) host.PreserveForDiagnosticRecovery = false;
            Assert.AreEqual("Completed", state["State"]); Assert.AreEqual(false, state["Stale"]);
            return report;
        }

        private static void AssertResults(IDictionary<string, object> report, string revision, string[] expected)
        {
            Assert.AreEqual(revision, report["revision"]); Assert.AreEqual(false, report["uncertain"]);
            Assert.IsTrue(string.IsNullOrEmpty(report["error"] as string), Convert.ToString(report["error"]));
            var results = ((object[])report["tests"]).Select(VbeBridgeClient.Object).OrderBy(item => (string)item["name"], StringComparer.Ordinal).ToArray();
            CollectionAssert.AreEqual(expected, results.Select(item => (string)item["outcome"]).ToArray());
            foreach (var result in results) Assert.AreEqual("Test", result["phase"]);
            foreach (var result in results.Where(item => Equals(item["outcome"], "Error"))) Assert.AreEqual(5, Convert.ToInt32(result["error"]));
        }

        private static void Create(ExcelVbeFixture host, string project, string module, string text)
        {
                var created = Data(host, new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 }, true);
            var before = Data(host, new { Command = "read_module", Project = project, Module = module });
            Data(host, new { Command = "replace_lines", Project = project, Module = module, StartLine = 1,
                Count = created["Lines"], ExpectedSha256 = before["Sha256"], ExpectedMode = 2, Text = text }, true);
            Assert.AreEqual(text.Replace("\r\n", "\n").TrimEnd('\n'),
                ((string)Data(host, new { Command = "read_module", Project = project, Module = module })["Code"]).Replace("\r\n", "\n").TrimEnd('\n'));
        }

        private static object DataValue(ExcelVbeFixture host, object request, bool mutation = false)
        {
            if (mutation) host.PreserveForDiagnosticRecovery = true;
            var reply = host.Command(request);
            Assert.IsNotNull(reply); Assert.AreEqual(true, reply["Ok"], Convert.ToString(reply["Error"]));
            var data = reply["Data"];
            if (mutation)
            {
                var fields = data as IDictionary<string, object>;
                object uncertain;
                if (fields != null && fields.TryGetValue("Uncertain", out uncertain)) Assert.AreEqual(false, uncertain);
                host.PreserveForDiagnosticRecovery = false;
            }
            return data;
        }
        private static IDictionary<string, object> Data(ExcelVbeFixture host, object request, bool mutation = false)
        { return VbeBridgeClient.Object(DataValue(host, request, mutation)); }
        private static byte[] ReadBytes(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var bytes = new MemoryStream()) { file.CopyTo(bytes); return bytes.ToArray(); }
        }
        private static void Save(ExcelVbeFixture host, string name, object value)
        { File.WriteAllText(host.File(name), new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 }.Serialize(value), new UTF8Encoding(false)); }
        private static VbaTestCatalog Review(IDictionary<string, object> catalog)
        {
            var review = new VbaTestCatalog { Project = new VbaTestProjectSnapshot() };
            foreach (var module in ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object))
                review.Modules.Add(new VbaTestModule { Name=(string)module["Name"], Diagnostic=module["Diagnostic"] as string,
                    Tests=((object[])module["Tests"]).Select(Descriptor).ToList(), ModuleInitialize=Descriptor(module["ModuleInitialize"]),
                    ModuleCleanup=Descriptor(module["ModuleCleanup"]), TestInitialize=Descriptor(module["TestInitialize"]), TestCleanup=Descriptor(module["TestCleanup"]) });
            return review;
        }
        private static VbaTestDescriptor Descriptor(object value)
        {
            if (value == null) return null;
            var item = VbeBridgeClient.Object(value);
            return new VbaTestDescriptor { Id=(string)item["Id"], Module=(string)item["Module"], Procedure=(string)item["Procedure"],
                Kind=(string)item["Kind"], Line=Convert.ToInt32(item["Line"]), Diagnostic=item["Diagnostic"] as string, IgnoreReason=item["IgnoreReason"] as string };
        }
        private const string ProductionSource = "Option Explicit\r\nPublic Function EnteredValue() As Long\r\n" +
            "    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1\r\n    EnteredValue = 7\r\nEnd Function\r\n" +
            "Public Function UnenteredValue() As Long\r\n    UnenteredValue = 9\r\nEnd Function\r\n";
        private const string TestSource = "Option Explicit\r\n'@TestModule\r\n'@TestInitialize\r\nPublic Sub InitializeTest()\r\n" +
            "    ThisWorkbook.Worksheets(1).Range(\"C1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"C1\").Value2 + 1\r\nEnd Sub\r\n" +
            "'@TestCleanup\r\nPublic Sub CleanupTest()\r\n    ThisWorkbook.Worksheets(1).Range(\"D1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"D1\").Value2 + 1\r\nEnd Sub\r\n" +
            "'@TestMethod\r\nPublic Function AExpectedValue() As Boolean\r\n    AExpectedValue = (RegisteredProduction.EnteredValue() = 7)\r\nEnd Function\r\n" +
            "'@TestMethod\r\nPublic Function BDeliberateFalse() As Boolean\r\n    BDeliberateFalse = False\r\nEnd Function\r\n" +
            "'@TestMethod\r\nPublic Sub CSwallowedAssertion()\r\n    On Error Resume Next\r\n    VBAiTestSupport.Fail \"Deliberate swallowed assertion\"\r\nEnd Sub\r\n" +
            "'@TestMethod\r\nPublic Sub DDeliberateError()\r\n    Err.Raise 5, \"RegisteredExcelFixture\", \"Deliberate runtime error\"\r\nEnd Sub\r\n";
    }
}
