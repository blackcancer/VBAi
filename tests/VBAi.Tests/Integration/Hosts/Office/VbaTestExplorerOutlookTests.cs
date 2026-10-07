using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies the registered add-in in a fresh owned Outlook process with an initially absent OTM without changing trust or registration.</summary>
    [TestClass, TestCategory("VbaTestOutlook"), DoNotParallelize]
    public sealed class VbaTestExplorerOutlookTests
    {
        private const string ModuleName = "VBAiOutlookModule";
        private const string RuntimeClsid = "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        /// <summary>Uses only explicit discovered IDs and the installed native callback, never a substitute execution adapter.</summary>
        [STATestMethod]
        [Timeout(240000)]
        public void RegisteredOutlookRunsBatchAndSingleWithVerifiedResults()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OUTLOOK_TESTS") != "1")
                Assert.Inconclusive("Native Outlook tests require VBAi_RUN_OUTLOOK_TESTS=1.");
            // Capture provenance before COM activation; an unavailable source identity must not become a fabricated result.
            const string host = "Outlook";
            string sourceRevision = Git("rev-parse HEAD").Trim();
            string sourceStatus = Git("status --porcelain");
            using (var fixture = OutlookVbaTestFixture.Start())
            {
                Console.WriteLine("Registered Outlook qualification=" + fixture.Root);
                var status = fixture.Data("status");
                Save(fixture, "identity.json", new
                {
                    CurrentSourceRevision = sourceRevision,
                    SourceStatus = sourceStatus,
                    Host = host,
                    fixture.ProcessId,
                    fixture.OtmPath,
                    fixture.Project,
                    LoadedAssembly = status,
                    ExecutionBoundary = "Registered in-process add-in with registered VBAi.TestRuntime callback"
                });
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Save(fixture, "callback-registration.json", RequireCallback((string)status["AssemblyPath"]));

                fixture.Data("create_module", "Module", ModuleName, "ExpectedMode", 2);
                ReplaceSource(fixture, Source);
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
                fixture.TrackOwnedModule(VbaTestRuntimeSource.ModuleName);
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

        private static IDictionary<string, object> Run(OutlookVbaTestFixture fixture, string revision, string[] ids, string name)
        {
            var started = fixture.Data("run_vba_tests", "ExpectedProjectVersion", revision, "ExpectedMode", 2, "Items", ids);
            Save(fixture, name + "-start.json", started);
            string query = (string)started["Query"];
            Assert.IsFalse(string.IsNullOrEmpty(query));
            var deadline = DateTime.UtcNow.AddSeconds(60);
            IDictionary<string, object> state;
            do
            {
                state = fixture.Data("vba_test_run_status", "Query", query, "Action", "compact");
                if (!Convert.ToBoolean(state["Pending"])) break;
                Thread.Sleep(100);
            } while (DateTime.UtcNow < deadline);
            Save(fixture, name + "-status.json", state);
            Assert.AreEqual(false, state["Pending"], "Native run observation timed out. No retry or reset was attempted.");
            Assert.AreEqual("Completed", state["State"]);
            Assert.AreEqual(false, state["Stale"]);
            var report = VbeBridgeClient.Object(state["Report"]);
            Save(fixture, name + "-compact.json", report);
            var human = fixture.Data("vba_test_run_status", "Query", query, "Action", "human");
            File.WriteAllText(Path.Combine(fixture.Root, name + "-human.txt"), (string)human["Report"], Utf8);
            Assert.AreEqual(query, report["run"]);
            return report;
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

        private static void ReplaceSource(OutlookVbaTestFixture fixture, string source, string module = ModuleName)
        {
            var before = fixture.Data("read_module", "Module", module);
            int lines = Convert.ToInt32(fixture.Data("component_properties", "Module", module)["CodeLines"]);
            fixture.Data("replace_lines", "Module", module, "StartLine", 1, "Count", lines,
                "ExpectedSha256", before["Sha256"], "Text", source);
            Assert.AreEqual(source.Replace("\r\n", "\n").TrimEnd('\n'),
                ((string)fixture.Data("read_module", "Module", module)["Code"]).Replace("\r\n", "\n").TrimEnd('\n'),
                "Synthetic test source was not written exactly.");
            fixture.TrackOwnedModule(ModuleName);
        }

        private static VbaTestCatalog ReviewCatalogue(IDictionary<string, object> catalog)
        {
            var review = new VbaTestCatalog { Project = new VbaTestProjectSnapshot() };
            foreach (var item in ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object))
                review.Modules.Add(new VbaTestModule
                {
                    Name = (string)item["Name"],
                    Diagnostic = item["Diagnostic"] as string,
                    Tests = ((object[])item["Tests"]).Select(Descriptor).ToList(),
                    ModuleInitialize = Descriptor(item["ModuleInitialize"]),
                    ModuleCleanup = Descriptor(item["ModuleCleanup"]),
                    TestInitialize = Descriptor(item["TestInitialize"]),
                    TestCleanup = Descriptor(item["TestCleanup"])
                });
            return review;
        }

        private static VbaTestDescriptor Descriptor(object value)
        {
            if (value == null) return null;
            var item = VbeBridgeClient.Object(value);
            return new VbaTestDescriptor
            {
                Id = (string)item["Id"],
                Module = (string)item["Module"],
                Procedure = (string)item["Procedure"],
                Kind = (string)item["Kind"],
                Line = Convert.ToInt32(item["Line"]),
                Diagnostic = item["Diagnostic"] as string,
                IgnoreReason = item["IgnoreReason"] as string
            };
        }

        private static void AssertSource(OutlookVbaTestFixture fixture, IDictionary<string, object> expected, string module = ModuleName)
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
                return new
                {
                    CodeBase = codeBase,
                    Class = server.GetValue("Class"),
                    Assembly = server.GetValue("Assembly"),
                    ProgId = "VBAi.TestRuntime",
                    Clsid = RuntimeClsid,
                    VerifiedLoadedAssembly = loadedAssembly
                };
            }
        }

        private static void Save(OutlookVbaTestFixture fixture, string name, object value)
        {
            File.WriteAllText(Path.Combine(fixture.Root, name), new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value), Utf8);
        }

        private static string Git(string arguments)
        {
            var start = new ProcessStartInfo("git", arguments)
            {
                WorkingDirectory = Environment.CurrentDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using (var process = Process.Start(start))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                Assert.IsTrue(process.WaitForExit(10000), "Source revision query timed out.");
                Assert.AreEqual(0, process.ExitCode, "Source revision unavailable: " + error);
                return output;
            }
        }

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
    Err.Raise 5, ""OutlookExplorerFixture"", ""Deliberate runtime error.""
End Sub
";
    }
}
