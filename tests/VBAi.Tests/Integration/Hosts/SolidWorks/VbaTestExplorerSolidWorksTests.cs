using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies only an explicitly selected, preloaded SOLIDWORKS process and an already opened blank disposable SWP.</summary>
    [TestClass, TestCategory("VbaTestSolidWorks"), DoNotParallelize]
    public sealed class VbaTestExplorerSolidWorksTests
    {
        private const string ModuleName = "VBAiSolidWorksTests";
        private const string RuntimeClsid = "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out WindowRect rectangle);
        [StructLayout(LayoutKind.Sequential)] private struct WindowRect { public int Left, Top, Right, Bottom; }

        [TestMethod, Timeout(240000)]
        public void PreloadedBlankMacroRunsVerifiedBatchSingleReportsAndExplorer()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_SOLIDWORKS_TEST_EXPLORER") != "1")
                Assert.Inconclusive("SOLIDWORKS Test Explorer qualification requires VBAi_RUN_SOLIDWORKS_TEST_EXPLORER=1.");
            string suppliedPid = Environment.GetEnvironmentVariable("VBAi_SOLIDWORKS_PID");
            string suppliedPath = Environment.GetEnvironmentVariable("VBAi_SOLIDWORKS_TEST_MACRO");
            if (string.IsNullOrWhiteSpace(suppliedPid) || string.IsNullOrWhiteSpace(suppliedPath))
                Assert.Inconclusive("Supply VBAi_SOLIDWORKS_PID and VBAi_SOLIDWORKS_TEST_MACRO for a blank saved disposable SWP already opened through Tools > Macro > Edit.");
            int pid;
            Assert.IsTrue(int.TryParse(suppliedPid, out pid) && pid > 0);
            Assert.IsTrue(Path.IsPathRooted(suppliedPath), "The explicitly authorized macro path must be absolute.");
            string path = Path.GetFullPath(suppliedPath);
            Assert.AreEqual(".swp", Path.GetExtension(path), true);
            Assert.IsTrue(File.Exists(path), "The disposable macro must already exist; this fixture never opens or creates an SWP.");
            Assert.IsFalse((File.GetAttributes(path) & (FileAttributes.ReparsePoint | FileAttributes.ReadOnly)) != 0);
            string root = Path.Combine(Path.GetTempPath(), "VBAi-TestExplorer-SolidWorks", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string revision = Git("rev-parse HEAD").Trim(), sourceStatus = Git("status --porcelain");
            using (var process = Process.GetProcessById(pid))
            {
                Assert.AreEqual("SLDWORKS", process.ProcessName, true);
                Assert.IsFalse(process.HasExited);
                _ = process.Handle; // Retain the exact selected process identity without activating COM.
                var fixture = new Fixture(process, path, root);
                Exception primary = null, cleanup = null;
                try { fixture.Preflight(revision, sourceStatus); fixture.Run(); }
                catch (Exception error) { primary = error; }
                finally
                {
                    try { fixture.RestoreOwnedModules(); }
                    catch (Exception error) { cleanup = error; }
                    try { fixture.Save("lifecycle.json", new { PID = pid, Macro = path, PrimaryError = primary?.ToString(), CleanupError = cleanup?.ToString(),
                        fixture.Unsettled, ProcessRetained = !process.HasExited, MacroLeftOpen = true,
                        OtherProjectSourcesInspected = false, CoverageQualification = "Not requested; prerequisites require a separate decision" }); }
                    catch (Exception error) { cleanup = cleanup == null ? error : new AggregateException(cleanup, error); }
                    Console.WriteLine("SOLIDWORKS qualification evidence=" + root);
                }
                if (primary != null && cleanup != null) throw new AggregateException("Qualification failed; cleanup failure is preserved separately.", primary, cleanup);
                if (primary != null) ExceptionDispatchInfo.Capture(primary).Throw();
                if (cleanup != null) ExceptionDispatchInfo.Capture(cleanup).Throw();
                Assert.IsFalse(process.HasExited, "The selected SOLIDWORKS process unexpectedly exited.");
            }
        }

        private sealed class Fixture
        {
            private readonly Process process;
            private readonly string path, root;
            private readonly Dictionary<string, string> owned = new Dictionary<string, string>(StringComparer.Ordinal);
            private IDictionary<string, object>[] baseline;
            private string referencesVersion, diskHash;
            private bool baselineVerified;
            internal bool Unsettled { get; private set; }
            internal Fixture(Process process, string path, string root) { this.process = process; this.path = path; this.root = root; }
            internal void Save(string name, object value) => File.WriteAllText(Path.Combine(root, name),
                new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value), Utf8);

            private IDictionary<string, object> Response(string command, params object[] pairs)
            {
                Assert.IsFalse(process.HasExited, "Selected SOLIDWORKS exited; no operation may be sent to a replacement PID.");
                var request = new Dictionary<string, object> { ["Command"] = command, ["Project"] = path };
                for (int index = 0; index < pairs.Length; index += 2) request[(string)pairs[index]] = pairs[index + 1];
                bool mutation = new[] { "create_module", "replace_lines", "install_vba_test_support", "remove_component" }.Contains(command);
                if (mutation || command == "run_vba_tests") Unsettled = true;
                var response = VbeBridgeClient.Read("VBAi." + process.Id, request, 65000, connectionAttempts: 1);
                Assert.IsNotNull(response, "Bridge connection unavailable; no request was resent.");
                if (mutation && Equals(response["Ok"], true))
                {
                    var data = response["Data"] as IDictionary<string, object>;
                    Unsettled = data != null && data.TryGetValue("Uncertain", out object uncertain) && Equals(uncertain, true);
                }
                if (command == "run_vba_tests" && Equals(response["Ok"], false) && Convert.ToString(response["Error"]).Contains("ExpectedProjectVersion"))
                    Unsettled = false;
                return response;
            }
            private IDictionary<string, object> Data(string command, params object[] pairs)
            {
                var response = Response(command, pairs);
                Assert.AreEqual(true, response["Ok"], command + ": " + Convert.ToString(response["Error"]));
                return VbeBridgeClient.Object(response["Data"]);
            }
            private IDictionary<string, object>[] Items(string command)
            {
                var response = Response(command);
                Assert.AreEqual(true, response["Ok"], command + ": " + Convert.ToString(response["Error"]));
                return ((object[])response["Data"]).Select(VbeBridgeClient.Object).ToArray();
            }
            internal void Preflight(string revision, string sourceStatus)
            {
                var status = Data("status");
                Assert.AreEqual(process.Id, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(64, Convert.ToInt32(status["ProcessBitness"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                string loaded = (string)status["AssemblyPath"];
                Assert.AreEqual(HashFile(typeof(VbeSession).Assembly.Location), HashFile(loaded), "Loaded candidate bytes differ from the test assembly reference.");
                Save("identity.json", new { CurrentSourceRevision = revision, SourceStatus = sourceStatus, LoadedAssembly = status,
                    PID = process.Id, HostFileVersion = process.MainModule.FileVersionInfo.FileVersion, Macro = path,
                    Boundary = "registered in-process VBE callback; preloaded SOLIDWORKS retained" });
                Save("callback-registration.json", RequireCallback(loaded));
                // Exact path resolution excludes other projects even if their VBA names collide.
                var properties = Data("project_properties");
                Assert.AreEqual(2, Convert.ToInt32(properties["Mode"]));
                var fields = ((object[])properties["Properties"]).Select(VbeBridgeClient.Object).ToArray();
                string type = Convert.ToString(fields.Single(field => (string)field["Name"] == "Type")["Value"]);
                string protection = Convert.ToString(fields.Single(field => (string)field["Name"] == "Protection")["Value"]);
                Assert.IsTrue(type == "100" || type == "vbext_pt_HostProject", "The selected SWP must be a host project.");
                Assert.IsTrue(protection == "0" || protection == "vbext_pp_none", "The selected SWP must be unprotected.");
                var persistence = Data("project_persistence_status");
                Assert.AreEqual("SOLIDWORKS", persistence["Host"]);
                Assert.AreEqual(true, persistence["HostAvailable"]);
                Assert.AreEqual(process.Id, Convert.ToInt32(persistence["OwnerProcessId"]));
                Assert.AreEqual(path, Path.GetFullPath((string)persistence["HostPath"]), true);
                Assert.AreEqual(true, persistence["ProjectSaved"], "Save the blank disposable macro before qualification.");
                baseline = ReadSources();
                Assert.IsTrue(baseline.Length >= 1 && baseline.All(module => Convert.ToInt32(module["Type"]) == 1), "Only blank standard modules are accepted.");
                foreach (var module in baseline)
                    Assert.IsTrue(Regex.IsMatch((string)module["Code"], @"\A\s*(?:Option\s+Explicit\s*)?\z", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
                        "The selected macro must be blank: no procedures, comments or arbitrary existing code may be executed or edited.");
                Assert.IsFalse(baseline.Any(module => (string)module["Name"] == ModuleName || (string)module["Name"] == VbaTestRuntimeSource.ModuleName));
                var references = Data("list_references");
                referencesVersion = (string)references["Version"];
                Save("baseline-sources.json", baseline); Save("baseline-references.json", references); Save("baseline-project.json", properties);
                diskHash = HashFile(path);
                File.Copy(path, Path.Combine(root, "baseline.swp.backup"), false);
                Assert.AreEqual(diskHash, HashFile(Path.Combine(root, "baseline.swp.backup")));
                Assert.AreEqual(diskHash, HashFile(path), "The macro file changed during backup; qualification was refused.");
                baselineVerified = true;
            }
            private IDictionary<string, object>[] ReadSources() => Items("list_modules").Select(module => {
                var source = Data("read_module", "Module", module["Name"]);
                return (IDictionary<string, object>)new Dictionary<string, object> { ["Name"] = module["Name"], ["Type"] = module["Type"],
                    ["Code"] = source["Code"], ["Sha256"] = source["Sha256"] };
            }).OrderBy(module => (string)module["Name"], StringComparer.Ordinal).ToArray();

            internal void Run()
            {
                Data("create_module", "Module", ModuleName, "ExpectedMode", 2);
                owned[ModuleName] = (string)Data("read_module", "Module", ModuleName)["Sha256"];
                Replace(Source);
                var original = Data("read_module", "Module", ModuleName);
                var discovery = Data("discover_vba_tests");
                var tests = Catalogue(discovery);
                var preview = Data("preview_vba_test_support");
                string support = (string)preview["Text"];
                Assert.AreEqual("3", VbaTestRuntimeSource.Version);
                Assert.AreEqual(VbaTestRuntimeSource.Generate(Review(discovery)), support);
                Assert.IsTrue(support.Length <= 1024 * 1024);
                Save("support-preview.json", preview);
                File.WriteAllText(Path.Combine(root, "reviewed-support.bas"), support, Utf8);
                var installed = Data("install_vba_test_support", "ExpectedProjectVersion", preview["ExpectedProjectVersion"], "ExpectedMode", 2, "Text", support);
                Assert.AreEqual(true, installed["Applied"]);
                owned[VbaTestRuntimeSource.ModuleName] = (string)Data("read_module", "Module", VbaTestRuntimeSource.ModuleName)["Sha256"];
                Assert.AreEqual(Canonical(support), Canonical((string)Data("read_module", "Module", VbaTestRuntimeSource.ModuleName)["Code"]));
                var catalog = Data("discover_vba_tests");
                tests = Catalogue(catalog);
                Save("discovery-ready.json", catalog);
                Assert.IsTrue(string.IsNullOrEmpty(catalog["ExecutionUnavailableReason"] as string), Convert.ToString(catalog["ExecutionUnavailableReason"]));
                string version = (string)catalog["ExpectedProjectVersion"];
                string[] ids = tests.OrderBy(test => (string)test["Procedure"], StringComparer.Ordinal).Select(test => (string)test["Id"]).ToArray();
                Assert.AreEqual(4, ids.Distinct(StringComparer.Ordinal).Count());
                AssertReport(RunNative(version, ids, "batch"), version, tests);
                AssertSource(original);
                var passed = tests.Single(test => (string)test["Procedure"] == "ABooleanPass");
                var single = RunNative(version, new[] { (string)passed["Id"] }, "single");
                AssertReport(single, version, new[] { passed });
                AssertSource(original);
                var shown = Data("show_vba_test_explorer");
                Save("explorer.json", shown);
                Assert.AreEqual(true, shown["Opened"]); Assert.AreEqual(true, shown["Docked"]);
                var hwnd = new IntPtr(Convert.ToInt64(shown["Hwnd"]));
                Assert.IsTrue(IsWindow(hwnd));
                uint owner; Assert.AreNotEqual(0u, GetWindowThreadProcessId(hwnd, out owner)); Assert.AreEqual((uint)process.Id, owner);
                WindowRect bounds; Assert.IsTrue(GetWindowRect(hwnd, out bounds));
                Assert.IsTrue(bounds.Right > bounds.Left && bounds.Bottom > bounds.Top);
                Save("explorer-window.json", new { Hwnd = hwnd.ToInt64(), PID = owner, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom });
                Replace(((string)original["Code"]).TrimEnd() + "\r\n' stale-source qualification\r\n");
                var changed = Data("read_module", "Module", ModuleName);
                var refusal = Response("run_vba_tests", "ExpectedProjectVersion", version, "ExpectedMode", 2, "Items", new[] { (string)passed["Id"] });
                Save("stale-refusal.json", refusal);
                Assert.AreEqual(false, refusal["Ok"]); StringAssert.Contains(Convert.ToString(refusal["Error"]), "ExpectedProjectVersion");
                AssertSource(changed);
                var historical = Data("vba_test_run_status", "Query", single["run"], "Action", "compact");
                Save("stale-history.json", historical);
                Assert.AreEqual(true, historical["Stale"]); Assert.AreEqual(false, historical["Pending"]);
                Assert.AreEqual("Completed", historical["State"]);
            }
            private void Replace(string source)
            {
                var before = Data("read_module", "Module", ModuleName);
                if (owned.TryGetValue(ModuleName, out string hash)) Assert.AreEqual(hash, before["Sha256"], "Concurrent edits refuse fixture mutation.");
                int lines = Convert.ToInt32(Data("component_properties", "Module", ModuleName)["CodeLines"]);
                Data("replace_lines", "Module", ModuleName, "StartLine", 1, "Count", lines, "ExpectedSha256", before["Sha256"], "Text", source);
                var after = Data("read_module", "Module", ModuleName);
                Assert.AreEqual(Canonical(source), Canonical((string)after["Code"]));
                owned[ModuleName] = (string)after["Sha256"];
            }
            private void AssertSource(IDictionary<string, object> expected)
            {
                var actual = Data("read_module", "Module", ModuleName);
                Assert.AreEqual(expected["Sha256"], actual["Sha256"]); Assert.AreEqual(expected["Code"], actual["Code"]);
                Assert.AreEqual(referencesVersion, Data("list_references")["Version"]);
            }
            private IDictionary<string, object> RunNative(string revision, string[] ids, string name)
            {
                var start = Data("run_vba_tests", "ExpectedProjectVersion", revision, "ExpectedMode", 2, "Items", ids);
                Save(name + "-start.json", start);
                string query = (string)start["Query"];
                Assert.IsFalse(string.IsNullOrEmpty(query));
                DateTime deadline = DateTime.UtcNow.AddSeconds(60);
                IDictionary<string, object> state;
                do
                {
                    state = Data("vba_test_run_status", "Query", query, "Action", "compact");
                    if (Equals(state["Pending"], false)) break;
                    Thread.Sleep(100);
                } while (DateTime.UtcNow < deadline);
                Save(name + "-status.json", state);
                Assert.AreEqual(false, state["Pending"], "Native completion deadline exceeded; no retry or host reset was attempted.");
                var report = VbeBridgeClient.Object(state["Report"]);
                if (Equals(report["uncertain"], false)) Unsettled = false;
                Assert.AreEqual("Completed", state["State"]); Assert.AreEqual(false, state["Stale"]);
                Save(name + "-compact.json", report);
                var human = Data("vba_test_run_status", "Query", query, "Action", "human");
                File.WriteAllText(Path.Combine(root, name + "-human.txt"), (string)human["Report"], Utf8);
                Assert.AreEqual(query, report["run"]);
                return report;
            }
            internal void RestoreOwnedModules()
            {
                if (!baselineVerified) return; // Preflight has never mutated the selected project.
                if (Unsettled) throw new InvalidOperationException("Native outcome is uncertain; all cleanup mutations were refused. The macro and host remain open.");
                Assert.IsFalse(process.HasExited);
                Assert.AreEqual(referencesVersion, Data("list_references")["Version"], "Reference changes refuse cleanup.");
                var current = ReadSources();
                CollectionAssert.AreEquivalent(baseline.Select(module => (string)module["Name"]).Concat(owned.Keys).ToArray(), current.Select(module => (string)module["Name"]).ToArray());
                foreach (var module in baseline)
                    Assert.AreEqual(module["Sha256"], current.Single(item => Equals(item["Name"], module["Name"]))["Sha256"], "An original blank module changed; cleanup refused.");
                foreach (var module in owned)
                    Assert.AreEqual(module.Value, current.Single(item => Equals(item["Name"], module.Key))["Sha256"], "Owned module has concurrent edits; cleanup refused.");
                foreach (var module in owned)
                {
                    var component = Data("component_properties", "Module", module.Key);
                    Assert.AreEqual(module.Value, component["CodeSha256"], "An owned module changed after cleanup inspection; its removal was refused.");
                    Data("remove_component", "Module", module.Key, "ExpectedComponentVersion", component["Version"], "ExpectedProjectVersion", Data("project_properties")["Version"]);
                    if (Unsettled) throw new InvalidOperationException("Component removal is uncertain; further cleanup was refused.");
                }
                var after = ReadSources();
                Save("restored-selected-sources.json", after);
                Assert.AreEqual(baseline.Length, after.Length);
                foreach (var module in baseline)
                    Assert.AreEqual(module["Sha256"], after.Single(item => Equals(item["Name"], module["Name"]))["Sha256"]);
                Assert.AreEqual(referencesVersion, Data("list_references")["Version"]);
                Assert.AreEqual(diskHash, HashFile(path), "The selected disk macro changed outside the fixture; its backup was retained without overwriting it.");
                Save("restoration.json", new { Verified = true, Macro = path, DiskSha256 = diskHash, SavedAutomatically = false,
                    HostRetained = true, MacroLeftOpen = true, OtherProjectsInspected = false });
            }
        }

        private static IDictionary<string, object>[] Catalogue(IDictionary<string, object> catalog)
        {
            var modules = ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object).ToArray();
            var module = modules.Single(item => (string)item["Name"] == ModuleName);
            Assert.IsTrue(string.IsNullOrEmpty(module["Diagnostic"] as string));
            foreach (string fixture in new[] { "ModuleInitialize", "ModuleCleanup", "TestInitialize", "TestCleanup" }) Assert.IsNotNull(module[fixture]);
            var tests = ((object[])module["Tests"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(4, tests.Length); Assert.AreEqual(4, modules.Sum(item => ((object[])item["Tests"]).Length));
            foreach (var test in tests) Assert.IsTrue(string.IsNullOrEmpty(test["Diagnostic"] as string));
            return tests;
        }
        private static void AssertReport(IDictionary<string, object> report, string revision, IDictionary<string, object>[] selected)
        {
            Assert.AreEqual(revision, report["revision"]); Assert.AreEqual(false, report["uncertain"]);
            Assert.IsTrue(string.IsNullOrEmpty(report["error"] as string), Convert.ToString(report["error"]));
            var tests = ((object[])report["tests"]).Select(VbeBridgeClient.Object).ToArray();
            Assert.AreEqual(selected.Length, tests.Length);
            CollectionAssert.AreEquivalent(selected.Select(test => (string)test["Id"]).ToArray(), tests.Select(test => (string)test["id"]).ToArray());
            foreach (var test in selected)
            {
                string procedure = (string)test["Procedure"];
                var result = tests.Single(item => (string)item["id"] == (string)test["Id"]);
                Assert.AreEqual(ModuleName + "." + procedure, result["name"]);
                Assert.AreEqual(procedure == "ABooleanPass" ? "Passed" : procedure == "DRuntimeError" ? "Error" : "Failed", result["outcome"]);
                Assert.AreEqual("Test", result["phase"]);
                if (procedure == "DRuntimeError") Assert.AreEqual(5, Convert.ToInt32(result["error"]));
                if (procedure == "CSwallowedAssertion") StringAssert.Contains((string)result["message"], "Deliberate swallowed assertion");
            }
        }
        private static VbaTestCatalog Review(IDictionary<string, object> catalog)
        {
            var review = new VbaTestCatalog { Project = new VbaTestProjectSnapshot() };
            foreach (var module in ((object[])catalog["Modules"]).Select(VbeBridgeClient.Object))
                review.Modules.Add(new VbaTestModule { Name = (string)module["Name"], Diagnostic = module["Diagnostic"] as string,
                    Tests = ((object[])module["Tests"]).Select(Descriptor).ToList(), ModuleInitialize = Descriptor(module["ModuleInitialize"]),
                    ModuleCleanup = Descriptor(module["ModuleCleanup"]), TestInitialize = Descriptor(module["TestInitialize"]), TestCleanup = Descriptor(module["TestCleanup"]) });
            return review;
        }
        private static VbaTestDescriptor Descriptor(object value)
        {
            if (value == null) return null;
            var item = VbeBridgeClient.Object(value);
            return new VbaTestDescriptor { Id = (string)item["Id"], Module = (string)item["Module"], Procedure = (string)item["Procedure"],
                Kind = (string)item["Kind"], Line = Convert.ToInt32(item["Line"]), Diagnostic = item["Diagnostic"] as string, IgnoreReason = item["IgnoreReason"] as string };
        }
        private static string Canonical(string source) => source.Replace("\r\n", "\n").TrimEnd('\n');
        private static string HashFile(string path) { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", ""); }
        private static object RequireCallback(string loaded)
        {
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var server = classes.OpenSubKey(@"CLSID\" + RuntimeClsid + @"\InprocServer32"))
            using (var progId = classes.OpenSubKey(@"VBAi.TestRuntime\CLSID"))
            {
                Assert.IsNotNull(server); Assert.IsNotNull(progId);
                Assert.AreEqual(RuntimeClsid, Convert.ToString(progId.GetValue("")), true);
                Uri uri; Assert.IsTrue(Uri.TryCreate(Convert.ToString(server.GetValue("CodeBase")), UriKind.Absolute, out uri) && uri.IsFile);
                Assert.AreEqual(Path.GetFullPath(loaded), Path.GetFullPath(uri.LocalPath), true);
                Assert.AreEqual(HashFile(loaded), HashFile(uri.LocalPath));
                Assert.AreEqual(typeof(VbaTestRuntime).FullName, server.GetValue("Class")); Assert.AreEqual(typeof(VbaTestRuntime).Assembly.FullName, server.GetValue("Assembly"));
                Assert.AreEqual("v4.0.30319", server.GetValue("RuntimeVersion")); Assert.AreEqual("Both", Convert.ToString(server.GetValue("ThreadingModel")), true);
                Assert.AreEqual("mscoree.dll", Convert.ToString(server.GetValue("")), true);
                return new { CodeBase = server.GetValue("CodeBase"), Clsid = RuntimeClsid, ProgId = "VBAi.TestRuntime", LoadedAssembly = loaded, Sha256 = HashFile(loaded) };
            }
        }
        private static string Git(string arguments)
        {
            using (var process = Process.Start(new ProcessStartInfo("git", arguments) { WorkingDirectory = Environment.CurrentDirectory,
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true }))
            {
                var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                Assert.IsTrue(process.WaitForExit(10000), "Source revision unavailable; no native mutation was started.");
                Assert.AreEqual(0, process.ExitCode, error.GetAwaiter().GetResult()); return output.GetAwaiter().GetResult();
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
    VBAiTestSupport.AreEqual mInitialized, mCleanup, ""Previous cleanup must have completed.""
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
    VBAiTestSupport.IsFalse mReady, ""Cleanup must reset the fixture.""
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
    Err.Raise 5, ""SolidWorksExplorerFixture"", ""Deliberate runtime error""
End Sub
";
    }
}
