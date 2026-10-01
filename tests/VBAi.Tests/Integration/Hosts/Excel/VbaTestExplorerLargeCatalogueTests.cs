using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in native compilation of the complete partitioned dispatcher, without running the full catalogue.</summary>
    [TestClass, TestCategory("VbaTestExcelLarge"), DoNotParallelize]
    public sealed class VbaTestExplorerLargeCatalogueTests
    {
        private const int TextLimit = 1024 * 1024;
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

        [STATestMethod, Timeout(240000)]
        public void TenThousandShortTestsCompileAndBoundarySelectionsReturnVerifiedResults()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_LARGE_TESTS") != "1")
                Assert.Inconclusive("Native large-catalogue qualification requires VBAi_RUN_EXCEL_LARGE_TESTS=1.");
            Type type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null) Assert.Inconclusive("Excel is not installed.");
            string sourceRevision = Git("rev-parse HEAD").Trim();
            string sourceStatus = Git("status --porcelain");
            int[] existing;
            var inventory = Process.GetProcessesByName("EXCEL");
            try { existing = inventory.Select(item => item.Id).ToArray(); }
            finally { foreach (var item in inventory) item.Dispose(); }
            object application = null, books = null, workbook = null, vbe = null;
            Process process = null;
            bool owned = false, uncertain = false;
            string root = Path.Combine(Path.GetTempPath(), "VBAi-TestExplorer-Large", Guid.NewGuid().ToString("N"));
            var previous = SynchronizationContext.Current;
            Directory.CreateDirectory(root);
            try
            {
                application = Activator.CreateInstance(type);
                uint pid;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out pid);
                owned = pid != 0 && !existing.Contains((int)pid);
                if (!owned) Assert.Inconclusive("Excel reused an existing PID; no workbook was created.");
                process = Process.GetProcessById((int)pid);
                _ = process.Handle;
                ((dynamic)application).DisplayAlerts = false;
                ((dynamic)application).AutomationSecurity = 3; // Never lower macro security or change Trust Center.
                books = ((dynamic)application).Workbooks;
                if ((int)((dynamic)books).Count != 0)
                {
                    uncertain = true;
                    Assert.Inconclusive("The new Excel process contains unexpected startup workbooks; none was changed and the process was retained.");
                }
                workbook = ((dynamic)books).Add();
                vbe = ((dynamic)application).VBE;
                RunScenario(application, workbook, vbe, (int)pid, root, sourceRevision, sourceStatus, value => uncertain = value);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
                try
                {
                    if (owned && !uncertain)
                    {
                        uint cleanupPid;
                        GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out cleanupPid);
                        Assert.AreEqual(process.Id, (int)cleanupPid, "Cleanup refused: Excel window ownership changed.");
                        if (workbook != null) ((dynamic)workbook).Close(false);
                        if (books != null && (int)((dynamic)books).Count != 0)
                        {
                            uncertain = true;
                            Assert.Fail("Unexpected workbooks remain after closing the owned fixture; Excel was retained without Quit.");
                        }
                        if (application != null) ((dynamic)application).Quit();
                    }
                }
                finally
                {
                    foreach (object item in new[] { workbook, books, vbe, application })
                        if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
                    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                    if (process != null)
                    {
                        try
                        {
                            if (!uncertain)
                            {
                                Assert.IsTrue(process.WaitForExit(15000), "Owned Excel did not exit normally; no automatic termination was attempted.");
                                Assert.AreEqual(0, process.ExitCode, "Owned Excel exited abnormally.");
                            }
                            else Console.WriteLine("Uncertain execution: owned PID retained without Close/Quit=" + process.Id);
                        }
                        finally { process.Dispose(); }
                    }
                    Console.WriteLine("Large catalogue evidence=" + root);
                }
            }
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void RunScenario(object application, object workbook, object vbe, int pid, string root, string sourceRevision, string sourceStatus, Action<bool> uncertain)
        {
            object project = ((dynamic)workbook).VBProject;
            object components = ((dynamic)project).VBComponents;
            try
            {
                AddModule(components, "BulkState", StateSource);
                for (int module = 0; module < 100; module++)
                {
                    var source = new StringBuilder("Option Explicit\r\n'@TestModule\r\n");
                    for (int test = 0; test < 100; test++)
                        source.Append("'@TestMethod\r\nPublic Sub T" + test.ToString("D3")
                            + "()\r\n    BulkState.Seen = BulkState.Seen + 1\r\nEnd Sub\r\n");
                    AddModule(components, "L" + module.ToString("D3"), source.ToString());
                }
                AddModule(components, "ZFailures", FailureSource);
                string path = Path.Combine(root, "large-catalogue.xlsm");
                ((dynamic)workbook).SaveAs(path, 52);
                using (var dispatcher = new Control())
                {
                    _ = dispatcher.Handle;
                    SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                    using (var service = new VbeTestExplorerService(vbe, dispatcher))
                    {
                        var host = new FixtureHost(application, workbook, pid);
                        service.IsExecutionHost = () => true;
                        service.Host = host;
                        service.BackupRoot = () => Path.Combine(root, "support-backups");
                        var before = service.DiscoverSelector(path);
                        Assert.AreEqual(10004, before.Tests.Count());
                        Assert.AreEqual(100, before.Modules.Count(module => module.Name.StartsWith("L", StringComparison.Ordinal)));
                        Assert.AreEqual(10000, before.Tests.Count(test => test.Module.StartsWith("L", StringComparison.Ordinal)));
                        Assert.AreEqual(10004, before.Tests.Select(test => test.Id).Distinct(StringComparer.Ordinal).Count());
                        Assert.IsFalse(before.Tests.Any(test => !string.IsNullOrEmpty(test.Diagnostic)));
                        string generated = VbaTestRuntimeSource.Generate(before);
                        Assert.AreEqual("3", VbaTestRuntimeSource.Version);
                        StringAssert.Contains(generated, "' VBAi test support version 3");
                        Assert.IsTrue(generated.Length <= TextLimit, "This shape exceeds the existing 1 MiB support-text boundary; native installation must be refused.");
                        int leafCount = Regex.Matches(generated, @"(?m)^Private Function VBAiDispatchLeaf").Count;
                        int routeCount = Regex.Matches(generated, @"(?m)^Private Function VBAiDispatchRoute").Count;
                        Assert.IsTrue(leafCount > 32, "The fixture must cross a complete route fanout.");
                        Assert.IsTrue(routeCount > 1);
                        WriteEvidence(root, "before-install", new { PID = pid, ExcelVersion = Convert.ToString(((dynamic)application).Version),
                            Assembly = typeof(VbeTestExplorerService).Assembly.Location, MVID = typeof(VbeTestExplorerService).Module.ModuleVersionId,
                            CurrentSourceRevision = sourceRevision, SourceStatus = sourceStatus,
                            ProjectRevision = before.Project.Revision, ReferencesHash = before.Project.ReferencesHash,
                            SupportVersion = VbaTestRuntimeSource.Version, SupportSignature = VbaTestRuntimeSource.DispatchSignature(before),
                            SupportHash = VbeTestExplorerService.Hash(generated), SupportCharacters = generated.Length,
                            TextLimit, LeafCount = leafCount, RouteCount = routeCount, TestCount = before.Tests.Count(),
                            Modules = before.Project.Modules.Select(module => new { module.Name, module.Hash, module.ComponentType }).ToArray() });
                        File.WriteAllText(Path.Combine(root, "generated-support.bas"), generated, new UTF8Encoding(false));
                        service.ConfirmSupport = (review, oldSource, newSource) => {
                            Assert.AreEqual(Canonical(generated), Canonical(newSource));
                            Assert.IsTrue(newSource.Length <= TextLimit);
                            return true;
                        };
                        service.InstallSupport(before);
                        var installed = service.DiscoverSelector(path);
                        Assert.AreEqual(10004, installed.Tests.Count());
                        Assert.AreEqual(before.Project.Modules.Length + 1, installed.Project.Modules.Length);
                        CollectionAssert.AreEquivalent(before.Tests.Select(test => test.Id).ToArray(), installed.Tests.Select(test => test.Id).ToArray());
                        AssertSources(before, installed);
                        Assert.AreEqual(Canonical(generated), Canonical(installed.Project.Modules.Single(module => module.Name == VbaTestRuntimeSource.ModuleName).Source));
                        Assert.IsNull(service.ExecutionUnavailableReason(installed));
                        // Unknown-key traversal visits every leaf without executing any catalogue entry.
                        // Excel.Run must return the VBA verdict; a command ACK cannot establish compilation.
                        object verdict = InvokeVerified(host, workbook, "CompileVerdict", uncertain);
                        Assert.IsTrue(verdict is bool compiled && compiled, "The complete VBA dispatcher did not return its verified native compilation verdict.");
                        Assert.AreEqual(0, Convert.ToInt32(InvokeVerified(host, workbook, "ExecutedCount", uncertain)));
                        var bulk = installed.Tests.Where(test => test.Module.StartsWith("L", StringComparison.Ordinal))
                            .OrderBy(test => test.Module, StringComparer.OrdinalIgnoreCase).ThenBy(test => test.Procedure, StringComparer.OrdinalIgnoreCase).ToArray();
                        int[] positions = { 0, 127, 128, 4095, 4096, 8191, 8192, 9999 };
                        var selected = positions.Select(position => bulk[position]).Concat(installed.Tests.Where(test => test.Module == "ZFailures")).ToArray();
                        Assert.AreEqual(12, selected.Length);
                        uncertain(true);
                        var run = Pump(service.RunAsync(installed, selected, null, CancellationToken.None), () => uncertain(true));
                        uncertain(run.OutcomeUnknown);
                        File.WriteAllText(Path.Combine(root, "results-human.txt"), VbaTestReports.Human(run), new UTF8Encoding(false));
                        File.WriteAllText(Path.Combine(root, "results-compact.json"), VbaTestReports.Compact(run), new UTF8Encoding(false));
                        Assert.IsNull(run.Error, run.Error);
                        Assert.IsFalse(run.OutcomeUnknown);
                        Assert.AreEqual(12, run.Results.Count);
                        Assert.AreEqual(9, run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed));
                        Assert.AreEqual(2, run.Results.Count(result => result.Outcome == VbaTestOutcome.Failed));
                        Assert.AreEqual(1, run.Results.Count(result => result.Outcome == VbaTestOutcome.Error));
                        CollectionAssert.AreEquivalent(selected.Select(test => test.Id).ToArray(), run.Results.Select(result => result.Test.Id).ToArray());
                        foreach (var test in positions.Select(position => bulk[position]))
                            Assert.AreEqual(VbaTestOutcome.Passed, run.Results.Single(result => result.Test.Id == test.Id).Outcome);
                        Assert.AreEqual(VbaTestOutcome.Passed, Outcome(run, "BooleanPass").Outcome);
                        Assert.AreEqual(VbaTestOutcome.Failed, Outcome(run, "BooleanFail").Outcome);
                        Assert.AreEqual(VbaTestOutcome.Failed, Outcome(run, "SwallowedAssertion").Outcome);
                        StringAssert.Contains(Outcome(run, "SwallowedAssertion").Message, "Sticky failure");
                        Assert.AreEqual(VbaTestOutcome.Error, Outcome(run, "NativeError").Outcome);
                        Assert.AreEqual(5, Outcome(run, "NativeError").ErrorNumber);
                        Assert.AreEqual(8, Convert.ToInt32(InvokeVerified(host, workbook, "ExecutedCount", uncertain)), "Only selected bulk entries may execute.");
                        var after = service.DiscoverSelector(path);
                        AssertSources(before, after);
                        Assert.AreEqual(installed.Project.Revision, after.Project.Revision);
                        Assert.AreEqual(10004, after.Tests.Count());
                        WriteEvidence(root, "verified", new { PID = pid, MVID = typeof(VbeTestExplorerService).Module.ModuleVersionId,
                            BeforeRevision = before.Project.Revision, InstalledRevision = installed.Project.Revision, FinalRevision = after.Project.Revision,
                            NativeCompileVerdict = verdict, SelectedIds = selected.Select(test => test.Id).ToArray(),
                            Positions = positions, Passed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed),
                            Failed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Failed),
                            Errors = run.Results.Count(result => result.Outcome == VbaTestOutcome.Error),
                            Boundary = "external STA adapter; installed add-in loading is not qualified" });
                    }
                }
            }
            finally
            {
                foreach (object item in new[] { components, project })
                    if (Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
            }
        }

        private static void AddModule(object components, string name, string source)
        {
            object component = ((dynamic)components).Add(1);
            object code = null;
            try { ((dynamic)component).Name = name; code = ((dynamic)component).CodeModule; ((dynamic)code).AddFromString(source); }
            finally { foreach (object item in new[] { code, component }) if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item); }
        }

        private static string Canonical(string source) => (source ?? "").Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n') + "\n";
        private static object InvokeVerified(FixtureHost host, object workbook, string procedure, Action<bool> uncertain)
        {
            bool returned = false;
            uncertain(true);
            try
            {
                object value = host.Invoke(workbook, "BulkState", procedure, new object[0]);
                returned = true;
                return value;
            }
            finally
            {
                // A malformed returned value is a known failure; an interrupted native call must retain the host.
                if (returned) uncertain(false);
            }
        }
        private static string Git(string arguments)
        {
            var start = new ProcessStartInfo("git", arguments) { WorkingDirectory = Environment.CurrentDirectory,
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using (var process = Process.Start(start))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var error = process.StandardError.ReadToEndAsync();
                Assert.IsTrue(process.WaitForExit(10000), "Source identity query timed out; qualification was not started.");
                Assert.AreEqual(0, process.ExitCode, "Source identity unavailable: " + error.GetAwaiter().GetResult());
                return output.GetAwaiter().GetResult();
            }
        }
        private static VbaTestResult Outcome(VbaTestRun run, string procedure) => run.Results.Single(result => result.Test.Module == "ZFailures" && result.Test.Procedure == procedure);
        private static void AssertSources(VbaTestCatalog expected, VbaTestCatalog actual)
        {
            Assert.AreEqual(expected.Project.ReferencesHash, actual.Project.ReferencesHash);
            foreach (var module in expected.Project.Modules)
            {
                var observed = actual.Project.Modules.Single(item => item.Name == module.Name);
                Assert.AreEqual(module.Hash, observed.Hash, module.Name);
                Assert.AreEqual(module.ComponentType, observed.ComponentType, module.Name);
            }
        }
        private static void WriteEvidence(string root, string name, object evidence) => File.WriteAllText(Path.Combine(root, name + ".json"),
            new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(evidence), new UTF8Encoding(false));
        private static T Pump<T>(Task<T> task, Action uncertain)
        {
            DateTime deadline = DateTime.UtcNow.AddSeconds(60);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(10); }
            if (!task.IsCompleted) { uncertain(); Assert.Fail("Execution exceeded its deadline; no retry, Close, Quit or process termination was attempted."); }
            return task.GetAwaiter().GetResult();
        }

        private sealed class FixtureHost : VbeDebug.IProcedureValuesHost
        {
            private readonly object excel, workbook;
            private readonly int pid;
            internal FixtureHost(object excel, object workbook, int pid) { this.excel = excel; this.workbook = workbook; this.pid = pid; }
            public object ResolveTarget(object project, string expectedHostPath)
            {
                uint actual;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)excel).Hwnd)), out actual);
                Assert.AreEqual(pid, (int)actual);
                Assert.IsTrue(VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, (object)((dynamic)workbook).VBProject));
                Assert.AreEqual(Path.GetFullPath(expectedHostPath), Path.GetFullPath((string)((dynamic)workbook).FullName), true);
                return workbook;
            }
            public object Invoke(object target, string module, string procedure, object[] arguments)
            {
                Assert.IsTrue(VbeDebug.NativeProcedureValuesHost.SameComIdentity(target, workbook));
                uint actual;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)excel).Hwnd)), out actual);
                Assert.AreEqual(pid, (int)actual);
                var parameters = new object[arguments.Length + 1];
                parameters[0] = "'" + ((string)((dynamic)workbook).FullName).Replace("'", "''") + "'!" + module + "." + procedure;
                Array.Copy(arguments, 0, parameters, 1, arguments.Length);
                return excel.GetType().InvokeMember("Run", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, excel, parameters, System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        private const string StateSource = @"Option Explicit
Public Seen As Long
Public Function ExecutedCount() As Long
    ExecutedCount = Seen
End Function
Public Function CompileVerdict() As Boolean
    Dim result As Variant
    result = VBAiTestSupport.VBAiExecuteTest(""UnknownModule"", ""UnknownProcedure"")
    CompileVerdict = (CStr(result(0)) = ""Error"") And (InStr(1, CStr(result(1)), ""absent from the installed test dispatch table"", vbBinaryCompare) > 0) And (Seen = 0)
End Function
";
        private const string FailureSource = @"Option Explicit
'@TestModule
'@TestMethod
Public Function BooleanPass() As Boolean
    BooleanPass = True
End Function
'@TestMethod
Public Function BooleanFail() As Boolean
    BooleanFail = False
End Function
'@TestMethod
Public Sub SwallowedAssertion()
    On Error Resume Next
    VBAiTestSupport.Fail ""Sticky failure""
End Sub
'@TestMethod
Public Sub NativeError()
    Err.Raise 5, ""LargeCatalogueFixture"", ""Deliberate native error""
End Sub
";
    }
}
