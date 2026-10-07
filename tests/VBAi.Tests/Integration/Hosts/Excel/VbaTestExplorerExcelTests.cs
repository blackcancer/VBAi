using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Integration
{
    /// <summary>Real disposable VBA execution through the compiled service in the external STA test host.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("VbaTestExcel")]
    public sealed class VbaTestExplorerExcelTests
    {
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);

        [STATestMethod]
        public void DisposableWorkbookRunsBooleanAssertionsFixturesAndReportsWithoutChangingTestSource()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1") Assert.Inconclusive("Native Excel tests require VBAi_RUN_EXCEL_TESTS=1.");
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel is not installed.");
            int[] existing;
            using (var inventory = new ProcessInventory()) existing = inventory.Ids;
            object application = null, workbook = null, vbe = null, workbooks = null;
            bool owned = false;
            Process process = null;
            string root = Path.Combine(Path.GetTempPath(), "VBAi-TestExplorer", Guid.NewGuid().ToString("N"));
            var previousContext = SynchronizationContext.Current;
            try
            {
                application = Activator.CreateInstance(excelType);
                dynamic excel = application;
                uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out pid);
                owned = pid != 0 && !existing.Contains((int)pid);
                if (!owned) Assert.Inconclusive("Excel returned an existing process; no document was created.");
                process = Process.GetProcessById((int)pid);
                _ = process.Handle;
                excel.DisplayAlerts = false;
                workbooks = excel.Workbooks;
                workbook = ((dynamic)workbooks).Add();
                vbe = excel.VBE; // Never change Trust Center settings to make the test succeed.
                RunScenario(application, workbook, vbe, (int)pid, root);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previousContext);
                if (owned)
                {
                    if (workbook != null) ((dynamic)workbook).Close(false);
                    if (application != null) ((dynamic)application).Quit();
                }
                foreach (object item in new[] { workbook, workbooks, vbe, application }) if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                if (process != null)
                {
                    try
                    {
                        Assert.IsTrue(process.WaitForExit(15000), "Owned Excel did not shut down normally; it was not terminated automatically.");
                        Assert.AreEqual(0, process.ExitCode, "Owned Excel exited abnormally.");
                    }
                    finally { process.Dispose(); }
                }
                // Retain the disposable workbook and logs when diagnosis is needed.
                Console.WriteLine("Disposable fixture=" + root);
            }
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void RunScenario(object application, object workbook, object vbe, int pid, string root)
        {
            dynamic excel = application;
            dynamic project = ((dynamic)workbook).VBProject;
            dynamic tests = project.VBComponents.Add(1);
            tests.Name = "ExplorerTests";
            tests.CodeModule.AddFromString(Source);
            dynamic production = project.VBComponents.Add(1);
            production.Name = "Production";
            production.CodeModule.AddFromString("Option Explicit\r\nPublic Function Calc(ByVal inputValue As Long) As Long\r\n    Calc = inputValue * 2\r\nEnd Function\r\nPublic Function Unused() As Long\r\n    Unused = 99\r\nEnd Function\r\n");
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "explorer.xlsm");
            ((dynamic)workbook).SaveAs(path, 52);
            using (var dispatcher = new Control())
            {
                _ = dispatcher.Handle;
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                using (var service = new VbeTestExplorerService(vbe, dispatcher))
                {
                    service.IsExecutionHost = () => true;
                    service.Host = new FixtureHost(application, workbook, (int)pid);
                    service.CoverageRoot = () => Path.Combine(root, "coverage-runs");
                    service.CreateCoverageClone = (sourceProject, sourcePath, folder) =>
                    {
                        Assert.IsTrue(VbeDebug.NativeProcedureValuesHost.SameComIdentity(sourceProject, (object)project));
                        Directory.CreateDirectory(folder);
                        string copyPath = Path.Combine(folder, "coverage.xlsm");
                        ((dynamic)workbook).SaveCopyAs(copyPath);
                        bool events = (bool)excel.EnableEvents;
                        object copied = null;
                        try { excel.EnableEvents = false; copied = excel.Workbooks.Open(copyPath, 0, false); }
                        finally { excel.EnableEvents = events; }
                        return new VbaTestCoverageClone
                        {
                            Project = ((dynamic)copied).VBProject,
                            Path = copyPath,
                            Close = () => { bool previous = (bool)excel.EnableEvents; try { excel.EnableEvents = false; ((dynamic)copied).Close(false); } finally { excel.EnableEvents = previous; } }
                        };
                    };
                    service.ConfirmSupport = (review, before, after) => true;
                    var catalog = service.DiscoverSelector(path);
                    Assert.AreEqual(6, catalog.Tests.Count());
                    string original = (string)tests.CodeModule.Lines[1, (int)tests.CodeModule.CountOfLines];
                    service.InstallSupport(catalog);
                    catalog = service.DiscoverSelector(path);
                    Assert.IsNull(service.ExecutionUnavailableReason(catalog));
                    var run = Pump(service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                    Assert.IsNull(run.Error, run.Error);
                    Assert.IsFalse(run.OutcomeUnknown);
                    Assert.AreEqual(VbaTestOutcome.Passed, Result(run, "BooleanPass").Outcome);
                    Assert.AreEqual(VbaTestOutcome.Failed, Result(run, "BooleanFail").Outcome);
                    Assert.AreEqual(VbaTestOutcome.Failed, Result(run, "AssertionFail").Outcome);
                    Assert.AreEqual(VbaTestOutcome.Error, Result(run, "UnexpectedError").Outcome);
                    Assert.AreEqual(5, Result(run, "UnexpectedError").ErrorNumber);
                    Assert.AreEqual(VbaTestOutcome.Failed, Result(run, "SwallowedAssertion").Outcome);
                    Assert.AreEqual(VbaTestOutcome.Inconclusive, Result(run, "Placeholder").Outcome);
                    StringAssert.Contains(VbaTestReports.Human(run), "AssertionFail");
                    StringAssert.Contains(VbaTestReports.Compact(run), "\"available\":false");
                    string reports = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-explorer-native");
                    Directory.CreateDirectory(reports);
                    File.WriteAllText(Path.Combine(reports, "results-human.txt"), VbaTestReports.Human(run), new System.Text.UTF8Encoding(false));
                    File.WriteAllText(Path.Combine(reports, "results-llm.json"), VbaTestReports.Compact(run), new System.Text.UTF8Encoding(false));
                    Assert.AreEqual(original, (string)tests.CodeModule.Lines[1, (int)tests.CodeModule.CountOfLines]);
                    // Every completed test invokes cleanup, including assertions and VBA errors.
                    Assert.AreEqual(6, Convert.ToInt32(excel.Run("'" + path + "'!ExplorerTests.CleanupCount")));
                    Assert.AreEqual(1, Convert.ToInt32(excel.Run("'" + path + "'!ExplorerTests.ModuleCleanupCount")));
                    string originalProduction = (string)production.CodeModule.Lines[1, (int)production.CodeModule.CountOfLines];
                    var covered = Pump(service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                    Assert.IsNull(covered.Error, covered.Error);
                    Assert.IsFalse(covered.OutcomeUnknown);
                    Assert.IsTrue(covered.Coverage.Available);
                    Assert.IsTrue(covered.Coverage.Complete);
                    Assert.AreEqual(2, covered.Coverage.Eligible);
                    Assert.AreEqual(1, covered.Coverage.Hit);
                    Assert.AreEqual(50d, covered.Coverage.Percent.Value);
                    Assert.AreEqual(originalProduction, (string)production.CodeModule.Lines[1, (int)production.CodeModule.CountOfLines]);
                    Assert.AreEqual(original, (string)tests.CodeModule.Lines[1, (int)tests.CodeModule.CountOfLines]);
                    Assert.AreEqual(1, (int)excel.Workbooks.Count, "The owned coverage copy must be closed after verified collection.");
                    File.WriteAllText(Path.Combine(reports, "coverage-human.txt"), VbaTestReports.Human(covered));
                    File.WriteAllText(Path.Combine(reports, "coverage-llm.json"), VbaTestReports.Compact(covered));
                    tests.CodeModule.AddFromString("' source changed after run");
                    Assert.ThrowsException<InvalidOperationException>(() => service.RunAsync(catalog, catalog.Tests.Take(1).ToArray(), null, CancellationToken.None));
                    Console.WriteLine("Assembly=" + typeof(VbeTestExplorerService).Assembly.Location + "; MVID=" + typeof(VbeTestExplorerService).Module.ModuleVersionId);
                    Console.WriteLine("Excel=" + excel.Version + "; PID=" + pid + "; execution boundary=external STA fixture adapter; installed add-in not qualified.");
                }
            }
        }

        private static VbaTestResult Result(VbaTestRun run, string procedure) => run.Results.Single(result => result.Test.Procedure == procedure);

        private static T Pump<T>(Task<T> task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(30);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(10); }
            Assert.IsTrue(task.IsCompleted, "The call did not complete. No retry or host reset was attempted.");
            return task.GetAwaiter().GetResult();
        }

        private sealed class ProcessInventory : IDisposable
        {
            private readonly Process[] processes = Process.GetProcessesByName("EXCEL");
            internal int[] Ids => processes.Select(process => process.Id).ToArray();
            public void Dispose() { foreach (var process in processes) process.Dispose(); }
        }

        private sealed class FixtureHost : VbeDebug.IProcedureValuesHost
        {
            private readonly object excel, workbook;
            private readonly int pid;
            internal FixtureHost(object excel, object workbook, int pid) { this.excel = excel; this.workbook = workbook; this.pid = pid; }
            public object ResolveTarget(object project, string expectedHostPath)
            {
                uint actual; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)excel).Hwnd)), out actual);
                Assert.AreEqual(pid, (int)actual);
                foreach (object opened in ((dynamic)excel).Workbooks)
                    if (string.Equals(expectedHostPath, (string)((dynamic)opened).FullName, StringComparison.OrdinalIgnoreCase) &&
                        VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, (object)((dynamic)opened).VBProject)) return opened;
                throw new InvalidOperationException("The exact owned workbook was not resolved.");
            }
            public object Invoke(object target, string module, string procedure, object[] arguments)
            {
                var parameters = new object[arguments.Length + 1];
                parameters[0] = "'" + ((string)((dynamic)target).FullName).Replace("'", "''") + "'!" + module + "." + procedure;
                Array.Copy(arguments, 0, parameters, 1, arguments.Length);
                return excel.GetType().InvokeMember("Run", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, excel, parameters, System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        private const string Source = @"Option Explicit
'@TestModule
Private mCleanup As Long
Private mModuleCleanup As Long
Private mReady As Boolean
'@ModuleInitialize
Public Sub InitializeModule()
    mCleanup = 0
    mModuleCleanup = 0
End Sub
'@ModuleCleanup
Public Sub CleanupModule()
    mModuleCleanup = mModuleCleanup + 1
End Sub
'@TestInitialize
Public Sub InitializeTest()
    mReady = True
End Sub
'@TestCleanup
Public Sub CleanupTest()
    mCleanup = mCleanup + 1
    mReady = False
End Sub
'@TestMethod
Public Function BooleanPass() As Boolean
    BooleanPass = mReady And (Production.Calc(2) = 4)
End Function
'@TestMethod
Public Function BooleanFail() As Boolean
    BooleanFail = False
End Function
'@TestMethod
Public Sub AssertionFail()
    VBAiTestSupport.AreEqual CLng(4), CLng(3), ""Expected four.""
End Sub
'@TestMethod
Public Sub UnexpectedError()
    Err.Raise 5, ""ExplorerFixture"", ""Deliberate runtime error.""
End Sub
'@TestMethod
Public Sub SwallowedAssertion()
    On Error Resume Next
    VBAiTestSupport.Fail ""Swallowed failure must stay failed.""
End Sub
'@TestMethod
Public Sub Placeholder()
    VBAiTestSupport.Inconclusive ""Placeholder.""
End Sub
Public Function CleanupCount() As Long
    CleanupCount = mCleanup
End Function
Public Function ModuleCleanupCount() As Long
    ModuleCleanupCount = mModuleCleanup
End Function";
    }
}
