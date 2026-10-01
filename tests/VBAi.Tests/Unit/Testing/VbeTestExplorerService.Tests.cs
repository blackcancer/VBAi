using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed partial class VbeTestExplorerServiceTests
    {
        [STATestMethod]
        public void ProjectIdentityIsStableAndAmbiguousNamesRequireExactPaths()
        {
            using (var fixture = new Fixture())
            {
                var second = Project("WorkbookProject", @"C:\Temp\Second.xlsm");
                fixture.Vbe.VBProjects.Add(second);
                var first = fixture.Service.ReadProjects();
                var repeated = fixture.Service.ReadProjects();
                Assert.AreEqual(first[0].Id, repeated[0].Id);
                Assert.AreEqual(first[1].Id, repeated[1].Id);
                Assert.AreNotEqual(first[0].Id, first[1].Id);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.DiscoverSelector("WorkbookProject"));
                Assert.AreEqual(first[1].Id, fixture.Service.DiscoverSelector(second.FileName).Project.Id);
                var old = fixture.Catalog();
                fixture.Vbe.VBProjects[0] = Project(fixture.Project.Name, fixture.Project.FileName);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Validate(old));
                Assert.AreNotEqual(old.Project.Id, fixture.Service.ReadProjects()[0].Id);
            }
        }

        [STATestMethod]
        public void ValidationFingerprintsCalleesReferencesPathModeAndProtection()
        {
            foreach (string change in new[] { "callee", "reference", "path", "mode", "protection" })
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                if (change == "callee") fixture.Project.VBComponents[0].CodeModule.Source += "\n' production edit";
                if (change == "reference") fixture.Project.References[0].Minor++;
                if (change == "path") fixture.Project.FileName = @"C:\Temp\Moved.xlsm";
                if (change == "mode") fixture.Project.Mode = 1;
                if (change == "protection") fixture.Project.Protection = 1;
                Exception refusal = null;
                try { fixture.Service.Validate(catalog); }
                catch (Exception error) { refusal = error; }
                Assert.IsNotNull(refusal, change);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void DiscoverySurvivesProtectedProjectListingAndExecutionExplainsHostSupport()
        {
            using (var fixture = new Fixture())
            {
                fixture.Project.Protection = 1;
                var row = fixture.Service.ReadProjects().Single();
                Assert.AreEqual(fixture.Project.Name, row.Name);
                Assert.AreEqual(0, row.Modules.Length);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Discover(row.Id));
                fixture.Project.Protection = 0;
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                fixture.Service.IsExecutionHost = () => false;
                StringAssert.Contains(fixture.Service.ExecutionUnavailableReason(catalog), "in-process VBE host");
                fixture.Service.IsExecutionHost = () => true;
                fixture.Project.FileName = null;
                StringAssert.Contains(fixture.Service.ExecutionUnavailableReason(fixture.Catalog()), "Save");
                fixture.Project.FileName = @"C:\Temp\Fixture.xlsm";
                fixture.Project.VBComponents.Single(component => component.Name == VbaTestRuntimeSource.ModuleName).CodeModule.Source += "\n' edited support";
                StringAssert.Contains(fixture.Service.ExecutionUnavailableReason(fixture.Catalog()), "outdated or changed");
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void QueuedRunRevalidatesSourceModeAndLiveIdentityBeforeNativeInvocation()
        {
            foreach (string change in new[] { "source", "mode", "replacement" })
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                if (change == "source") fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed after scheduling";
                if (change == "mode") fixture.Project.Mode = 1;
                if (change == "replacement") fixture.Vbe.VBProjects[0] = Project(fixture.Project.Name, fixture.Project.FileName);
                var run = Pump(task);
                Assert.AreEqual(0, fixture.Host.Invocations, change);
                Assert.IsFalse(run.OutcomeUnknown, change);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked), change);
            }
        }

        [STATestMethod]
        public void FinalTargetResolutionCannotBypassRevisionOrPermissionRevalidation()
        {
            foreach (bool changesPermission in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                bool allowed = true;
                fixture.Host.Resolving = count =>
                {
                    if (count < 3) return;
                    if (changesPermission) allowed = false;
                    else fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed during target resolution";
                };
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray(), () => { if (!allowed) throw new InvalidOperationException("Permission withdrawn"); });
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, (string)start.Query, "compact")).Pending));
                Assert.AreEqual(0, fixture.Host.Invocations, changesPermission ? "Permission changed" : "Revision changed");
            }
        }

        [STATestMethod]
        public void InvalidNativeReturnAndTransportFailureLatchUnknownAndNeverRetry()
        {
            foreach (bool throws in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                fixture.Host.Returned = new object();
                fixture.Host.ThrowOnInvoke = throws;
                var run = Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.AreEqual(VbaTestOutcome.OutcomeUnknown, run.Results[0].Outcome);
                Assert.AreEqual(1, fixture.Host.Invocations);
                StringAssert.Contains(fixture.Service.ExecutionUnavailableReason(catalog), "uncertain");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                for (int i = 0; i < 3; i++)
                {
                    dynamic status = fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "compact");
                    Assert.AreEqual("OutcomeUnknown", (string)status.State);
                    Assert.IsFalse((bool)status.Pending);
                }
                Assert.AreEqual(1, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void ProgressReportsExposeUncertaintyBeforeTheRunSettles()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                fixture.Host.Returned = new object();
                var catalog = fixture.Catalog();
                bool observed = false;
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), result =>
                {
                    if (result.Outcome != VbaTestOutcome.OutcomeUnknown) return;
                    var entries = (System.Collections.IDictionary)typeof(VbeTestExplorerService)
                        .GetField("runs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                    string id = entries.Keys.Cast<string>().Single();
                    dynamic status = fixture.Service.RunStatus(fixture.Project.FileName, id, "compact");
                    Assert.IsTrue((bool)status.Pending, "Progress must expose uncertainty before final settlement.");
                    var report = (Dictionary<string, object>)status.Report;
                    Assert.IsTrue((bool)report["uncertain"]);
                    Assert.AreEqual(1, ((Dictionary<string, object>)report["counts"])["OutcomeUnknown"]);
                    dynamic human = fixture.Service.RunStatus(fixture.Project.FileName, id, "human");
                    StringAssert.Contains((string)human.Report, UiText.Get("Execution outcome is uncertain. Inspect the host before any further run."));
                    observed = true;
                }, CancellationToken.None);
                var run = Pump(task);
                Assert.IsTrue(observed);
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.AreEqual(1, fixture.Host.Invocations, "A status read must never repeat native execution.");
            }
        }

        [STATestMethod]
        public void SharedUiAndCommandRunsExcludeOverlapAndExposeReadOnlyHistoricalReports()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StartRun(fixture.Project.FileName,
                    catalog.Project.Revision, catalog.Tests.Select(test => test.Id).ToArray()));
                var run = Pump(task);
                Assert.AreEqual(2, fixture.Host.Invocations);
                dynamic compact = fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "compact");
                Assert.AreEqual("Completed", (string)compact.State);
                Assert.IsFalse((bool)compact.Pending);
                Assert.IsFalse((bool)compact.Stale);
                var report = (Dictionary<string, object>)compact.Report;
                Assert.AreEqual(fixture.Project.Name, report["project"]);
                dynamic human = fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "human");
                StringAssert.Contains((string)human.Report, fixture.Project.Name);
                fixture.Project.VBComponents[0].CodeModule.Source += "\n' edit after run";
                Assert.IsTrue((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "compact")).Stale);
                var second = Project("OtherProject", @"C:\Temp\Other.xlsm");
                fixture.Vbe.VBProjects.Add(second);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunStatus(second.FileName, run.Id, "compact"));
                Assert.AreEqual(2, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void StopQueuedRunPreventsNativeDispatchAndMarksRemainingTestsCancelled()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray());
                string id = start.Query;
                dynamic stop = fixture.Service.StopRun(fixture.Project.FileName, id);
                Assert.AreEqual("StopRequested", (string)stop.State);
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, id, "compact")).Pending));
                Assert.AreEqual(0, fixture.Host.Invocations);
                dynamic status = fixture.Service.RunStatus(fixture.Project.FileName, id, "compact");
                Assert.AreEqual("Cancelled", (string)status.State);
                var report = (Dictionary<string, object>)status.Report;
                var tests = ((object[])report["tests"]).Cast<Dictionary<string, object>>().ToArray();
                Assert.IsTrue(tests.All(test => (string)test["outcome"] == "Cancelled"));
            }
        }

        [STATestMethod]
        public void DisposalSettlesQueuedRunWithoutNativeInvocationEvenIfDispatcherIsDestroyed()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                int owner = Thread.CurrentThread.ManagedThreadId, releasedOn = 0;
                var continuations = OwnerContinuations(fixture.Service);
                Assert.IsTrue(continuations.IsHandleCreated);
                continuations.Disposed += (_, __) => { releasedOn = Thread.CurrentThread.ManagedThreadId; fixture.Service.Dispose(); };
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), _ =>
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, "Disposal results must be published on the owner."), CancellationToken.None);
                fixture.Service.Dispose();
                Assert.IsFalse(continuations.IsDisposed, "The active run still needs its independent continuation handle.");
                fixture.Dispatcher.Dispose();
                var run = Pump(task);
                Assert.AreEqual(0, fixture.Host.Invocations);
                Assert.IsFalse(run.OutcomeUnknown);
                Assert.IsTrue(continuations.IsDisposed);
                Assert.IsNull(OwnerContinuations(fixture.Service));
                Assert.AreEqual(owner, releasedOn);
                Assert.ThrowsException<ObjectDisposedException>(() => fixture.Service.ReadProjects());
            }
        }

        [STATestMethod]
        public void IdleDisposalReleasesTheIndependentContinuationHandleImmediately()
        {
            using (var fixture = new Fixture())
            {
                var continuations = OwnerContinuations(fixture.Service);
                int releasedOn = 0;
                continuations.Disposed += (_, __) => releasedOn = Thread.CurrentThread.ManagedThreadId;
                fixture.Service.Dispose(); fixture.Dispatcher.Dispose();
                Assert.IsTrue(continuations.IsDisposed);
                Assert.IsNull(OwnerContinuations(fixture.Service));
                Assert.AreEqual(Thread.CurrentThread.ManagedThreadId, releasedOn);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void ACompletedRunKeepsContinuationsAvailableUntilServiceDisposal()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var continuations = OwnerContinuations(fixture.Service);
                var run = Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsNull(run.Error); Assert.IsFalse(continuations.IsDisposed);
                fixture.Service.Dispose();
                Assert.IsTrue(continuations.IsDisposed);
                Assert.IsNull(OwnerContinuations(fixture.Service));
                Assert.AreEqual(2, fixture.Host.Invocations, "Disposal must not dispatch any additional native call.");
            }
        }

        private static Control OwnerContinuations(VbeTestExplorerService service)
            => (Control)typeof(VbeTestExplorerService).GetField("continuationDispatcher", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(service);

        [STATestMethod]
        public void SupportInstallRequiresExactReviewedSourcePreservesBackupAndNeverSavesWorkbook()
        {
            using (var fixture = new Fixture())
            {
                var catalog = fixture.Catalog();
                string generated = VbaTestRuntimeSource.Generate(catalog);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(catalog, generated + "\r\n' unexpected edit"));
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
                Assert.IsFalse(Directory.Exists(fixture.BackupRoot));
                fixture.Service.ConfirmSupport = (plan, before, after) => false;
                fixture.Service.InstallSupport(catalog);
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
                dynamic applied = fixture.Service.ApplySupport(catalog, generated);
                Assert.IsTrue((bool)applied.Applied);
                Assert.IsFalse((bool)applied.SavedToDisk);
                Assert.IsTrue(File.Exists((string)applied.Backup));
                Assert.AreEqual("", File.ReadAllText((string)applied.Backup));
                Assert.AreEqual(1, fixture.Project.VBComponents.Additions);
                Assert.IsTrue(VbaTestRuntimeSource.IsOwned(fixture.Project.VBComponents.Last().CodeModule.Source));
                Assert.IsNull(fixture.Service.ExecutionUnavailableReason(fixture.Catalog()));
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void PartialSupportWriteRetainsOriginalBackupWithoutRetryOrAutomaticRecovery()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var support = fixture.Project.VBComponents.Single(component => component.Name == VbaTestRuntimeSource.ModuleName);
                string original = support.CodeModule.Source;
                support.CodeModule.ThrowOnAdd = true;
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(catalog, VbaTestRuntimeSource.Generate(catalog)));
                StringAssert.Contains(error.Message, "partially written");
                Assert.AreEqual(1, support.CodeModule.Deletions);
                Assert.AreEqual(1, support.CodeModule.Insertions);
                Assert.AreEqual("", support.CodeModule.Source);
                var backups = Directory.GetFiles(fixture.BackupRoot, VbaTestRuntimeSource.ModuleName + ".bas", SearchOption.AllDirectories);
                Assert.AreEqual(1, backups.Length);
                Assert.AreEqual(original.Replace("\r\n", "\n").TrimEnd('\n'), File.ReadAllText(backups[0]).Replace("\r\n", "\n").TrimEnd('\n'));
            }
        }

        [STATestMethod]
        public void CollisionAndStaleNavigationAreRefusedWithoutCodeOrPaneMutations()
        {
            using (var fixture = new Fixture())
            {
                var catalog = fixture.Catalog();
                fixture.Service.Navigate(catalog, catalog.Tests.First());
                Assert.AreEqual(1, fixture.Project.VBComponents[0].CodeModule.CodePane.Shows);
                fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed location";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Navigate(catalog, catalog.Tests.First()));
                Assert.AreEqual(1, fixture.Project.VBComponents[0].CodeModule.CodePane.Shows);
                fixture.Project.VBComponents.Add(new FakeComponent { Name = VbaTestRuntimeSource.ModuleName, Type = 1,
                    CodeModule = new FakeCode { Source = "Public Sub UserCode()\nEnd Sub" } });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(fixture.Catalog(), "unreviewed"));
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
                Assert.IsFalse(Directory.Exists(fixture.BackupRoot));
            }
        }

        [STATestMethod]
        public void ProtocolRequiresRevisionModeAndExplicitIdsAndCoverageRemainsUnavailable()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var request = new Request { Command = "run_vba_tests", Project = fixture.Project.FileName,
                    ExpectedProjectVersion = "stale", ExpectedMode = 2, Items = catalog.Tests.Select(test => test.Id).ToArray() };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Command(request));
                request.ExpectedProjectVersion = catalog.Project.Revision;
                request.ExpectedMode = 1;
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(request));
                request.ExpectedMode = 2;
                request.Items = new string[0];
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(request));
                dynamic coverage = fixture.Service.Command(new Request { Command = "vba_test_coverage", Project = fixture.Project.FileName });
                Assert.IsFalse((bool)coverage.Available);
                StringAssert.Contains((string)coverage.Reason, "pass rate is a separate metric");
                dynamic preview = fixture.Service.PreviewSupport(fixture.Project.FileName);
                Assert.AreEqual(catalog.Project.Revision, (string)preview.ExpectedProjectVersion);
                Assert.AreEqual(VbaTestRuntimeSource.Generate(catalog), (string)preview.Text);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void InvalidPagingIsRefusedBeforeProjectLookupOrNativeDispatch()
        {
            using (var fixture = new Fixture())
            {
                foreach (string command in new[] { "vba_test_run_status", "vba_test_coverage" })
                    foreach (var pair in new[] { new[] { -1, 100 }, new[] { 0, -1 }, new[] { 0, 101 } })
                        Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(new Request {
                            Command = command, Project = "missing-project", Query = "missing-run", Offset = pair[0], Limit = pair[1] }));
                Assert.AreEqual(0, fixture.Host.Invocations);
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
            }
        }

        [STATestMethod]
        public void RunStatusPagesDoNotExecuteAgainAndKeepGlobalCountsAndHistoricalStaleness()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                int invocations = fixture.Host.Invocations;
                dynamic first = fixture.Service.Command(new Request { Command = "vba_test_run_status", Project = fixture.Project.FileName,
                    Query = run.Id, Offset = 0, Limit = 1 });
                dynamic second = fixture.Service.Command(new Request { Command = "vba_test_run_status", Project = fixture.Project.FileName,
                    Query = run.Id, Offset = 1, Limit = 1 });
                var left = (IDictionary<string, object>)first.Report;
                var right = (IDictionary<string, object>)second.Report;
                Assert.AreEqual(2, left["total"]); Assert.AreEqual(left["total"], right["total"]);
                Assert.AreEqual(1, ((object[])left["tests"]).Length); Assert.AreEqual(1, ((object[])right["tests"]).Length);
                Assert.AreNotEqual(((IDictionary<string, object>)((object[])left["tests"])[0])["id"],
                    ((IDictionary<string, object>)((object[])right["tests"])[0])["id"]);
                Assert.AreEqual(1, left["nextOffset"]); Assert.IsNull(right["nextOffset"]);
                Assert.AreEqual(left["passRate"], right["passRate"]);
                Assert.AreEqual(run.Revision, right["revision"]);
                fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed after run";
                dynamic historical = fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "compact", 1, 1);
                Assert.IsTrue((bool)historical.Stale);
                Assert.AreEqual(invocations, fixture.Host.Invocations, "Status pagination must never redispatch tests.");
            }
        }

        private static T Pump<T>(Task<T> task)
        {
            PumpMessagesUntil(() => task.IsCompleted);
            return task.GetAwaiter().GetResult();
        }

        private static void PumpMessagesUntil(Func<bool> completed)
        {
            var timer = Stopwatch.StartNew();
            while (!completed())
            {
                if (timer.ElapsedMilliseconds > 5000) Assert.Fail("The queued contract operation did not settle.");
                Application.DoEvents();
                Thread.Sleep(1);
            }
        }

        private static FakeProject Project(string name, string path)
        {
            var project = new FakeProject { Name = name, FileName = path };
            project.VBComponents.Add(new FakeComponent { Name = "TestsOne", Type = 1,
                CodeModule = new FakeCode { Source = "Option Explicit\n'@TestModule\n'@TestMethod\nPublic Sub Alpha()\nEnd Sub\n'@TestMethod\nPublic Sub Beta()\nEnd Sub" } });
            project.References.Add(new FakeReference { Name = "VBA", Guid = "reference-guid", Major = 4, Minor = 2 });
            return project;
        }

        private sealed class Fixture : IDisposable
        {
            internal readonly Control Dispatcher = new Control();
            internal readonly FakeVbe Vbe = new FakeVbe();
            internal readonly FakeProject Project = VbeTestExplorerServiceTests.Project("WorkbookProject", @"C:\Temp\Fixture.xlsm");
            internal readonly FakeNativeHost Host = new FakeNativeHost();
            internal readonly string BackupRoot = Path.Combine(Path.GetTempPath(), "VBAi-test-support-" + Guid.NewGuid().ToString("N"));
            internal readonly VbeTestExplorerService Service;
            internal Fixture()
            {
                var handle = Dispatcher.Handle;
                Vbe.VBProjects.Add(Project);
                Service = new VbeTestExplorerService(Vbe, Dispatcher) { Host = Host, IsExecutionHost = () => true, BackupRoot = () => BackupRoot };
                Service.ReadProjects();
            }
            internal VbaTestCatalog Catalog() => Service.DiscoverSelector(Project.FileName ?? Project.Name);
            internal void InstallFixtureSupport()
            {
                string source = VbaTestRuntimeSource.Generate(Catalog());
                Project.VBComponents.Add(new FakeComponent { Name = VbaTestRuntimeSource.ModuleName, Type = 1, CodeModule = new FakeCode { Source = source } });
            }
            public void Dispose()
            {
                Service.Dispose();
                Dispatcher.Dispose();
                if (Directory.Exists(BackupRoot)) Directory.Delete(BackupRoot, true);
            }
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Protection { get; set; }
            public int Mode { get; set; } = 2;
            public FakeComponents VBComponents { get; } = new FakeComponents();
            public List<FakeReference> References { get; } = new List<FakeReference>();
        }
        public sealed class FakeComponents : List<FakeComponent>
        {
            public int Additions { get; private set; }
            public FakeComponent Add(int type)
            {
                Additions++;
                var component = new FakeComponent { Name = "Module" + Additions, Type = type, CodeModule = new FakeCode() };
                Add(component);
                return component;
            }
        }
        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeCode CodeModule { get; set; }
        }
        public sealed class FakeReference
        {
            public string Name { get; set; }
            public string Guid { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
        }
        public sealed class FakeCode
        {
            public string Source { get; set; } = "";
            public int CountOfLines => TextLines.Length;
            public FakeLines Lines => new FakeLines(this);
            public FakePane CodePane { get; } = new FakePane();
            public bool ThrowOnAdd { get; set; }
            public Action OnAdd { get; set; }
            public Action OnInsert { get; set; }
            public int Deletions { get; private set; }
            public int Insertions { get; private set; }
            internal string[] TextLines => string.IsNullOrEmpty(Source) ? new string[0] : Source.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
            public void DeleteLines(int start, int count)
            {
                Deletions++;
                Source = string.Join("\r\n", TextLines.Take(start - 1).Concat(TextLines.Skip(start - 1 + count)));
            }
            public void AddFromString(string source)
            {
                Insertions++;
                OnAdd?.Invoke();
                if (ThrowOnAdd) throw new InvalidOperationException("Simulated write failure");
                Source = string.IsNullOrEmpty(Source) ? source : Source + "\r\n" + source;
            }
            public void InsertLines(int line, string source)
            {
                Insertions++;
                OnInsert?.Invoke();
                var lines = TextLines.ToList();
                lines.InsertRange(line - 1, source.Replace("\r\n", "\n").TrimEnd('\n').Split('\n'));
                Source = string.Join("\r\n", lines);
            }
            public void ReplaceLine(int line, string source)
            {
                Insertions++;
                OnInsert?.Invoke();
                var lines = TextLines;
                lines[line - 1] = source;
                Source = string.Join("\r\n", lines);
            }
        }
        public sealed class FakeLines
        {
            private readonly FakeCode code;
            internal FakeLines(FakeCode code) { this.code = code; }
            public string this[int first, int count] => string.Join("\r\n", code.TextLines.Skip(first - 1).Take(count));
        }
        public sealed class FakePane
        {
            public int Shows { get; private set; }
            public int SelectedLine { get; private set; }
            public void Show() { Shows++; }
            public void SetSelection(int firstLine, int firstColumn, int lastLine, int lastColumn) { SelectedLine = firstLine; }
        }
        private sealed class FakeNativeHost : VbeDebug.IProcedureValuesHost
        {
            internal int Resolutions, Invocations;
            internal Action<int> Resolving;
            internal bool ThrowOnInvoke;
            internal object Returned = new object[] { "Passed", "", "0" };
            public object ResolveTarget(object project, string expectedHostPath) { Resolutions++; Resolving?.Invoke(Resolutions); return project; }
            public object Invoke(object target, string module, string procedure, object[] arguments)
            {
                Assert.AreEqual(VbaTestRuntimeSource.ModuleName, module);
                Assert.AreEqual(VbaTestRuntimeSource.DispatcherProcedure, procedure);
                Assert.AreEqual("TestsOne", arguments[0]);
                Invocations++;
                if (ThrowOnInvoke) throw new InvalidOperationException("Simulated native transport failure");
                return Returned;
            }
        }
    }
}
