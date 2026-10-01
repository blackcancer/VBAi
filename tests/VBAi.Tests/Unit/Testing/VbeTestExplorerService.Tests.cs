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
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class VbeTestExplorerServiceTests
    {
        [STATestMethod]
        public void ReturnedWordValidationAndDispatchBalanceKnownLeasesAndRetainUnknown()
        {
            foreach (string outcome in new[] { "passed", "predispatch", "unknown" })
            using (var fixture = new Fixture())
            using (var leases = new ServiceWordLeaseRecorder())
            {
                fixture.InstallFixtureSupport();
                fixture.Host.TargetFactory = leases.Acquire;
                Assert.IsNull(fixture.Service.ExecutionUnavailableReason(fixture.Catalog()));
                Assert.AreEqual(leases.Targets.Count, leases.Released);
                if (outcome == "predispatch") fixture.Host.Resolving = count => {
                    if (count == 4) fixture.Project.Mode = 1;
                };
                fixture.Host.ThrowOnInvoke = outcome == "unknown";
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunAsync(catalog, new[] { catalog.Tests.First() }, null, CancellationToken.None));
                if (outcome == "unknown")
                {
                    Assert.IsTrue(run.OutcomeUnknown);
                    Assert.AreEqual(1, leases.Targets.Count(target => target.IsRetained));
                    Assert.AreEqual(leases.Targets.Count - 1, leases.Released);
                    var retained = leases.Targets.Single(target => target.IsRetained);
                    retained.Dispose(); Assert.IsNotNull(retained.Document);
                }
                else
                {
                    Assert.IsFalse(run.OutcomeUnknown);
                    Assert.AreEqual(leases.Targets.Count, leases.Released);
                    Assert.IsTrue(leases.Targets.All(target => target.Document == null));
                }
                Assert.AreEqual(outcome == "predispatch" ? 0 : 1, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void ReturnedWordReleaseFailureBeforeDispatchSettlesTheRunWithoutNativeRetry()
        {
            using (var fixture = new Fixture())
            using (var leases = new ServiceWordLeaseRecorder())
            {
                fixture.InstallFixtureSupport();
                fixture.Host.TargetFactory = leases.Acquire;
                fixture.Host.Resolving = count => { if (count == 3) fixture.Project.Mode = 1; };
                leases.ThrowOnRelease = 3;
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunAsync(catalog, new[] { catalog.Tests.First() }, null, CancellationToken.None));
                Assert.IsFalse(run.OutcomeUnknown);
                Assert.AreEqual(0, fixture.Host.Invocations);
                Assert.AreEqual(3, leases.Released);
                StringAssert.Contains(run.Results.Single().Message, "design mode");
                StringAssert.Contains(run.Results.Single().Message, "Release failed");
                Assert.AreEqual(VbaTestOutcome.Blocked, run.Results.Single().Outcome);
            }
        }
        private sealed class ServiceWordLeaseRecorder : IDisposable
        {
            private readonly Func<object, bool> priorCheck = VbaTestWordValuesHost.IsComReference;
            private readonly Func<object, int> priorRelease = VbaTestWordValuesHost.ReleaseComReference;
            internal readonly List<VbaTestWordValuesHost.OwnedTarget> Targets = new List<VbaTestWordValuesHost.OwnedTarget>();
            private readonly HashSet<object> documents = new HashSet<object>();
            internal int Released, ThrowOnRelease;
            internal ServiceWordLeaseRecorder()
            {
                VbaTestWordValuesHost.IsComReference = value => documents.Contains(value);
                VbaTestWordValuesHost.ReleaseComReference = value => { if (++Released == ThrowOnRelease) throw new InvalidOperationException("Release failed"); return 0; };
            }
            internal void RegisterDocument(object document) { documents.Add(document); }
            internal object Acquire(object project)
            {
                var document = new VbaTestWordValuesHostTests.Document(); documents.Add(document);
                var target = new VbaTestWordValuesHost.OwnedTarget { Owner = new VbaTestWordValuesHost(), Document = document, Project = project };
                Targets.Add(target); return target;
            }
            public void Dispose() { VbaTestWordValuesHost.IsComReference = priorCheck; VbaTestWordValuesHost.ReleaseComReference = priorRelease; }
        }
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
        public void CoverageStartReturnsItsReservedQueryBeforeAnyCopyPreparationOnTheOwner()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                int owner = Thread.CurrentThread.ManagedThreadId;
                coverage.AfterCopy = _ => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                var catalog = fixture.Catalog();
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray(), null, true);
                string query = start.Query;
                Assert.IsFalse(string.IsNullOrEmpty(query));
                Assert.AreEqual("Running", (string)start.State);
                Assert.IsTrue((bool)start.Pending);
                Assert.IsNull(coverage.Clone, "No copy may be opened before StartRun returns its query.");
                Assert.AreEqual(0, coverage.Events.Count);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StartRun(fixture.Project.FileName,
                    catalog.Project.Revision, catalog.Tests.Select(test => test.Id).ToArray(), null, true));
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, query, "compact")).Pending));
                Assert.AreEqual("Completed", (string)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, query, "compact")).State);
                CollectionAssert.AreEqual(new[] { "Compile", "Reset", "Test:Alpha", "Test:Beta", "Snapshot", "Close" }, coverage.Events);
            }
        }

        [STATestMethod]
        public void QueuedCoverageRevalidatesRevisionModeIdentityAndPermissionBeforeOpeningTheCopy()
        {
            foreach (string change in new[] { "source", "mode", "identity", "permission" })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                var catalog = fixture.Catalog();
                bool allowed = true;
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray(), () => {
                        if (!allowed) throw new InvalidOperationException("Coverage permission withdrawn.");
                    }, true);
                string query = start.Query;
                var entries = (System.Collections.IDictionary)typeof(VbeTestExplorerService)
                    .GetField("runs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                var entry = entries[query];
                var task = (Task<VbaTestRun>)entry.GetType().GetField("Completion", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(entry);
                if (change == "source") fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed before coverage preparation";
                if (change == "mode") fixture.Project.Mode = 1;
                if (change == "identity") fixture.Vbe.VBProjects[0] = Project(fixture.Project.Name, fixture.Project.FileName);
                if (change == "permission") allowed = false;
                var run = Pump(task);
                Assert.IsNull(coverage.Clone, change);
                Assert.AreEqual(0, coverage.Events.Count, change);
                Assert.IsFalse(run.OutcomeUnknown, change);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked), change);
                Assert.AreEqual("Aborted", entry.GetType().GetField("State", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(entry), change);
            }
        }

        [STATestMethod]
        public void StopAndExternalCancellationBeforeQueuedCoveragePreparationDoNotOpenACopy()
        {
            foreach (bool external in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            using (var cancellation = new CancellationTokenSource())
            {
                var catalog = fixture.Catalog();
                Task<VbaTestRun> task;
                if (external)
                {
                    task = fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, cancellation.Token);
                    cancellation.Cancel();
                }
                else
                {
                    dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                        catalog.Tests.Select(test => test.Id).ToArray(), null, true);
                    string query = start.Query;
                    fixture.Service.StopRun(fixture.Project.FileName, query);
                    var entries = (System.Collections.IDictionary)typeof(VbeTestExplorerService)
                        .GetField("runs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                    var entry = entries[query];
                    task = (Task<VbaTestRun>)entry.GetType().GetField("Completion", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(entry);
                }
                var run = Pump(task);
                Assert.IsNull(coverage.Clone);
                Assert.AreEqual(0, coverage.Events.Count);
                Assert.IsFalse(run.OutcomeUnknown);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Cancelled));
                Assert.AreEqual("Cancelled", (string)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, run.Id, "compact")).State);
            }
        }

        [STATestMethod]
        public void DisposedQueuedCoverageSettlesOnTheOwnerWithoutCopyPreparationOrAnOrphanedRun()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                int owner = Thread.CurrentThread.ManagedThreadId;
                var catalog = fixture.Catalog();
                var continuations = OwnerContinuations(fixture.Service);
                var task = fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), _ =>
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId), CancellationToken.None);
                fixture.Service.Dispose();
                fixture.Dispatcher.Dispose();
                Assert.IsFalse(continuations.IsDisposed, "The posted start still needs its owned continuation handle.");
                var run = Pump(task);
                Assert.IsNull(coverage.Clone);
                Assert.AreEqual(0, coverage.Events.Count);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Cancelled));
                Assert.IsTrue(continuations.IsDisposed);
                Assert.IsNull(typeof(VbeTestExplorerService).GetField("active", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service));
            }
        }

        [STATestMethod]
        public void FailedCoverageStartPublicationIsTerminalAndReleasesTheActiveReservationWithoutRetry()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                var catalog = fixture.Catalog();
                OwnerContinuations(fixture.Service).Dispose();
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray(), null, true);
                Assert.AreEqual("Aborted", (string)start.State);
                Assert.IsFalse((bool)start.Pending);
                Assert.IsFalse(string.IsNullOrEmpty((string)start.Query));
                Assert.IsNull(coverage.Clone);
                Assert.AreEqual(0, coverage.Events.Count);
                Assert.IsNull(typeof(VbeTestExplorerService).GetField("active", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service));
                var entries = (System.Collections.IDictionary)typeof(VbeTestExplorerService)
                    .GetField("runs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                var entry = entries[(string)start.Query];
                var task = (Task<VbaTestRun>)entry.GetType().GetField("Completion", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(entry);
                Assert.IsInstanceOfType(Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult()), typeof(InvalidOperationException));
            }
        }

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

        [STATestMethod]
        public void HostFactoriesAndRegistrationVerificationCheckEveryIdentityFieldWithoutRegistryWrites()
        {
            foreach (string name in new[] { "EXCEL", "WINWORD", "POWERPNT", "excel", "OUTLOOK", "testhost" })
            {
                Assert.AreEqual(new[] { "EXCEL", "WINWORD", "POWERPNT" }.Contains(name.ToUpperInvariant()), VbeTestExplorerService.IsReturnedValuesHost(name));
                object host = VbeTestExplorerService.CreateReturnedValuesHost(name);
                Assert.IsInstanceOfType(host, name.Equals("WINWORD", StringComparison.OrdinalIgnoreCase) ? typeof(VbaTestWordValuesHost)
                    : name.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase) ? typeof(VbaTestPowerPointValuesHost) : typeof(VbeDebug.NativeProcedureValuesHost));
            }
            var values = new Dictionary<string, object> { ["CodeBase"] = new Uri(typeof(VbaTestRuntime).Assembly.Location).AbsoluteUri,
                ["Class"] = typeof(VbaTestRuntime).FullName, ["Assembly"] = typeof(VbaTestRuntime).Assembly.FullName,
                ["RuntimeVersion"] = "v4.0.30319", ["ThreadingModel"] = "Both", [""] = "mscoree.dll" };
            Func<string, string, object> read = (key, name) => name == null ? (object)true
                : key.StartsWith("VBAi.", StringComparison.Ordinal) ? "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}" : values[name];
            Assert.IsNull(VbeTestExplorerService.NativeRuntimeRegistrationReason(read));
            foreach (string key in values.Keys.ToArray())
            {
                var original = values[key]; values[key] = "wrong";
                StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason(read), "does not belong");
                values[key] = original;
            }
            values["CodeBase"] = "https://example.invalid/runtime.dll";
            StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason(read), "does not belong");
            values["CodeBase"] = new Uri(Path.Combine(Path.GetTempPath(), "other.dll")).AbsoluteUri;
            StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason(read), "does not belong");
            values["CodeBase"] = new Uri(typeof(VbaTestRuntime).Assembly.Location).AbsoluteUri;
            StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason((key, name) => key.StartsWith("VBAi.", StringComparison.Ordinal) ? null : read(key, name)), "does not belong");
            StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason((key, name) => false), "Register");
            StringAssert.Contains(VbeTestExplorerService.NativeRuntimeRegistrationReason((key, name) => throw new InvalidOperationException("Read refused")), "Read refused");
            string absent = "VBAi.NonexistentCoverageFixture-" + Guid.NewGuid().ToString("N");
            Assert.AreEqual(false, VbeTestExplorerService.ReadRuntimeRegistrationValue(absent, null));
            Assert.IsNull(VbeTestExplorerService.ReadRuntimeRegistrationValue(absent, "CodeBase"));
            Assert.AreEqual(true, VbeTestExplorerService.ReadRuntimeRegistrationValue("CLSID", null));
            VbeTestExplorerService.ReadRuntimeRegistrationValue("CLSID", "NonexistentCoverageValue");
            string actual = VbeTestExplorerService.NativeRuntimeRegistrationReason();
            Assert.IsTrue(actual == null || actual.Length > 0);
        }

        [STATestMethod]
        public void ServiceRejectsMissingIdentitiesWrongThreadAndUnknownRunsWithoutInvokingTheHost()
        {
            using (var control = new Control())
            {
                Assert.ThrowsException<ArgumentNullException>(() => new VbeTestExplorerService(null, control));
                Assert.ThrowsException<ArgumentNullException>(() => new VbeTestExplorerService(new FakeVbe(), null));
            }
            using (var fixture = new Fixture())
            {
                Assert.IsNotNull(fixture.Service.ExecutionUnavailableReason(null));
                Assert.IsNotNull(fixture.Service.CoverageUnavailableReason(null));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Discover(null));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Discover("missing"));
                var refusal = Task.Run(() => { try { fixture.Service.ReadProjects(); return null; } catch (Exception error) { return error; } }).GetAwaiter().GetResult();
                Assert.IsInstanceOfType(refusal, typeof(InvalidOperationException));
                var catalog = fixture.Catalog();
                Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Service.InvokeAsync(catalog, catalog.Tests.First(), "Test"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunStatus(fixture.Project.FileName, null, "compact"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StopRun(fixture.Project.FileName, "absent"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StartRun(fixture.Project.FileName, null, new[] { catalog.Tests.First().Id }));
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision, null));
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision, new string[0]));
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void DirectBridgeCommandsValidateScopeAndReturnBothReportFormatsAndNavigation()
        {
            using (var fixture = new Fixture())
            {
                var catalog = fixture.Catalog();
                Func<string, Request> request = command => new Request { Command = command, Project = fixture.Project.FileName,
                    ExpectedProjectVersion = catalog.Project.Revision, ExpectedMode = 2, Items = new[] { catalog.Tests.First().Id } };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Command(request("show_vba_test_explorer")));
                fixture.Service.ShowExplorer = projectId => projectId;
                Assert.AreEqual(catalog.Project.Id, fixture.Service.Command(request("show_vba_test_explorer")));
                dynamic preview = fixture.Service.Command(request("preview_vba_test_support"));
                Assert.AreEqual(catalog.Project.Revision, (string)preview.ExpectedProjectVersion);
                fixture.Service.Command(request("discover_vba_tests"));
                var navigate = request("navigate_vba_test");
                Assert.IsTrue((bool)((dynamic)fixture.Service.Command(navigate)).Navigated);
                foreach (var items in new[] { null, new string[0], new[] { "a", "b" } })
                { navigate.Items = items; Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(navigate)); }
                var install = request("install_vba_test_support"); install.Text = preview.Text;
                Assert.IsTrue((bool)((dynamic)fixture.Service.Command(install)).Applied);
                catalog = fixture.Catalog();
                var invalid = request("unknown"); Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(invalid));
                invalid.ExpectedMode = 1; Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(invalid));
                invalid.ExpectedProjectVersion = "stale"; Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Command(invalid));
                var runRequest = request("run_vba_tests"); runRequest.Action = "invalid";
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(runRequest));
                runRequest.Action = "";
                dynamic start = fixture.Service.Command(runRequest);
                string id = start.Query;
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, id, "compact")).Pending));
                var status = request("vba_test_run_status"); status.Query = id; status.Action = "human";
                Assert.IsInstanceOfType(((dynamic)fixture.Service.Command(status)).Report, typeof(string));
                status.Action = "compact"; Assert.IsNotNull(fixture.Service.Command(status));
                status.Action = "invalid"; Assert.ThrowsException<ArgumentException>(() => fixture.Service.Command(status));
                var stop = request("stop_vba_tests"); stop.Query = id; fixture.Service.Command(stop);
                var coverage = request("vba_test_coverage"); coverage.Query = id;
                Assert.IsFalse((bool)((dynamic)fixture.Service.Command(coverage)).Available);
                coverage.Query = "missing"; Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Command(coverage));
                coverage.Query = null; Assert.IsNotNull(fixture.Service.Command(coverage));
                Assert.AreEqual(1, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void SupportAndRunSelectionGuardsRejectActiveRunsAndRetainOnlyTwentyReports()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.RunAsync(catalog, null, null, CancellationToken.None));
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.RunAsync(catalog, new VbaTestDescriptor[0], null, CancellationToken.None));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunAsync(catalog,
                    new[] { new VbaTestDescriptor { Id = catalog.Tests.First().Id, Module = "other", Procedure = catalog.Tests.First().Procedure } }, null, CancellationToken.None));
                var first = fixture.Service.RunAsync(catalog, new[] { catalog.Tests.First(), catalog.Tests.First() }, null, CancellationToken.None);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.InstallSupport(catalog));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(catalog, VbaTestRuntimeSource.Generate(catalog)));
                string firstId = Pump(first).Id;
                for (int index = 0; index < 20; index++) Pump(fixture.Service.RunAsync(catalog, new[] { catalog.Tests.First() }, null, CancellationToken.None));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunStatus(fixture.Project.FileName, firstId, "compact"));
                Assert.AreEqual(21, fixture.Host.Invocations, "Duplicate selection executes once; expired report reads do not execute.");
            }
        }

        [STATestMethod]
        public void NavigationAndReviewedSupportRefuseMissingModuleStaleSourceAndUnownedNameCollision()
        {
            using (var fixture = new Fixture())
            {
                var catalog = fixture.Catalog();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Navigate(catalog, new VbaTestDescriptor { Module = "missing", Line = 1 }));
                fixture.Project.VBComponents[0].CodeModule.Source += "\n' changed";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Navigate(catalog, catalog.Tests.First()));
                fixture.Service.ConfirmSupport = (plan, before, after) => true;
                fixture.Service.InstallSupport(fixture.Catalog());
                Assert.AreEqual(1, fixture.Project.VBComponents.Additions);
            }
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents.Add(new FakeComponent { Name = VbaTestRuntimeSource.ModuleName, Type = 1, CodeModule = new FakeCode() });
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Generate(fixture.Catalog()));
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
            }
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var support = fixture.Project.VBComponents.Last(); support.Type = 2;
                var catalog = fixture.Catalog();
                string generated = VbaTestRuntimeSource.Generate(catalog);
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(catalog, generated));
                StringAssert.Contains(error.Message, "occupied by user code");
                Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
                Assert.AreEqual(0, support.CodeModule.Insertions + support.CodeModule.Deletions);
                Assert.IsFalse(Directory.Exists(fixture.BackupRoot));
            }
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                string generated = VbaTestRuntimeSource.Generate(catalog);
                var support = fixture.Project.VBComponents.Last(); int reads = 0;
                support.CodeModule.BeforeReadLines = () => { if (++reads == 2) support.CodeModule.Source = "' user source after validation"; };
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(catalog, generated));
                StringAssert.Contains(error.Message, "occupied by user code");
                Assert.AreEqual(2, reads); Assert.AreEqual(0, fixture.Project.VBComponents.Additions);
                Assert.AreEqual(0, support.CodeModule.Insertions + support.CodeModule.Deletions);
                Assert.IsFalse(Directory.Exists(fixture.BackupRoot));
            }
        }

        [STATestMethod]
        public void NativeServiceRoutesVerifiedVerdictsAndLatchesUncertainTransportFailures()
        {
            foreach (string failure in new[] { "none", "prepare", "execute" })
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                fixture.Service.IsExecutionHost = () => false;
                fixture.Service.IsNativeExecutionHost = () => true;
                fixture.Service.NativeRuntimeReason = () => null;
                var native = (VbaNativeTestExecutionHost)typeof(VbeTestExplorerService).GetField("nativeExecutionHost", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                var probe = new ServiceNativeProbe(); native.Probe = probe;
                probe.Preparing = () => { if (failure == "prepare") throw new InvalidOperationException("Preparation refused"); };
                probe.Executing = () => {
                    if (failure == "execute") throw new InvalidOperationException("Native completion unavailable");
                    var runtime = new VbaTestRuntime();
                    var job = (object[])runtime.Request(VbaTestRuntimeSource.Version, VbaTestRuntimeSource.DispatchSignature(catalog));
                    Assert.IsTrue(runtime.Publish((string)job[0], (string)job[4], (string)job[5], (string)job[6], "Passed", "Verified callback", 0));
                };
                Assert.IsNull(fixture.Service.ExecutionUnavailableReason(catalog));
                var run = Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.AreEqual(failure == "execute", run.OutcomeUnknown);
                Assert.AreEqual(failure == "none" ? VbaTestOutcome.Passed : failure == "prepare" ? VbaTestOutcome.Blocked : VbaTestOutcome.OutcomeUnknown, run.Results.First().Outcome, failure);
                Assert.AreEqual(2, run.Results.Count);
                Assert.AreEqual(failure == "none" ? 2 : failure == "prepare" ? 0 : 1, probe.Executions);
                Assert.AreEqual(0, fixture.Host.Invocations);
                if (failure == "execute") Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Service.Validate(catalog));
            }
        }

        [STATestMethod]
        public void NativeGuardPermitsOnlyCleanupAfterStopAndChecksTheRunAuthority()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                var active = typeof(VbeTestExplorerService).GetField("active", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                ((CancellationTokenSource)active.GetType().GetField("Stop", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(active)).Cancel();
                int guards = 0;
                active.GetType().GetField("ExecutionGuard", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(active, (Action)(() => guards++));
                var guard = typeof(VbeTestExplorerService).GetMethod("NativeExecutionGuard", BindingFlags.Instance | BindingFlags.NonPublic);
                foreach (string phase in new[] { "Test", "ModuleInitialize", "TestCleanup", "ModuleCleanup" })
                {
                    typeof(VbeTestExplorerService).GetField("nativePhase", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture.Service, phase);
                    if (phase.EndsWith("Cleanup", StringComparison.Ordinal)) guard.Invoke(fixture.Service, null);
                    else Assert.IsInstanceOfType(Assert.ThrowsException<TargetInvocationException>(() => guard.Invoke(fixture.Service, null)).InnerException, typeof(VbaTestInvocationException));
                }
                typeof(VbeTestExplorerService).GetField("nativePhase", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture.Service, null);
                Pump(task);
                guard.Invoke(fixture.Service, null);
                var invoke = typeof(VbeTestExplorerService).GetMethod("InvokeNativeAsync", BindingFlags.Instance | BindingFlags.NonPublic);
                fixture.Service.NativeRuntimeReason = () => "Registration missing";
                fixture.Service.IsExecutionHost = () => false; fixture.Service.IsNativeExecutionHost = () => true;
                var refused = (Task<VbaTestResult>)invoke.Invoke(fixture.Service, new object[] { catalog, catalog.Tests.First(), "Test" });
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => refused.GetAwaiter().GetResult());
                Assert.IsFalse(error.Uncertain); Assert.AreEqual("Registration missing", error.Message);
                Assert.IsNull(typeof(VbeTestExplorerService).GetField("nativePhase", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service));
            }
        }

        private sealed class ServiceNativeProbe : VbaNativeTestExecutionHost.IProbe
        {
            internal Action Owning, Preparing, Executing;
            internal int Executions;
            public void RequireOwner(object vbe) { Owning?.Invoke(); }
            public object Prepare(object vbe, object project, string source) { Preparing?.Invoke(); return this; }
            public void Revalidate(object vbe, object project, object prepared) { Assert.AreSame(this, prepared); }
            public void Execute(object prepared) { Executions++; Executing?.Invoke(); }
            public int ReadMode(object project) => 2;
        }

        [STATestMethod]
        public void DefaultsUseLocalPathsAndOwnedModalReviewAndContinuationFailuresAreReleased()
        {
            using (var dispatcher = new Control())
            using (var service = new VbeTestExplorerService(new FakeVbe(), dispatcher))
            {
                Assert.IsFalse(service.IsExecutionHost());
                StringAssert.Contains(service.BackupRoot(), "VBAi");
                StringAssert.Contains(service.CoverageRoot(), "CoverageRuns");
                service.NativeRuntimeReason();
                var disposedControl = new Control(); disposedControl.Dispose();
                Assert.ThrowsException<ObjectDisposedException>(() => new VbeTestExplorerService(new FakeVbe(), dispatcher, () => disposedControl));
                Assert.IsTrue(disposedControl.IsDisposed);
                var error = Task.Run(() => { try { service.AwaitOwner(Task.FromResult(1)).GetAwaiter().GetResult(); return null; } catch (Exception caught) { return caught; } }).GetAwaiter().GetResult();
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
                using (var timer = new System.Windows.Forms.Timer { Interval = 20 })
                {
                    timer.Tick += (_, __) => {
                        foreach (Form form in Application.OpenForms)
                            if (form is TestSupportReviewDialog) { timer.Stop(); form.DialogResult = DialogResult.Cancel; break; }
                    };
                    timer.Start();
                    Assert.IsFalse(service.ConfirmSupport(new VbaTestCatalog { Project = new VbaTestProjectSnapshot { Name = "Disposable review" } }, "before", "after"));
                }
            }
        }

        [STATestMethod]
        public void SnapshotLimitsAndMissingSavedPathsRefuseBeforeTestExecution()
        {
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents[0].CodeModule.Source = new string('x', 1) + string.Concat(Enumerable.Repeat("\n'", 200001));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Catalog());
            }
            using (var fixture = new Fixture())
            {
                for (int index = 0; index < 1000; index++) fixture.Project.VBComponents.Add(new FakeComponent { Name = "Module" + index, Type = 1, CodeModule = new FakeCode() });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Catalog());
            }
            foreach (string path in new[] { null, "relative.xlsm" })
            using (var fixture = new Fixture())
            {
                fixture.Project.FileName = path; fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(catalog), "Save");
                StringAssert.Contains(fixture.Service.ExecutionUnavailableReason(catalog), "Save");
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void ProgressFailureAbortsTheStoredRunAndReleasesOwnershipForANewRun()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), _ => throw new InvalidOperationException("Progress consumer failed"), CancellationToken.None);
                PumpMessagesUntil(() => task.IsCompleted);
                Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                var records = (System.Collections.IDictionary)typeof(VbeTestExplorerService).GetField("runs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                string id = records.Keys.Cast<string>().Single();
                dynamic status = fixture.Service.RunStatus(fixture.Project.FileName, id, "compact");
                Assert.AreEqual("Aborted", (string)status.State); Assert.IsFalse((bool)status.Pending);
                Assert.AreEqual(2, Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None)).Results.Count);
            }
        }

        [STATestMethod]
        public void SynchronousIgnoredRunsAreStoredAndUnavailableSupportAfterQueueCannotDispatch()
        {
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents[0].CodeModule.Source = "'@TestModule\n'@TestMethod\n'@Ignore reviewed skip\nPublic Sub Alpha()\nEnd Sub";
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision, catalog.Tests.Select(test => test.Id).ToArray());
                Assert.IsFalse((bool)start.Pending); Assert.AreEqual("Completed", (string)start.State);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                var task = fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                fixture.Host.Resolving = _ => throw new InvalidOperationException("Selected host no longer available");
                var run = Pump(task);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked));
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void NativeRegistrationRefusalFallsBackOnlyForReturnedValuesHostsAndOwnerFailuresAreReported()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                fixture.Service.IsNativeExecutionHost = () => true;
                fixture.Service.NativeRuntimeReason = () => "Callback missing";
                Assert.IsNull(fixture.Service.ExecutionUnavailableReason(catalog));
                Assert.AreEqual(2, Pump(fixture.Service.RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None)).Results.Count);
                fixture.Service.IsExecutionHost = () => false; fixture.Service.NativeRuntimeReason = () => null;
                var host = (VbaNativeTestExecutionHost)typeof(VbeTestExplorerService).GetField("nativeExecutionHost", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fixture.Service);
                host.Probe = new ServiceNativeProbe { Owning = () => throw new InvalidOperationException("Native owner unavailable") };
                Assert.AreEqual("Native owner unavailable", fixture.Service.ExecutionUnavailableReason(catalog));
                var runId = (Func<string>)typeof(VbaNativeTestExecutionHost).GetField("runId", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(host);
                Assert.ThrowsException<InvalidOperationException>(() => runId());
                host.Dispose();
                typeof(VbeTestExplorerService).GetField("nativeExecutionHost", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture.Service, null);
                fixture.Service.Dispose();
                typeof(VbeTestExplorerService).GetMethod("ReleaseOwnerContinuationsWhenIdle", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fixture.Service, null);
            }
        }

        [STATestMethod]
        public void SourceNormalizationMissingRevisionAndOtherProjectStopDoNotBypassTheBoundaries()
        {
            Assert.AreEqual(VbeTestExplorerService.Hash(""), VbeTestExplorerService.Hash(null));
            var canonical = typeof(VbeTestExplorerService).GetMethod("Canonical", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.AreEqual("", canonical.Invoke(null, new object[] { null }));
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport(); var catalog = fixture.Catalog();
                var request = new Request { Command = "run_vba_tests", Project = fixture.Project.FileName, ExpectedMode = 2, Items = catalog.Tests.Select(test => test.Id).ToArray() };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Command(request));
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision, request.Items);
                string id = start.Query;
                var other = Project("Other", @"C:\Temp\Other.xlsm"); fixture.Vbe.VBProjects.Add(other);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StopRun(other.FileName, id));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.StopRun(fixture.Project.FileName, null));
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, id, "compact")).Pending));
                fixture.Service.ConfirmSupport = (plan, before, after) => false;
                fixture.Service.InstallSupport(fixture.Catalog());
                var support = fixture.Project.VBComponents.Single(component => component.Name == VbaTestRuntimeSource.ModuleName);
                support.CodeModule.AfterAdd = () => support.CodeModule.Source += "\n' unexpected native edit";
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplySupport(fixture.Catalog(), VbaTestRuntimeSource.Generate(fixture.Catalog())));
                StringAssert.Contains(error.Message, "did not match"); StringAssert.Contains(error.Message, "Preserved backup");
            }
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
            internal VbaTestCatalog Catalog() => Service.DiscoverSelector(!string.IsNullOrEmpty(Project.FileName) && Path.IsPathRooted(Project.FileName) ? Project.FileName : Project.Name);
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

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
            public object ActiveVBProject { get; set; }
            public object ActiveCodePane { get; set; }
            public VbeTestExplorerServiceCompilationTests.CompilerWindow MainWindow { get; } = new VbeTestExplorerServiceCompilationTests.CompilerWindow();
            public VbeTestExplorerServiceCompilationTests.CompilerCommandBars CommandBars { get; } = new VbeTestExplorerServiceCompilationTests.CompilerCommandBars();
        }
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
            public Action AfterAdd { get; set; }
            public Action BeforeReadLines { get; set; }
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
                AfterAdd?.Invoke();
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
            public string this[int first, int count] { get { code.BeforeReadLines?.Invoke(); return string.Join("\r\n", code.TextLines.Skip(first - 1).Take(count)); } }
        }
        public sealed class FakePane
        {
            public VbeTestExplorerServiceCompilationTests.CompilerWindow Window { get; } = new VbeTestExplorerServiceCompilationTests.CompilerWindow();
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
            internal Func<object, object> TargetFactory;
            internal object Returned = new object[] { "Passed", "", "0" };
            public object ResolveTarget(object project, string expectedHostPath) { Resolutions++; Resolving?.Invoke(Resolutions); return TargetFactory == null ? project : TargetFactory(project); }
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
