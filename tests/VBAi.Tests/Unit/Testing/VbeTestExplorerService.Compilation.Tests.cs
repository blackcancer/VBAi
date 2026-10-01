using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeTestExplorerServiceCompilationTests
    {
        [STATestMethod]
        public void CompilationExecutesOnceThenWaitsForOwnerIdleAndDelayedDisabledState()
        {
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                Assert.AreEqual(1, fixture.Control.Executions);
                int readsBeforeObserver = fixture.Control.Reads, guards = 0;
                var observed = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => guards++);
                Assert.IsFalse(observed.IsCompleted);
                Assert.IsNull(fixture.Timer, "The observer must first yield the owner UI.");
                Assert.AreEqual(readsBeforeObserver, fixture.Control.Reads);
                Application.DoEvents();
                Assert.IsNotNull(fixture.Timer);
                Assert.AreEqual(readsBeforeObserver, fixture.Control.Reads, "Starting the timer cannot claim compilation completion.");
                fixture.Timer.Tick();
                Assert.IsFalse(observed.IsCompleted);
                fixture.Control.State = false;
                fixture.Timer.Tick();
                Assert.IsTrue(observed.GetAwaiter().GetResult());
                Assert.AreEqual(3, guards, "A disabled state requires a second guard after the COM getter returns.");
                Assert.AreEqual(1, fixture.Control.Executions);
                Assert.IsTrue(fixture.Timer.Disposed);
            }
        }

        [STATestMethod]
        public void PermanentlyEnabledControlHasABoundedKnownRefusalWithoutSecondCompilation()
        {
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                var observed = fixture.Observe();
                fixture.Timer.Tick();
                Assert.IsFalse(observed.IsCompleted);
                fixture.Clock = 3000;
                fixture.Timer.Tick();
                var error = Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                StringAssert.Contains(error.Message, "three seconds");
                Assert.IsFalse((Exception)error is VbaTestInvocationException);
                Assert.AreEqual(1, fixture.Control.Executions);
                Assert.IsTrue(fixture.Timer.Disposed);
            }
        }

        [STATestMethod]
        public void ChangedProjectAndNonDesignModeCannotUseAnotherDisabledCompileControlAsProof()
        {
            foreach (string fault in new[] { "identity", "mode" })
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                var observed = fixture.Observe();
                int reads = fixture.Control.Reads;
                fixture.Control.State = false;
                if (fault == "identity") fixture.Vbe.ActiveVBProject = new CompilerProject { Name = fixture.Project.Name };
                else fixture.Project.Mode = 1;
                fixture.Timer.Tick();
                Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                Assert.AreEqual(reads, fixture.Control.Reads);
                Assert.AreEqual(1, fixture.Control.Executions);
                Assert.IsTrue(fixture.Timer.Disposed);
            }
        }

        [STATestMethod]
        public void RevokedGuardAndCancellationStopBeforeReadingCompilationState()
        {
            foreach (bool cancel in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                int guards = 0, reads = fixture.Control.Reads;
                var observed = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => {
                    guards++;
                    if (cancel) throw new OperationCanceledException();
                    throw new InvalidOperationException("The original project approval was revoked.");
                });
                Application.DoEvents(); fixture.Timer.Tick();
                if (cancel) Assert.ThrowsException<OperationCanceledException>(() => observed.GetAwaiter().GetResult());
                else Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                Assert.AreEqual(1, guards);
                Assert.AreEqual(reads, fixture.Control.Reads);
                Assert.IsTrue(fixture.Timer.Disposed);
            }
        }

        [STATestMethod]
        public void ExecuteFailureOrProjectChangeDuringExecuteIsUncertainAndNeverRetried()
        {
            foreach (bool changedProject in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.Control.OnExecute = () => {
                    if (changedProject) fixture.Vbe.ActiveVBProject = new CompilerProject { Name = "Other" };
                    else throw new InvalidOperationException("Native command completion unavailable");
                };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Service.CompileCoverageProject(fixture.Project));
                Assert.IsTrue(error.Uncertain);
                Assert.AreEqual(1, fixture.Control.Executions);
                Assert.IsNull(fixture.Timer);
            }
        }

        [STATestMethod]
        public void DispatcherOrServiceDisposalSettlesObserverAndPreventsFurtherComReads()
        {
            foreach (bool disposeDispatcher in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                var observed = fixture.Observe();
                int reads = fixture.Control.Reads;
                if (disposeDispatcher) fixture.Dispatcher.Dispose();
                else { fixture.Service.Dispose(); fixture.Timer.Tick(); }
                Assert.ThrowsException<ObjectDisposedException>(() => observed.GetAwaiter().GetResult());
                Assert.IsTrue(fixture.Timer.Disposed);
                fixture.Timer.Tick();
                Assert.AreEqual(reads, fixture.Control.Reads);
            }
        }

        [STATestMethod]
        public void AlreadyDisabledControlStillRequiresOneDeferredOwnerObservation()
        {
            using (var fixture = new Fixture())
            {
                fixture.Control.State = false;
                fixture.Service.CompileCoverageProject(fixture.Project);
                var observed = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => { });
                Assert.IsFalse(observed.IsCompleted);
                Application.DoEvents(); fixture.Timer.Tick();
                Assert.IsTrue(observed.GetAwaiter().GetResult());
                Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void FocusPumpingRevalidatesProjectAndApprovalBeforeExecutingCompiler()
        {
            foreach (bool revokeApproval in new[] { false, true })
            using (var fixture = new Fixture())
            {
                bool permitted = true;
                typeof(VbeTestExplorerService).GetField("currentCoverageCompileGuard", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(fixture.Service, new Action(() => {
                        if (!permitted) throw new InvalidOperationException("Approval revoked during focus.");
                    }));
                fixture.Pane.Window.OnFocus = () => {
                    if (revokeApproval) permitted = false;
                    else fixture.Vbe.ActiveVBProject = new CompilerProject { Name = fixture.Project.Name };
                };
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.CompileCoverageProject(fixture.Project));
                Assert.IsFalse((Exception)error is VbaTestInvocationException, "No compiler command was dispatched.");
                Assert.IsTrue(fixture.Vbe.MainWindow.Visible);
                Assert.AreEqual(1, fixture.Pane.Shows);
                Assert.AreEqual(1, fixture.Pane.Window.Focuses);
                Assert.AreSame(fixture.Pane, fixture.Vbe.ActiveCodePane);
                Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void ObserverStartupFailureSettlesTaskWithoutReexecutingCommand()
        {
            foreach (bool throwOnStart in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                fixture.Service.StartCoverageCompilationTimer = tick => {
                    if (throwOnStart) throw new InvalidOperationException("Timer unavailable.");
                    return null;
                };
                var observed = fixture.Observe();
                Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                Assert.AreEqual(1, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void CustomOrReplacedCommandCannotExecuteOrProveCompilation()
        {
            foreach (string fault in new[] { "id", "type", "builtin", "action" })
            foreach (bool afterExecute in new[] { false, true })
            using (var fixture = new Fixture())
            {
                Task<bool> observed = null;
                if (afterExecute) { fixture.Service.CompileCoverageProject(fixture.Project); observed = fixture.Observe(); }
                if (fault == "id") fixture.Control.Id = 999;
                else if (fault == "type") fixture.Control.Type = 2;
                else if (fault == "builtin") fixture.Control.BuiltIn = false;
                else fixture.Control.OnAction = "UserMacro";
                fixture.Control.State = false;
                if (afterExecute)
                {
                    fixture.Timer.Tick();
                    Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                    Assert.IsTrue(fixture.Timer.Disposed);
                }
                else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.CompileCoverageProject(fixture.Project));
                Assert.AreEqual(afterExecute ? 1 : 0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void DisabledGetterCannotProveCompilationAfterChangingContextSourceOrApproval()
        {
            foreach (string fault in new[] { "identity", "mode", "source", "approval", "cancellation" })
            using (var fixture = new Fixture())
            {
                fixture.Service.CompileCoverageProject(fixture.Project);
                string revision = "reviewed";
                bool permitted = true, cancelled = false;
                int guards = 0;
                var observed = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => {
                    guards++;
                    if (cancelled) throw new OperationCanceledException();
                    if (!permitted || revision != "reviewed") throw new InvalidOperationException("Original/copy authority changed during state read.");
                });
                Application.DoEvents();
                fixture.Control.State = false;
                fixture.Control.OnRead = () => {
                    if (fault == "identity") fixture.Vbe.ActiveVBProject = new CompilerProject { Name = fixture.Project.Name };
                    else if (fault == "mode") fixture.Project.Mode = 1;
                    else if (fault == "source") revision = "edited";
                    else if (fault == "approval") permitted = false;
                    else cancelled = true;
                };
                fixture.Timer.Tick();
                Assert.IsTrue(observed.IsCompleted);
                if (fault == "cancellation") Assert.ThrowsException<OperationCanceledException>(() => observed.GetAwaiter().GetResult());
                else Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                Assert.AreEqual(2, guards);
                Assert.AreEqual(1, fixture.Control.Executions);
                Assert.IsTrue(fixture.Timer.Disposed);
            }
        }

        [STATestMethod]
        public void ObserverRejectsMissingInputsNullCommandAndDestroyedDispatcherWithoutCompilationRetry()
        {
            using (var fixture = new Fixture())
            {
                Assert.ThrowsException<ArgumentNullException>(() => fixture.Service.VerifyCoverageCompilationAsync(null, () => { }));
                Assert.ThrowsException<ArgumentNullException>(() => fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, null));
                fixture.Vbe.CommandBars.Control = null;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.CompileCoverageProject(fixture.Project));
                fixture.Vbe.CommandBars.Control = fixture.Control;
                typeof(Control).GetMethod("DestroyHandle", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fixture.Dispatcher, null);
                var observed = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => { });
                Assert.ThrowsException<InvalidOperationException>(() => observed.GetAwaiter().GetResult());
                Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void CompilationTimerDisposesOnStartupFailureAndDefaultTicksUseTheOwnedUiThread()
        {
            bool disposed = false;
            var error = Assert.ThrowsException<InvalidOperationException>(() => VbeTestExplorerService.StartCompilationTimer(() => Assert.Fail("No tick on failed startup"), timer => {
                timer.Disposed += (_, __) => disposed = true;
                throw new InvalidOperationException("Timer startup failed");
            }));
            Assert.IsTrue(disposed); Assert.AreEqual("Timer startup failed", error.Message);
            using (var fixture = new Fixture())
            {
                fixture.Control.State = false;
                fixture.Service.StartCoverageCompilationTimer = tick => (IDisposable)typeof(VbeTestExplorerService)
                    .GetMethod("StartCompilationTimer", BindingFlags.Static | BindingFlags.NonPublic, null, new[] { typeof(Action) }, null).Invoke(null, new object[] { tick });
                var task = fixture.Service.VerifyCoverageCompilationAsync(fixture.Project, () => { });
                var deadline = DateTime.UtcNow.AddSeconds(3);
                while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(5); }
                Assert.IsTrue(task.IsCompleted); Assert.IsTrue(task.GetAwaiter().GetResult());
                Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void CompilationSkipsOtherModulesAndReentrantTimerCompletionDisposesTheReturnedTimer()
        {
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents.Insert(0, new CompilerComponent { Name = "Other", CodeModule = new CompilerCode { CodePane = fixture.Pane } });
                fixture.Control.State = false;
                fixture.Service.CompileCoverageProject(fixture.Project);
                fixture.Service.StartCoverageCompilationTimer = tick => { tick(); return fixture.Timer = new ManualTimer(tick); };
                var task = fixture.Observe();
                Assert.IsTrue(task.IsCompleted); Assert.IsTrue(task.GetAwaiter().GetResult()); Assert.IsTrue(fixture.Timer.Disposed);
                fixture.Timer.Tick(); Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        [STATestMethod]
        public void MissingCoverageRuntimeCannotSelectAPaneOrExecuteCompilation()
        {
            foreach (bool otherModule in new[] { false, true })
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents.Clear();
                if (otherModule) fixture.Project.VBComponents.Add(new CompilerComponent { Name = "Other", CodeModule = new CompilerCode { CodePane = fixture.Pane } });
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.CompileCoverageProject(fixture.Project));
                StringAssert.Contains(error.Message, "runtime module is missing");
                Assert.AreEqual(0, fixture.Pane.Shows); Assert.AreEqual(0, fixture.Control.Executions);
            }
        }

        private sealed class Fixture : IDisposable
        {
            internal readonly Control Dispatcher = new Control();
            internal readonly CompilerVbe Vbe = new CompilerVbe();
            internal readonly CompilerProject Project = new CompilerProject { Name = "OwnedCoverage" };
            internal readonly CompilerPane Pane = new CompilerPane();
            internal readonly CompileControl Control = new CompileControl();
            internal readonly VbeTestExplorerService Service;
            internal ManualTimer Timer;
            internal long Clock;
            internal Fixture()
            {
                var handle = Dispatcher.Handle;
                Project.VBComponents.Add(new CompilerComponent { CodeModule = new CompilerCode { CodePane = Pane } });
                Vbe.VBProjects.Add(Project); Vbe.ActiveVBProject = Project;
                Vbe.CommandBars.Control = Control;
                Service = new VbeTestExplorerService(Vbe, Dispatcher);
                Service.CoverageCompilationClock = () => Clock;
                Service.StartCoverageCompilationTimer = tick => Timer = new ManualTimer(tick);
            }
            internal Task<bool> Observe()
            {
                var result = Service.VerifyCoverageCompilationAsync(Project, () => { });
                Application.DoEvents();
                return result;
            }
            public void Dispose() { Service.Dispose(); Dispatcher.Dispose(); }
        }
        private sealed class ManualTimer : IDisposable
        {
            private readonly Action tick;
            internal bool Disposed;
            internal ManualTimer(Action tick) { this.tick = tick; }
            internal void Tick() { tick(); }
            public void Dispose() { Disposed = true; }
        }
        public sealed class CompilerVbe
        {
            public List<CompilerProject> VBProjects { get; } = new List<CompilerProject>();
            public object ActiveVBProject { get; set; }
            public object ActiveCodePane { get; set; }
            public CompilerWindow MainWindow { get; } = new CompilerWindow();
            public CompilerCommandBars CommandBars { get; } = new CompilerCommandBars();
        }
        public sealed class CompilerProject
        {
            public string Name { get; set; }
            public int Mode { get; set; } = 2;
            public List<CompilerComponent> VBComponents { get; } = new List<CompilerComponent>();
        }
        public sealed class CompilerComponent
        {
            public string Name { get; set; } = VbaCoverageInstrumentation.ModuleName;
            public CompilerCode CodeModule { get; set; }
        }
        public sealed class CompilerCode
        {
            public CompilerPane CodePane { get; set; }
        }
        public sealed class CompilerPane
        {
            public CompilerWindow Window { get; } = new CompilerWindow();
            public int Shows;
            public void Show() { Shows++; }
        }
        public sealed class CompilerWindow
        {
            public bool Visible { get; set; }
            public int Focuses;
            public Action OnFocus;
            public void SetFocus() { Focuses++; OnFocus?.Invoke(); }
        }
        public sealed class CompilerCommandBars
        {
            public CompileControl Control { get; set; }
            public object FindControl(int type, int id) { Assert.AreEqual(1, type); Assert.AreEqual(578, id); return Control; }
        }
        public sealed class CompileControl
        {
            public int Id { get; set; } = 578;
            public int Type { get; set; } = 1;
            public bool BuiltIn { get; set; } = true;
            public string OnAction { get; set; } = "";
            public bool State = true;
            public int Reads, Executions;
            public Action OnExecute;
            public Action OnRead;
            public bool Enabled { get { Reads++; OnRead?.Invoke(); return State; } }
            public void Execute() { Executions++; OnExecute?.Invoke(); }
        }
    }
}
