using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaNativeTestExecutionHostTests
    {
        public sealed class DiagnosticProject
        {
            private string name;
            public int NameReads { get; private set; }
            public string Name { get { NameReads++; return name; } set { name = value; } }
        }
        public sealed class DiagnosticCollection { public DiagnosticProject Parent { get; set; } }
        public sealed class DiagnosticComponent
        {
            private string name;
            public int NameReads { get; private set; }
            public string Name { get { NameReads++; return name; } set { name = value; } }
            public DiagnosticCollection Collection { get; set; }
        }
        public sealed class DiagnosticBodyLine { public int this[string procedure, int kind] => 123; }
        public sealed class DiagnosticProcedureLines : DynamicObject
        {
            public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
            {
                indexes[1] = 0;
                result = (int)indexes[0] == 8 ? "OtherProcedure" : "AnotherProcedure";
                return true;
            }
        }
        public sealed class DiagnosticModule
        {
            public DiagnosticComponent Parent { get; set; }
            public DiagnosticBodyLine ProcBodyLine { get; } = new DiagnosticBodyLine();
            private readonly DiagnosticProcedureLines procedureLines = new DiagnosticProcedureLines();
            public int ProcedureReads { get; private set; }
            public DiagnosticProcedureLines ProcOfLine { get { ProcedureReads++; return procedureLines; } }
            public int PaneAcquisitions { get; private set; }
            public object CodePane
            {
                get { PaneAcquisitions++; throw new AssertFailedException("Refusal diagnostics must not acquire or activate a CodePane."); }
            }
        }
        public sealed class DiagnosticPane
        {
            public DiagnosticModule CodeModule { get; set; }
            public DiagnosticWindow Window { get; } = new DiagnosticWindow();
            public int Mutations { get; private set; }
            public bool FailSelection { get; set; }
            public void GetSelection(ref int start, ref int column, ref int end, ref int endColumn)
            {
                if (FailSelection) throw new InvalidOperationException("The editor disconnected.");
                start = 8; column = 3; end = 9; endColumn = 4;
            }
            public void Show() { Mutations++; }
            public void SetSelection(int start, int column, int end, int endColumn) { Mutations++; }
        }
        public sealed class DiagnosticEditor
        {
            public DiagnosticProject ActiveVBProject { get; set; }
            public DiagnosticPane ActiveCodePane { get; set; }
            public DiagnosticWindow MainWindow { get; } = new DiagnosticWindow { Type = 12, HWnd = 123, Visible = true };
            public DiagnosticWindow ActiveWindow { get; set; }
            public DiagnosticWindow[] Windows { get; set; } = new DiagnosticWindow[0];
        }
        public sealed class DiagnosticWindow
        {
            public int Type { get; set; }
            public bool Visible { get; set; }
            public long HWnd { get; set; }
            public int WindowState { get; set; }
            public int CaptionReads { get; private set; }
            public string Caption { get { CaptionReads++; throw new AssertFailedException("Window diagnostics must not inspect captions."); } }
        }
        public sealed class RunControl
        {
            public int Id { get; set; } = 186;
            public int Type { get; set; } = 1;
            public bool BuiltIn { get; set; } = true;
            public bool Enabled { get; set; } = true;
            public string Caption { get; set; } = "Run Sub";
            public string OnAction { get; set; } = "";
            public Action OnExecute;
            public void Execute() { OnExecute?.Invoke(); }
        }
        public sealed class SelectionWindow
        {
            public readonly List<string> Events;
            private bool visible;
            public bool AllowVisibility = true;
            public int Type { get; set; }
            public long HWnd { get; set; }
            public string Caption { get; set; } = "Owned support (Code)";
            public Action OnFocus;
            public SelectionWindow(List<string> events) { Events = events; }
            public bool Visible { get => visible; set { Events.Add("Visible"); visible = value && AllowVisibility; } }
            public void SetFocus() { Events.Add(Type == 12 ? "MainFocus" : "Focus"); Assert.IsTrue(Visible); OnFocus?.Invoke(); }
        }
        public sealed class SelectionEditor
        {
            public readonly List<string> Events = new List<string>();
            public object VBProjects { get; set; }
            public NativeCommandBars CommandBars { get; set; }
            public object Windows { get; set; } = new object[0];
            public SelectionWindow MainWindow { get; }
            public object ActiveVBProject { get; set; }
            public object ActiveWindow { get; set; }
            private SelectionPane pane;
            public SelectionEditor() { MainWindow = new SelectionWindow(Events) { Type = 12 }; }
            public SelectionPane ActiveCodePane { get => pane; set { Events.Add("ActivePane"); pane = value; } }
        }
        public sealed class SelectionModule
        {
            public SelectionEditor Editor { get; }
            public SelectionPane Pane { get; }
            public SelectionBodyLine ProcBodyLine { get; }
            public SelectionModule(SelectionEditor editor) { Editor = editor; Pane = new SelectionPane(editor) { CodeModule = this }; ProcBodyLine = new SelectionBodyLine(Pane); }
            public SelectionPane CodePane
            {
                get { Editor.Events.Add("CodePane"); Assert.IsTrue(Editor.MainWindow.Visible); return Pane; }
            }
        }
        public sealed class SelectionBodyLine
        {
            private readonly SelectionPane pane;
            public int? Override { get; set; }
            public SelectionBodyLine(SelectionPane pane) { this.pane = pane; }
            public int this[string procedure, int kind]
            {
                get { Assert.AreEqual(VbaTestRuntimeSource.PendingProcedure, procedure); Assert.AreEqual(0, kind); return Override ?? pane.StartLine; }
            }
        }
        public sealed class SelectionPane
        {
            public SelectionEditor Editor { get; }
            public SelectionWindow Window { get; }
            public object CodeModule { get; set; }
            public bool Materialize = true;
            public int StartLine, StartColumn, EndLine, EndColumn;
            public SelectionPane(SelectionEditor editor) { Editor = editor; Window = new SelectionWindow(editor.Events); }
            public void Show()
            {
                Editor.Events.Add("Show"); Assert.IsTrue(Editor.MainWindow.Visible);
                Window.Visible = true;
                if (Materialize) Window.HWnd = 123;
            }
            public void SetSelection(int start, int column, int end, int endColumn)
            {
                Editor.Events.Add("Selection"); StartLine = start; StartColumn = column; EndLine = end; EndColumn = endColumn;
            }
            public void GetSelection(ref int start, ref int column, ref int end, ref int endColumn)
            { start = StartLine; column = StartColumn; end = EndLine; endColumn = EndColumn; }
        }
        private sealed class Probe : VbaNativeTestExecutionHost.IProbe
        {
            internal Action OnPrepare, OnRevalidate, OnExecute, OnReadMode;
            internal int Mode = 2, Invocations;
            public void RequireOwner(object vbe) { }
            public object Prepare(object vbe, object project, string source) { OnPrepare?.Invoke(); return new object(); }
            public void Revalidate(object vbe, object project, object prepared) { OnRevalidate?.Invoke(); }
            public void Execute(object prepared) { Invocations++; OnExecute?.Invoke(); }
            public int ReadMode(object project) { OnReadMode?.Invoke(); return Mode; }
        }
        private sealed class Cleanup : IDisposable
        {
            private readonly Action action;
            internal Cleanup(Action action) { this.action = action; }
            public void Dispose() { action(); }
        }
        private sealed class Fixture : IDisposable
        {
            internal readonly object Project = new object();
            internal readonly VbaTestResultSink Sink = new VbaTestResultSink(ReferenceEquals);
            internal readonly Control Dispatcher = new Control();
            internal readonly Probe Native = new Probe();
            internal readonly VbaTestDescriptor Test = new VbaTestDescriptor { Id = "test", Module = "Tests", Procedure = "Check", Kind = "Sub" };
            internal readonly VbaTestCatalog Catalog = new VbaTestCatalog
            {
                Project = new VbaTestProjectSnapshot { Id = "project", HostPath = @"C:\fixture\owned.swp", Revision = "revision",
                    Modules = new[] { new VbaTestModuleSnapshot { Name = VbaTestRuntimeSource.ModuleName, Source = "reviewed source" } } }
            };
            internal readonly VbaNativeTestExecutionHost Host;
            internal Action Queued, Tick, GuardAction, ValidateAction;
            internal int Guards;
            internal Fixture()
            {
                Host = new VbaNativeTestExecutionHost(new object(), Dispatcher, Sink, _ => Project,
                    _ => ValidateAction?.Invoke(), () => { Guards++; GuardAction?.Invoke(); }, _ => "signature", () => "run")
                { Probe = Native, Post = action => Queued = action, StartPolling = action => { Tick = action; return new Cleanup(() => Tick = null); } };
                Native.OnExecute = () => Publish();
            }
            internal void Publish(string status = "Passed", string message = "", int error = 0)
            {
                var runtime = new VbaTestRuntime();
                var job = (object[])runtime.Request(VbaTestRuntimeSource.Version, "signature");
                runtime.Publish((string)job[0], (string)job[4], (string)job[5], (string)job[6], status, message, error);
            }
            internal Task<VbaTestResult> Start()
            {
                var task = Host.InvokeAsync(Catalog, Test, "Test");
                Queued(); Queued = null;
                return task;
            }
            public void Dispose() { Host.Dispose(); Sink.Dispose(); Dispatcher.Dispose(); }
        }

        [TestMethod]
        public void NativeReturnAndCallbackVerdictCompleteOneAttemptAndRevokeFactoryExposure()
        {
            using (var f = new Fixture())
            {
                var result = f.Start().GetAwaiter().GetResult();
                Assert.AreEqual(VbaTestOutcome.Passed, result.Outcome); Assert.AreSame(f.Test, result.Test);
                Assert.AreEqual("Test", result.Phase); Assert.AreEqual(1, f.Native.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(2, f.Native.Invocations);
            }
        }

        [TestMethod]
        public void DispatchReservesNativeSelectionAcrossReentrantAuthorityAndRuntimeCallbacks()
        {
            Assert.IsFalse(VbeDebugInspection.IsActive);
            using (var f = new Fixture())
            {
                var stages = new List<string>();
                Action<string> observe = stage =>
                {
                    Assert.IsTrue(VbeDebugInspection.IsActive, stage);
                    using (new VbeDebugInspection()) Assert.IsTrue(VbeDebugInspection.IsActive, "Nested " + stage);
                    Assert.IsTrue(VbeDebugInspection.IsActive, "Nested disposal must retain dispatch ownership.");
                    stages.Add(stage);
                };
                f.Native.OnPrepare = () => observe("Prepare");
                f.ValidateAction = () => { if (f.Guards > 1) observe("Authority"); };
                f.Native.OnRevalidate = () => observe("Revalidate");
                f.GuardAction = () => { if (f.Guards > 1) observe("Guard"); };
                f.Native.OnExecute = () => { observe("Execute"); f.Publish(); };
                f.Native.OnReadMode = () => observe("Observe");
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                foreach (string stage in new[] { "Prepare", "Authority", "Revalidate", "Guard", "Execute", "Observe" })
                    Assert.IsTrue(stages.Contains(stage), stage + " was not exercised.");
                Assert.AreEqual(1, f.Native.Invocations);
                Assert.IsFalse(VbeDebugInspection.IsActive, "Successful dispatch must release selection ownership.");
            }
        }

        [DataTestMethod]
        [DataRow("Prepare", false)]
        [DataRow("Authority", false)]
        [DataRow("Revalidate", false)]
        [DataRow("Guard", false)]
        [DataRow("Execute", true)]
        [DataRow("Observe", true)]
        public void ReentrantDispatchFailureReleasesSelectionAndRetainsAuthorityAndUncertainty(string stage, bool invoked)
        {
            Assert.IsFalse(VbeDebugInspection.IsActive);
            using (var f = new Fixture())
            {
                Action refuse = () =>
                {
                    Assert.IsTrue(VbeDebugInspection.IsActive, stage);
                    throw new InvalidOperationException("Reentrant " + stage + " authority refusal.");
                };
                if (stage == "Prepare") f.Native.OnPrepare = refuse;
                if (stage == "Authority") f.ValidateAction = () => { if (f.Guards == 3) refuse(); };
                if (stage == "Revalidate") f.Native.OnRevalidate = refuse;
                if (stage == "Guard") f.GuardAction = () => { if (f.Guards == 4) refuse(); };
                if (stage == "Execute") f.Native.OnExecute = refuse;
                if (stage == "Observe") f.Native.OnReadMode = refuse;
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                StringAssert.Contains(error.Message, stage);
                Assert.AreEqual(invoked, error.Uncertain);
                Assert.AreEqual(invoked ? 1 : 0, f.Native.Invocations);
                Assert.IsFalse(VbeDebugInspection.IsActive, "A refusal must release selection ownership.");
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void PendingNativeCompletionReleasesDispatchSelectionBeforeOwnerPolling()
        {
            using (var f = new Fixture())
            {
                f.Native.Mode = 0;
                f.Native.OnReadMode = () => Assert.IsTrue(VbeDebugInspection.IsActive);
                var task = f.Start();
                Assert.IsFalse(task.IsCompleted);
                Assert.IsFalse(VbeDebugInspection.IsActive, "A pending callback must not retain a thread-static inspection.");
                f.Native.OnReadMode = () => Assert.IsFalse(VbeDebugInspection.IsActive);
                f.Native.Mode = 2;
                f.Tick();
                Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
                Assert.IsFalse(VbeDebugInspection.IsActive);
            }
        }
        [TestMethod]
        public void AReceivedVerdictWaitsForTheLiveProjectToReturnToDesignMode()
        {
            using (var f = new Fixture())
            {
                f.Native.Mode = 0; var task = f.Start();
                Assert.IsFalse(task.IsCompleted); Assert.IsNotNull(f.Tick);
                Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.InvokeAsync(f.Catalog, f.Test, "Test"));
                f.Native.Mode = 2; f.Tick();
                Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
                Assert.IsNull(f.Tick); Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [TestMethod]
        public void CallbackDoesNotCompleteTheTaskBeforeTheNativeCommandReturns()
        {
            using (var f = new Fixture())
            {
                Task<VbaTestResult> task = f.Host.InvokeAsync(f.Catalog, f.Test, "Test");
                f.Native.OnExecute = () => { f.Publish(); Assert.IsFalse(task.IsCompleted); };
                f.Queued();
                Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
            }
        }

        [TestMethod]
        public void MissingCallbackAtTheObservationDeadlineIsUnknownAndCannotBeRetried()
        {
            using (var f = new Fixture())
            {
                f.Native.OnExecute = () => { }; f.Host.VerificationTimeout = TimeSpan.Zero;
                var task = f.Start();
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.InvokeAsync(f.Catalog, f.Test, "Test"));
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
                Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("throw")]
        [DataRow("break")]
        [DataRow("callback")]
        public void DispatchedNativeFailuresNeverProducePassedOrRetry(string failure)
        {
            using (var f = new Fixture())
            {
                f.Native.OnExecute = () =>
                {
                    if (failure == "callback")
                    {
                        var runtime = new VbaTestRuntime();
                        try { runtime.Request(VbaTestRuntimeSource.Version, "foreign signature"); } catch (InvalidOperationException) { }
                    }
                    else { f.Publish(); if (failure == "throw") throw new InvalidOperationException("Native dispatch failed after VBA ran."); else f.Native.Mode = 1; }
                };
                var task = f.Start();
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.AreEqual(1, f.Native.Invocations);
                Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.InvokeAsync(f.Catalog, f.Test, "Test"));
            }
        }

        [TestMethod]
        public void AuthorityIsRecheckedAfterPaneSelectionAndAtTheLastDispatchBoundary()
        {
            using (var f = new Fixture())
            {
                bool refuse = false;
                f.Native.OnPrepare = () => refuse = true;
                f.GuardAction = () => { if (refuse) throw new InvalidOperationException("Execution approval changed."); };
                var task = f.Start();
                Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.AreEqual(0, f.Native.Invocations);
                refuse = false; f.Native.OnPrepare = null;
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
            }
        }

        [TestMethod]
        public void FinalGuardRefusalCancelsTheArmedAttemptWithoutExposingAJob()
        {
            using (var f = new Fixture())
            {
                f.GuardAction = () => { if (f.Guards == 4) throw new InvalidOperationException("Privacy changed after native control inspection."); };
                var task = f.Start();
                Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
                f.GuardAction = null;
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
            }
        }

        [TestMethod]
        public void ChangedNativeSelectionBeforeDispatchIsARefusal()
        {
            using (var f = new Fixture())
            {
                f.Native.OnRevalidate = () => { throw new InvalidOperationException("Another code pane now owns selection."); };
                var task = f.Start();
                Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.AreEqual(0, f.Native.Invocations);
            }
        }

        [TestMethod]
        public void HiddenVbeIsShownBeforePaneAcquisitionSelectionAndFocusAndDispatchesOnce()
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor();
                var module = new SelectionModule(editor);
                f.Native.OnPrepare = () => Assert.AreSame(module.Pane,
                    VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, shown =>
                    {
                        Assert.AreSame(module.Pane, shown);
                        CollectionAssert.AreEqual(new[] { "Visible", "MainFocus", "CodePane", "Show", "Visible" }, editor.Events);
                        Assert.IsNull(editor.ActiveCodePane);
                        Assert.AreEqual(0, module.Pane.StartLine);
                    }, ReferenceEquals));
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                CollectionAssert.AreEqual(new[] { "Visible", "MainFocus", "CodePane", "Show", "Visible", "ActivePane", "Selection", "Focus" }, editor.Events);
                Assert.AreSame(module.Pane, editor.ActiveCodePane);
                Assert.AreEqual(146, module.Pane.StartLine); Assert.AreEqual(1, module.Pane.StartColumn);
                Assert.AreEqual(146, module.Pane.EndLine); Assert.AreEqual(1, module.Pane.EndColumn);
                Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("mainWindow")]
        [DataRow("paneType")]
        [DataRow("paneVisibility")]
        [DataRow("paneModule")]
        public void InvalidNativeWindowsRefuseBeforePaneFocusAndNeverDispatch(string failure)
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor();
                var module = new SelectionModule(editor);
                if (failure == "mainWindow") editor.MainWindow.AllowVisibility = false;
                if (failure == "paneType") module.Pane.Window.Type = 15;
                if (failure == "paneVisibility") module.Pane.Window.AllowVisibility = false;
                if (failure == "paneModule") module.Pane.CodeModule = new object();
                f.Native.OnPrepare = () => VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, identity: ReferenceEquals);
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain);
                Assert.IsFalse(editor.Events.Contains("Focus"));
                if (failure == "mainWindow") CollectionAssert.AreEqual(new[] { "Visible" }, editor.Events);
                else CollectionAssert.AreEqual(new[] { "Visible", "MainFocus", "CodePane", "Show", "Visible", "ActivePane", "Selection" }, editor.Events);
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void ZeroCodeWindowHandleAllowsOneFocusAndOneDispatchOnlyWithExactActiveContext()
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor { ActiveVBProject = f.Project };
                var module = new SelectionModule(editor);
                module.Pane.Materialize = false;
                f.Native.OnPrepare = () =>
                {
                    VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, identity: ReferenceEquals);
                    editor.ActiveWindow = module.Pane.Window;
                };
                f.Native.OnRevalidate = () => VbaNativeTestExecutionHost.ValidateNativePaneContext(editor, f.Project,
                    module, module.Pane, 146, ReferenceEquals);
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(0L, module.Pane.Window.HWnd);
                Assert.AreEqual(1, editor.Events.FindAll(item => item == "MainFocus").Count);
                Assert.AreEqual(1, editor.Events.FindAll(item => item == "Focus").Count);
                Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("projectNull")]
        [DataRow("projectChanged")]
        [DataRow("windowNull")]
        [DataRow("windowMain")]
        [DataRow("windowTool")]
        [DataRow("windowHidden")]
        [DataRow("windowType")]
        [DataRow("paneNull")]
        [DataRow("paneChanged")]
        [DataRow("moduleChanged")]
        [DataRow("selectionLine")]
        [DataRow("selectionColumn")]
        [DataRow("wrapperBody")]
        public void LastActivePaneCannotDispatchWithoutExactProjectFocusedCodeWindowAndWrapperSelection(string failure)
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor { ActiveVBProject = f.Project };
                var module = new SelectionModule(editor);
                module.Pane.Materialize = false;
                f.Native.OnPrepare = () =>
                {
                    VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, identity: ReferenceEquals);
                    editor.ActiveWindow = module.Pane.Window;
                    switch (failure)
                    {
                        case "projectNull": editor.ActiveVBProject = null; break;
                        case "projectChanged": editor.ActiveVBProject = new object(); break;
                        case "windowNull": editor.ActiveWindow = null; break;
                        case "windowMain": editor.ActiveWindow = editor.MainWindow; break;
                        case "windowTool": editor.ActiveWindow = new SelectionWindow(editor.Events) { Type = 15 }; break;
                        case "windowHidden": module.Pane.Window.Visible = false; break;
                        case "windowType": module.Pane.Window.Type = 15; break;
                        case "paneNull": editor.ActiveCodePane = null; break;
                        case "paneChanged": editor.ActiveCodePane = new SelectionPane(editor); break;
                        case "moduleChanged": module.Pane.CodeModule = new object(); break;
                        case "selectionLine": module.Pane.EndLine = 147; break;
                        case "selectionColumn": module.Pane.StartColumn = 2; break;
                        case "wrapperBody": module.ProcBodyLine.Override = 147; break;
                        default: Assert.Fail("Unknown refusal scenario."); break;
                    }
                };
                int validations = 0;
                f.Native.OnRevalidate = () =>
                {
                    validations++;
                    VbaNativeTestExecutionHost.ValidateNativePaneContext(editor, f.Project, module, module.Pane, 146, ReferenceEquals);
                };
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain);
                Assert.AreEqual(1, validations, "Invalid native context must refuse before arming the attempt.");
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.AreEqual(1, editor.Events.FindAll(item => item == "MainFocus").Count);
                Assert.AreEqual(1, editor.Events.FindAll(item => item == "Focus").Count);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NativeFallbackRunsOnlyWhenComFocusFailedAndMustRestoreExactComContext(bool comAlreadyFocused)
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor { ActiveVBProject = f.Project };
                editor.MainWindow.HWnd = 1;
                var module = new SelectionModule(editor); module.Pane.Materialize = false;
                var windows = new VbaNativeTestWindowFocusTests.Windows();
                VbaNativeTestWindowFocus.Target target = null;
                windows.FocusAction = () => { editor.ActiveWindow = module.Pane.Window; editor.ActiveCodePane = module.Pane; windows.Focused = new IntPtr(4); };
                f.Native.OnPrepare = () =>
                {
                    VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, identity: ReferenceEquals);
                    editor.ActiveWindow = comAlreadyFocused ? (object)module.Pane.Window : editor.MainWindow;
                    if (!comAlreadyFocused) editor.ActiveCodePane = null;
                    target = VbaNativeTestExecutionHost.EnsureNativePaneFocus(editor, f.Project, module, module.Pane, 146, windows, ReferenceEquals);
                };
                f.Native.OnRevalidate = () =>
                {
                    VbaNativeTestExecutionHost.ValidateNativePaneContext(editor, f.Project, module, module.Pane, 146, ReferenceEquals);
                    VbaNativeTestExecutionHost.ValidateRecoveredNativeFocus(editor, module.Pane, windows, target);
                };
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(comAlreadyFocused ? 0 : 1, windows.Activations);
                Assert.AreEqual(comAlreadyFocused ? 0 : 1, windows.Focuses);
                Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("nativeAmbiguous")]
        [DataRow("nativeProcess")]
        [DataRow("nativeThread")]
        [DataRow("nativeParent")]
        [DataRow("nativeClass")]
        [DataRow("nativeReused")]
        [DataRow("focusFailed")]
        [DataRow("comWindow")]
        [DataRow("comPane")]
        [DataRow("comModule")]
        [DataRow("comSelection")]
        [DataRow("comProject")]
        [DataRow("laterFocus")]
        [DataRow("laterMain")]
        [DataRow("approval")]
        [DataRow("source")]
        public void FailedFallbackOrPumpedContextNeverRunsOrExposesTheCallback(string change)
        {
            using (var f = new Fixture())
            {
                var editor = new SelectionEditor { ActiveVBProject = f.Project };
                editor.MainWindow.HWnd = 1;
                var module = new SelectionModule(editor); module.Pane.Materialize = false;
                var windows = new VbaNativeTestWindowFocusTests.Windows();
                VbaNativeTestWindowFocus.Target target = null;
                bool prepared = false;
                if (change == "nativeAmbiguous") { windows.Add(5, 2, "VbaWindow"); windows.Captions[new IntPtr(5)] = "Owned support (Code)"; }
                if (change == "nativeProcess") windows.Items[new IntPtr(3)].Process = 9;
                if (change == "nativeThread") windows.Items[new IntPtr(3)].Thread = 9;
                if (change == "nativeParent") windows.Items[new IntPtr(3)].Parent = new IntPtr(1);
                if (change == "nativeClass") windows.Items[new IntPtr(3)].Class = "ToolWindow";
                if (change == "nativeReused") windows.Activated = () => windows.Items[new IntPtr(3)].Thread = 9;
                windows.FocusAction = () =>
                {
                    editor.ActiveWindow = module.Pane.Window; editor.ActiveCodePane = module.Pane;
                    switch (change)
                    {
                        case "focusFailed": windows.Focused = IntPtr.Zero; break;
                        case "comWindow": editor.ActiveWindow = editor.MainWindow; break;
                        case "comPane": editor.ActiveCodePane = null; break;
                        case "comModule": module.Pane.CodeModule = new object(); break;
                        case "comSelection": module.Pane.EndColumn = 2; break;
                        case "comProject": editor.ActiveVBProject = new object(); break;
                    }
                };
                f.Native.OnPrepare = () =>
                {
                    VbaNativeTestExecutionHost.PrepareNativePane(editor, module, 146, identity: ReferenceEquals);
                    editor.ActiveWindow = editor.MainWindow; editor.ActiveCodePane = null;
                    target = VbaNativeTestExecutionHost.EnsureNativePaneFocus(editor, f.Project, module, module.Pane, 146, windows, ReferenceEquals);
                    prepared = true;
                    if (change == "laterFocus") windows.Focused = new IntPtr(1);
                    if (change == "laterMain") editor.MainWindow.HWnd = 5;
                };
                f.GuardAction = () => { if (prepared && change == "approval") throw new InvalidOperationException("Approval changed during focus."); };
                f.ValidateAction = () => { if (prepared && change == "source") throw new InvalidOperationException("Source changed during focus."); };
                f.Native.OnRevalidate = () =>
                {
                    VbaNativeTestExecutionHost.ValidateNativePaneContext(editor, f.Project, module, module.Pane, 146, ReferenceEquals);
                    VbaNativeTestExecutionHost.ValidateRecoveredNativeFocus(editor, module.Pane, windows, target);
                };
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain); Assert.AreEqual(0, f.Native.Invocations);
                Assert.IsTrue(windows.Activations <= 1 && windows.Focuses <= 1);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void PostShowWindowDiagnosticsObserveZeroComHandlesAndOwnedNativeClassesWithoutCaptionsOrCode()
        {
            var project = new DiagnosticProject { Name = "Private project name" };
            var module = new DiagnosticModule { Parent = new DiagnosticComponent { Name = "Private module name",
                Collection = new DiagnosticCollection { Parent = project } } };
            var expectedPane = new DiagnosticPane { CodeModule = module };
            var editor = new DiagnosticEditor { ActiveWindow = expectedPane.Window, ActiveCodePane = expectedPane,
                Windows = new[] { expectedPane.Window } };
            string observation = VbaNativeTestExecutionHost.DescribeNativeWindows(editor, expectedPane, module,
                ReferenceEquals, main => { Assert.AreEqual(new IntPtr(123), main); return "{HWnd=456,PID=1234,Class=VbaWindow,Parent=123,Visible=True}"; });
            StringAssert.Contains(observation, "paneWindow={Type=0,Visible=False,HWnd=0,State=0}");
            StringAssert.Contains(observation, "activeWindowMatches=True; activePaneAvailable=True; activePaneMatches=True; activeModuleMatches=True");
            StringAssert.Contains(observation, "selection=8:3-9:4");
            StringAssert.Contains(observation, "Class=VbaWindow");
            Assert.AreEqual(0, module.PaneAcquisitions + module.ProcedureReads + module.Parent.NameReads + project.NameReads);
            Assert.AreEqual(0, expectedPane.Window.CaptionReads + editor.MainWindow.CaptionReads);
            Assert.AreEqual(0, expectedPane.Mutations);
            Assert.IsFalse(observation.Contains("Private project name") || observation.Contains("Private module name"));
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void PaneRefusalDiagnosticsReadNamesAndProceduresOnlyWithinTheSelectedProject(bool withinSelectedProject)
        {
            var expectedProject = new DiagnosticProject { Name = "Owned project" };
            var otherProject = new DiagnosticProject { Name = "Other project" };
            var expectedModule = new DiagnosticModule { Parent = new DiagnosticComponent { Name = VbaTestRuntimeSource.ModuleName,
                Collection = new DiagnosticCollection { Parent = expectedProject } } };
            var otherModule = new DiagnosticModule { Parent = new DiagnosticComponent { Name = "OtherModule",
                Collection = new DiagnosticCollection { Parent = withinSelectedProject ? expectedProject : otherProject } } };
            var expectedPane = new DiagnosticPane { CodeModule = expectedModule };
            var actualPane = new DiagnosticPane { CodeModule = otherModule };
            var editor = new DiagnosticEditor { ActiveVBProject = withinSelectedProject ? expectedProject : otherProject, ActiveCodePane = actualPane };
            string diagnostic = VbaNativeTestExecutionHost.DescribeNativeSelection(editor, expectedProject,
                expectedModule, expectedPane, 123, ReferenceEquals);
            StringAssert.Contains(diagnostic, "project=Owned project");
            StringAssert.Contains(diagnostic, "procedure=" + VbaTestRuntimeSource.PendingProcedure);
            StringAssert.Contains(diagnostic, "selection=123:1-123:1");
            StringAssert.Contains(diagnostic, "currentBodyLine=123");
            StringAssert.Contains(diagnostic, "paneMatches=False");
            Assert.IsFalse(diagnostic.Contains("currentModulePaneMatches"));
            Assert.AreEqual(0, expectedModule.PaneAcquisitions + otherModule.PaneAcquisitions);
            StringAssert.Contains(diagnostic, "selection=8:3-9:4");
            Assert.AreEqual(0, otherProject.NameReads);
            if (withinSelectedProject)
            {
                StringAssert.Contains(diagnostic, "activeProject=Owned project; activeProjectMatches=True");
                StringAssert.Contains(diagnostic, "module=OtherModule; moduleMatches=False");
                StringAssert.Contains(diagnostic, "paneProject=Owned project; paneProjectMatches=True");
                StringAssert.Contains(diagnostic, "startProcedure=OtherProcedure (kind=0)");
                StringAssert.Contains(diagnostic, "endProcedure=AnotherProcedure (kind=0)");
                Assert.AreEqual(2, otherModule.ProcedureReads);
                Assert.AreEqual(1, otherModule.Parent.NameReads);
            }
            else
            {
                StringAssert.Contains(diagnostic, "activeProject=redacted(outside selected project); activeProjectMatches=False");
                StringAssert.Contains(diagnostic, "module=redacted(outside selected project); moduleMatches=False");
                StringAssert.Contains(diagnostic, "paneProject=redacted(outside selected project); paneProjectMatches=False");
                StringAssert.Contains(diagnostic, "startProcedure=redacted(outside selected project)");
                StringAssert.Contains(diagnostic, "endProcedure=redacted(outside selected project)");
                Assert.AreEqual(0, otherModule.ProcedureReads);
                Assert.AreEqual(0, otherModule.Parent.NameReads);
                Assert.IsFalse(diagnostic.Contains("Other project") || diagnostic.Contains("OtherModule")
                    || diagnostic.Contains("OtherProcedure") || diagnostic.Contains("AnotherProcedure"));
            }
            Assert.AreSame(actualPane, editor.ActiveCodePane);
            Assert.AreEqual(0, expectedPane.Mutations + actualPane.Mutations);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void MissingOrDisconnectedPaneDiagnosticsPreserveTheUndispatchedRefusal(bool disconnected)
        {
            using (var f = new Fixture())
            {
                var project = new DiagnosticProject { Name = "Owned project" };
                var module = new DiagnosticModule { Parent = new DiagnosticComponent { Name = VbaTestRuntimeSource.ModuleName,
                    Collection = new DiagnosticCollection { Parent = project } } };
                var pane = new DiagnosticPane { CodeModule = module, FailSelection = disconnected };
                var editor = new DiagnosticEditor { ActiveVBProject = project, ActiveCodePane = disconnected ? pane : null };
                f.Native.OnRevalidate = () => throw new InvalidOperationException("The support code pane no longer owns the native run selection."
                    + VbaNativeTestExecutionHost.DescribeNativeSelection(editor, project, module, pane, 123, ReferenceEquals));
                var task = f.Start();
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain);
                StringAssert.Contains(refusal.Message, "The support code pane no longer owns the native run selection.");
                StringAssert.Contains(refusal.Message, "selection=unavailable(");
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.AreEqual(0, pane.Mutations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [DataTestMethod]
        [DataRow("id")]
        [DataRow("type")]
        [DataRow("builtIn")]
        [DataRow("onAction")]
        [DataRow("enabled")]
        [DataRow("caption")]
        [DataRow("changedCaption")]
        public void ReplacedOrReassignedRunControlIsRefusedBeforeNativeExecution(string changed)
        {
            using (var f = new Fixture())
            {
                var control = new RunControl();
                VbaNativeTestExecutionHost.ValidateNativeRunControl(control, "Run Sub");
                if (changed == "id") control.Id = 187;
                if (changed == "type") control.Type = 2;
                if (changed == "builtIn") control.BuiltIn = false;
                if (changed == "onAction") control.OnAction = "ForeignMacro";
                if (changed == "enabled") control.Enabled = false;
                if (changed == "caption") control.Caption = "Foreign command";
                if (changed == "changedCaption") control.Caption = "Exécuter Sub";
                f.Native.OnRevalidate = () => VbaNativeTestExecutionHost.ValidateNativeRunControl(control, "Run Sub");
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain);
                StringAssert.Contains(refusal.Message, "verified built-in native Run Sub command");
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [DataTestMethod]
        [DataRow("queued")]
        [DataRow("selection")]
        public void SourceChangesBeforeDispatchAreRefusedWithoutExecutingOrExposingAJob(string when)
        {
            using (var f = new Fixture())
            {
                string liveRevision = f.Catalog.Project.Revision;
                f.ValidateAction = () =>
                {
                    if (liveRevision != f.Catalog.Project.Revision)
                        throw new VbaTestInvocationException("The project source changed since discovery.", false);
                };
                if (when == "selection") f.Native.OnPrepare = () => liveRevision = "changed source";
                var task = f.Host.InvokeAsync(f.Catalog, f.Test, "Test");
                if (when == "queued") liveRevision = "changed source";
                f.Queued(); f.Queued = null;
                var refusal = Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult());
                Assert.IsFalse(refusal.Uncertain);
                StringAssert.Contains(refusal.Message, "source changed");
                Assert.AreEqual(0, f.Native.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
                liveRevision = f.Catalog.Project.Revision; f.Native.OnPrepare = null;
                Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(1, f.Native.Invocations);
            }
        }

        [TestMethod]
        public void DisconnectingAStartedPendingCallRevokesItsCallbackAndMarksItsOutcomeUnknown()
        {
            using (var f = new Fixture())
            {
                f.Native.Mode = 0; var task = f.Start(); f.Host.Dispose();
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
                Assert.IsNull(f.Tick);
            }
        }
        public sealed class NativeProject
        {
            public string Name { get; set; } = "Owned project";
            public int Mode { get; set; } = 2;
            public List<NativeComponent> VBComponents { get; } = new List<NativeComponent>();
        }
        public sealed class NativeComponent
        {
            public string Name { get; set; } = VbaTestRuntimeSource.ModuleName;
            public int Type { get; set; } = 1;
            public NativeModule CodeModule { get; set; }
        }
        public sealed class NativeLines
        {
            public string Source = "Public Sub " + VbaTestRuntimeSource.PendingProcedure + "()";
            public string this[int line, int count] => Source;
        }
        public sealed class NativeModule
        {
            public NativeComponent Parent { get; set; }
            public int CountOfLines { get; set; } = 1;
            public NativeLines Lines { get; } = new NativeLines();
            public SelectionPane CodePane { get; set; }
            public SelectionBodyLine ProcBodyLine { get; set; }
        }
        public sealed class NativeCommandBars
        {
            public RunControl Control { get; } = new RunControl();
            public object FindControl(int type, int id) { Assert.AreEqual(1, type); Assert.AreEqual(186, id); return Control; }
        }
        private static void RunOnSta(Action action)
        {
            Exception error = null;
            var thread = new System.Threading.Thread(() => { try { action(); } catch (Exception caught) { error = caught; } });
            thread.SetApartmentState(System.Threading.ApartmentState.STA); thread.Start(); thread.Join();
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
        private sealed class NativeFixture
        {
            internal readonly SelectionEditor Editor = new SelectionEditor();
            internal readonly NativeProject Project = new NativeProject();
            internal readonly VbaNativeTestWindowFocusTests.Windows Windows = new VbaNativeTestWindowFocusTests.Windows();
            internal readonly NativeModule Module;
            internal readonly VbaNativeTestExecutionHost.NativeProbe Probe;
            internal NativeFixture()
            {
                var pane = new SelectionPane(Editor); pane.StartLine = 1; pane.EndLine = 1; pane.StartColumn = pane.EndColumn = 1;
                Module = new NativeModule { CodePane = pane, ProcBodyLine = new SelectionBodyLine(pane) };
                pane.CodeModule = Module;
                var component = new NativeComponent { CodeModule = Module }; Module.Parent = component;
                Project.VBComponents.Add(component);
                Editor.MainWindow.HWnd = 1;
                Editor.ActiveVBProject = Project; Editor.ActiveWindow = pane.Window;
                Editor.VBProjects = new object[] { new object(), Project };
                Editor.CommandBars = new NativeCommandBars();
                Probe = new VbaNativeTestExecutionHost.NativeProbe(Windows, ReferenceEquals);
            }
            internal object Prepare() => Probe.Prepare(Editor, Project, Module.Lines.Source);
        }

        [TestMethod]
        public void NativeProbePreparesRevalidatesAndExecutesTheExactBuiltInCommandOnce()
        {
            RunOnSta(() =>
            {
                var f = new NativeFixture(); int executions = 0;
                f.Editor.CommandBars.Control.OnExecute = () => executions++;
                var plan = f.Prepare(); f.Probe.Revalidate(f.Editor, f.Project, plan); f.Probe.Execute(plan);
                Assert.AreEqual(1, executions); Assert.AreEqual(2, f.Probe.ReadMode(f.Project));
                f.Editor.VBProjects = new object[0];
                Assert.ThrowsException<InvalidOperationException>(() => f.Probe.Revalidate(f.Editor, f.Project, plan));
                f.Editor.VBProjects = new object[] { f.Project }; f.Project.Mode = 0;
                Assert.ThrowsException<InvalidOperationException>(() => f.Probe.Revalidate(f.Editor, f.Project, plan));
                f.Project.Mode = 2; f.Module.Lines.Source = "Changed";
                Assert.ThrowsException<InvalidOperationException>(() => f.Probe.Revalidate(f.Editor, f.Project, plan));
                Assert.AreEqual(1, executions);
            });
        }

        [TestMethod]
        public void NativeProbeRejectsUnavailableOwnerAndSupportBeforeDispatch()
        {
            RunOnSta(() =>
            {
                foreach (string fault in new[] { "zero", "pid", "thread", "duplicate", "missing", "type", "source", "negative", "huge", "empty", "wrapper", "main", "pane", "selection" })
                {
                    var f = new NativeFixture(); string expected = f.Module.Lines.Source;
                    switch (fault)
                    {
                        case "zero": f.Editor.MainWindow.HWnd = 0; break;
                        case "pid": f.Windows.Items[new IntPtr(1)].Process = 9; break;
                        case "thread": f.Windows.Items[new IntPtr(1)].Thread = 9; break;
                        case "duplicate": f.Project.VBComponents.Add(f.Project.VBComponents[0]); break;
                        case "missing": f.Project.VBComponents.Clear(); break;
                        case "type": f.Project.VBComponents[0].Type = 2; break;
                        case "source": expected = "Other source"; break;
                        case "negative": f.Module.CountOfLines = -1; break;
                        case "huge": f.Module.CountOfLines = 200001; break;
                        case "empty": f.Module.CountOfLines = 0; break;
                        case "wrapper": expected = f.Module.Lines.Source = "Private Sub Wrong()"; break;
                        case "main": f.Editor.MainWindow.AllowVisibility = false; break;
                        case "pane": f.Module.CodePane.Window.Type = 1; break;
                        case "selection": f.Module.CodePane.Window.OnFocus = () => f.Module.CodePane.StartColumn = 2; break;
                    }
                    Assert.ThrowsException<InvalidOperationException>(() => f.Probe.Prepare(f.Editor, f.Project, expected), fault);
                    Assert.AreEqual(0, f.Windows.Activations, fault);
                }
                var valid = new NativeFixture(); var plan = valid.Prepare();
                valid.Editor.ActiveVBProject = new object();
                Assert.ThrowsException<InvalidOperationException>(() => valid.Probe.Revalidate(valid.Editor, valid.Project, plan));
            });
            var mta = new NativeFixture();
            Exception refused = null;
            var thread = new System.Threading.Thread(() => { try { mta.Probe.RequireOwner(mta.Editor); } catch (Exception e) { refused = e; } });
            thread.SetApartmentState(System.Threading.ApartmentState.MTA); thread.Start(); thread.Join();
            Assert.IsInstanceOfType(refused, typeof(InvalidOperationException));
        }

        [TestMethod]
        public void ConstructorNullsAndInvalidPlansRefuseWithoutDispatch()
        {
            using (var dispatcher = new Control())
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                for (int missing = 0; missing < 8; missing++)
                {
                    int argument = missing;
                    Assert.ThrowsException<ArgumentNullException>(() => new VbaNativeTestExecutionHost(argument == 0 ? null : new object(), argument == 1 ? null : dispatcher,
                        argument == 2 ? null : sink, argument == 3 ? (Func<VbaTestCatalog, object>)null : _ => new object(),
                        argument == 4 ? (Action<VbaTestCatalog>)null : _ => { }, argument == 5 ? (Action)null : () => { },
                        argument == 6 ? (Func<VbaTestCatalog, string>)null : _ => "signature", argument == 7 ? (Func<string>)null : () => "run"));
                }
            }
            using (var f = new Fixture())
            {
                Assert.ThrowsException<ArgumentException>(() => f.Host.InvokeAsync(null, f.Test, "Test"));
                Assert.ThrowsException<ArgumentException>(() => f.Host.InvokeAsync(new VbaTestCatalog(), f.Test, "Test"));
                Assert.ThrowsException<ArgumentException>(() => f.Host.InvokeAsync(f.Catalog, null, "Test"));
                f.Host.Post = _ => { throw new InvalidOperationException("Dispatch queue disconnected"); };
                var task = f.Host.InvokeAsync(f.Catalog, f.Test, "Test");
                Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                f.Host.Dispose();
                Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Validate(f.Catalog));
            }
        }

        [TestMethod]
        public void DuplicateQueuedCallbacksAndUnavailableModeCannotDispatchTwice()
        {
            using (var f = new Fixture())
            {
                f.Native.Mode = 0;
                var task = f.Host.InvokeAsync(f.Catalog, f.Test, "Test"); var queued = f.Queued;
                queued(); var tick = f.Tick;
                f.Native.Mode = 2; tick();
                Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
                queued(); tick(); Assert.AreEqual(1, f.Native.Invocations);
            }
            using (var f = new Fixture())
            {
                f.Native.Mode = -1;
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult()).Uncertain);
            }
            using (var f = new Fixture())
            {
                Exception error = null;
                var thread = new System.Threading.Thread(() => { try { f.Host.Validate(f.Catalog); } catch (Exception caught) { error = caught; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
            }
        }

        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        private static extern int CreateStreamOnHGlobal(IntPtr memory, bool deleteOnRelease, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Interface)] out object stream);

        [TestMethod]
        public void ReadOnlyDiagnosticsBoundTextAndHandleNullComAndFailedIdentityReads()
        {
            var read = typeof(VbaNativeTestExecutionHost).GetMethod("DiagnosticRead", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var matches = typeof(VbaNativeTestExecutionHost).GetMethod("DiagnosticMatches", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            Assert.AreEqual("null", read.Invoke(null, new object[] { new Func<object>(() => null) }));
            Assert.AreEqual(new string('x', 256), read.Invoke(null, new object[] { new Func<object>(() => new string('x', 257)) }));
            Assert.AreEqual("unavailable(InvalidOperationException)", read.Invoke(null, new object[] { new Func<object>(() => { throw new InvalidOperationException(); }) }));
            object stream; Assert.AreEqual(0, CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
            try { Assert.AreEqual("available", read.Invoke(null, new object[] { new Func<object>(() => stream) })); }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(stream); }
            Assert.AreEqual(false, matches.Invoke(null, new object[] { new Func<object, object, bool>((a, b) => { throw new InvalidOperationException(); }), new object(), new object() }));
            Assert.AreEqual(false, matches.Invoke(null, new object[] { new Func<object, object, bool>(ReferenceEquals), null, new object() }));
            var editor = new DiagnosticEditor { Windows = new DiagnosticWindow[13] };
            string diagnostics = VbaNativeTestExecutionHost.DescribeNativeWindows(editor, null, null, nativeWindows: _ => null);
            StringAssert.Contains(diagnostics, "bounded"); StringAssert.Contains(diagnostics, "ownedNative=[null]");
            VbaNativeTestExecutionHost.DescribeNativeWindows(new object(), null, null);
            VbaNativeTestExecutionHost.DescribeNativeSelection(editor, null, null, null, 1);
            var selectionEditor = new SelectionEditor();
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestExecutionHost.PrepareNativePane(selectionEditor, new SelectionModule(selectionEditor), 1));
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestExecutionHost.ValidateNativePaneContext(new SelectionEditor(), new object(), new object(), new object(), 1));
            Assert.ThrowsException<Microsoft.CSharp.RuntimeBinder.RuntimeBinderException>(() => VbaNativeTestExecutionHost.EnsureNativePaneFocus(new object(), new object(), new object(), new object(), 1, null));
        }

        [TestMethod]
        public void NativeDiagnosticObservationBoundsCountAndLengthWithoutReadingForeignWindowTitles()
        {
            var f = new VbaNativeTestWindowFocusTests.Windows();
            var root = f.Items[new IntPtr(1)];
            Assert.AreEqual("unavailable(owner mismatch)", VbaNativeTestExecutionHost.NativeWindowObservation.Read(IntPtr.Zero, f, (_, __) => Assert.Fail()));
            root.Exists = false;
            Assert.AreEqual("unavailable(owner mismatch)", VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (_, __) => Assert.Fail()));
            root.Exists = true; root.Process = 9;
            Assert.AreEqual("unavailable(owner mismatch)", VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (_, __) => Assert.Fail()));
            root.Process = 7; root.Thread = 9;
            Assert.AreEqual("unavailable(owner mismatch)", VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (_, __) => Assert.Fail()));
            root.Thread = 8;
            f.Add(5, 1, "Private"); f.Items[new IntPtr(5)].Process = 9;
            string foreign = VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (parent, visit) => visit(new IntPtr(5)));
            Assert.IsFalse(foreign.Contains("Private")); Assert.AreEqual(0, f.CaptionReads.Count);
            root.Class = null;
            f.Items[new IntPtr(2)].Class = null;
            string bounded = VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (parent, visit) => { for (int i = 0; i < 80 && visit(new IntPtr(2)); i++) { } });
            StringAssert.Contains(bounded, "bounded");
            f.Items[new IntPtr(2)].Class = new string('x', 200);
            int visits = 0;
            string longText = VbaNativeTestExecutionHost.NativeWindowObservation.Read(new IntPtr(1), f, (parent, visit) => { while (visit(new IntPtr(2))) visits++; });
            Assert.IsTrue(visits < 64); Assert.IsTrue(longText.Length >= 3900); Assert.IsFalse(longText.Contains(new string('x', 96)));
            RunOnSta(() => {
                using (var form = new Form())
                using (var child = new Control())
                {
                    form.Controls.Add(child); var rootHandle = form.Handle; var childHandle = child.Handle;
                    string native = VbaNativeTestExecutionHost.NativeWindowObservation.Read(rootHandle);
                    StringAssert.Contains(native, "HWnd=" + rootHandle.ToInt64());
                    StringAssert.Contains(native, "HWnd=" + childHandle.ToInt64());
                }
            });
        }

        [TestMethod]
        public void NativeProbeHostRechecksIdentityAfterPanePreparationBeforeArming()
        {
            RunOnSta(() =>
            {
                foreach (bool changed in new[] { false, true })
                {
                    var f = new NativeFixture();
                    var catalog = new VbaTestCatalog { Project = new VbaTestProjectSnapshot { Id = "owned", Revision = "revision", Modules = new[] { new VbaTestModuleSnapshot { Name = VbaTestRuntimeSource.ModuleName, Source = f.Module.Lines.Source } } } };
                    var test = new VbaTestDescriptor { Id = "test", Module = "Tests", Procedure = "Alpha", Kind = "Sub" };
                    int reads = 0, executions = 0;
                    using (var sink = new VbaTestResultSink(ReferenceEquals))
                    using (var dispatcher = new Control())
                    using (var host = new VbaNativeTestExecutionHost(f.Editor, dispatcher, sink, _ => ++reads == 2 && changed ? new object() : f.Project, _ => { }, () => { }, _ => "signature", () => "run") { Probe = f.Probe, Post = action => action() })
                    {
                        f.Editor.CommandBars.Control.OnExecute = () => {
                            executions++;
                            var runtime = new VbaTestRuntime();
                            var job = (object[])runtime.Request(VbaTestRuntimeSource.Version, "signature");
                            runtime.Publish((string)job[0], (string)job[4], (string)job[5], (string)job[6], "Passed", "", 0);
                        };
                        var task = host.InvokeAsync(catalog, test, "Test");
                        if (changed) Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                        else Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
                        Assert.AreEqual(changed ? 0 : 1, executions);
                    }
                }
            });
        }

        [TestMethod]
        public void ChangedFocusTargetContextCannotInvokeNativeRecovery()
        {
            foreach (string fault in new[] { "project", "module", "type", "visible" })
            {
                var f = new NativeFixture(); f.Module.CodePane.Window.Visible = true;
                if (fault == "project") f.Editor.ActiveVBProject = new object();
                if (fault == "module") f.Module.CodePane.CodeModule = new object();
                if (fault == "type") f.Module.CodePane.Window.Type = 1;
                if (fault == "visible") f.Module.CodePane.Window.Visible = false;
                Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestExecutionHost.EnsureNativePaneFocus(f.Editor, f.Project, f.Module, f.Module.CodePane, 1, f.Windows, ReferenceEquals));
                Assert.AreEqual(0, f.Windows.Activations); Assert.AreEqual(0, f.Windows.Focuses);
            }
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestExecutionHost.ValidateNativeRunControl(null));
            var project = new DiagnosticProject();
            var module = new DiagnosticModule { Parent = new DiagnosticComponent { Collection = new DiagnosticCollection() } };
            var pane = new DiagnosticPane { CodeModule = module };
            var editor = new DiagnosticEditor { ActiveCodePane = pane };
            VbaNativeTestExecutionHost.DescribeNativeSelection(editor, project, module, pane, 1, ReferenceEquals);
            pane.CodeModule = null;
            VbaNativeTestExecutionHost.DescribeNativeSelection(editor, project, module, pane, 1, ReferenceEquals);
        }

    }
}
