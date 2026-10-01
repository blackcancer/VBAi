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
        }
        public sealed class SelectionWindow
        {
            public readonly List<string> Events;
            private bool visible;
            public bool AllowVisibility = true;
            public int Type { get; set; }
            public long HWnd { get; set; }
            public string Caption { get; set; } = "Owned support (Code)";
            public SelectionWindow(List<string> events) { Events = events; }
            public bool Visible { get => visible; set { Events.Add("Visible"); visible = value && AllowVisibility; } }
            public void SetFocus() { Events.Add(Type == 12 ? "MainFocus" : "Focus"); Assert.IsTrue(Visible); }
        }
        public sealed class SelectionEditor
        {
            public readonly List<string> Events = new List<string>();
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
            internal Action OnPrepare, OnRevalidate, OnExecute;
            internal int Mode = 2, Invocations;
            public void RequireOwner(object vbe) { }
            public object Prepare(object vbe, object project, string source) { OnPrepare?.Invoke(); return new object(); }
            public void Revalidate(object vbe, object project, object prepared) { OnRevalidate?.Invoke(); }
            public void Execute(object prepared) { Invocations++; OnExecute?.Invoke(); }
            public int ReadMode(object project) => Mode;
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
    }
}
