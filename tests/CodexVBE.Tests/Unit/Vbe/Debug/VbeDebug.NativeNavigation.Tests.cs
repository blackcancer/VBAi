using System.Dynamic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeNavigationWorkflowTests
    {
        private SynchronizationContext saved;
        private NativeNavigationContext context;
        [TestInitialize] public void IsolateDispatcher() { saved = SynchronizationContext.Current; context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context); }
        [TestCleanup] public void RestoreDispatcher() { SynchronizationContext.SetSynchronizationContext(saved); }
        private static Request Last(EditorDebugFixture f)
        {
            var r = f.Location("last_position"); r.StartLine = f.Pane.Start; r.StartColumn = f.Pane.Column; r.EndColumn = f.Pane.EndColumn; return r;
        }
        private static dynamic Status(EditorDebugFixture f, string id) { return f.Service.NativeNavigation(new Request { Action = "status", Query = id }); }
        [TestMethod]
        public void NavigationRequiresKnownActionContextModeAndExactEnabledCommand()
        {
            var f = new EditorDebugFixture();
            Assert.ThrowsException<InvalidOperationException>(() => Status(f, "unknown"));
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeNavigation(f.Location("wrong")));
            SynchronizationContext.SetSynchronizationContext(null);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(Last(f)));
            SynchronizationContext.SetSynchronizationContext(context);
            var r = Last(f);
            foreach (int mode in new[] { 0, 3 }) { f.Project.Mode = mode; r.ExpectedMode = mode; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r)); }
            f.Project.Mode = 1; r.ExpectedMode = 2; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            r.ExpectedMode = 1; r.ControlCaption = null; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            r.ControlCaption = "last_position"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            var cmd = f.Command(10, "last_position"); Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            cmd.Id = 1822; cmd.Enabled = false; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            cmd.Enabled = true; cmd.Caption = "wrong"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            Assert.AreEqual(0, context.Pending);
        }
        [TestMethod]
        public void DefinitionValidatesSourceSelectionAndCapturesRequestBeforeScheduling()
        {
            var f = new EditorDebugFixture(); f.Command(939, "definition", () => f.Pane.EndColumn = 7);
            var r = f.Location("definition");
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeNavigation(r));
            r.Expression = "Debug"; r.StartColumn = 0; r.EndColumn = 6;
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeNavigation(r));
            r.StartColumn = 1; r.EndColumn = 1; Assert.ThrowsException<ArgumentException>(() => f.Service.NativeNavigation(r));
            r.EndColumn = 6;
            dynamic queued = f.Service.NativeNavigation(r); string id = queued.OperationId;
            Assert.AreEqual("Queued", (string)queued.State); Assert.AreEqual(6, f.Pane.EndColumn);
            r.Project = "changed by caller"; r.Module = "changed by caller"; r.ExpectedSha256 = "changed"; r.ControlCaption = "changed"; r.ExpectedMode = 0; r.StartLine = 99;
            context.RunAll(); dynamic result = Status(f,id);
            Assert.AreEqual("Completed", (string)result.State); Assert.IsTrue((bool)result.NavigationObserved);
            Assert.IsTrue((bool)result.CommandCompleted); Assert.IsFalse((bool)result.DefinitionResolved);
        }
        [TestMethod]
        public void LastPositionRequiresTheExactActiveModuleAndSelection()
        {
            var f = new EditorDebugFixture(); f.Command(1822, "last_position"); var r = Last(f);
            f.Vbe.ActiveCodePane = null; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            f.Vbe.ActiveCodePane = f.AddModule("other").CodePane; Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            f.Vbe.ActiveCodePane = f.Pane;
            foreach (int mismatch in new[] { 0, 1, 2, 3 })
            {
                f.Pane.Start = 1; f.Pane.End = 1; f.Pane.Column = 1; f.Pane.EndColumn = 1;
                if (mismatch == 0) f.Pane.Start = 2; if (mismatch == 1) f.Pane.End = 2;
                if (mismatch == 2) f.Pane.Column = 2; if (mismatch == 3) f.Pane.EndColumn = 2;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(r));
            }
        }
        [TestMethod]
        public void NavigationRejectsNewRequestsWhileQueuedRunningOrObserving()
        {
            var f = new EditorDebugFixture(); string id = null;
            f.Command(1822, "last_position", () => {
                Assert.AreEqual("Running", (string)Status(f,id).State);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(Last(f)));
            });
            dynamic queued = f.Service.NativeNavigation(Last(f)); id = queued.OperationId;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(Last(f)));
            context.RunNext(); Assert.AreEqual("Observing", (string)Status(f,id).State);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeNavigation(Last(f)));
            context.RunAll(); Thread.Sleep(2100);
            dynamic result = Status(f,id); Assert.AreEqual("Completed", (string)result.State);
            Assert.IsFalse((bool)result.NavigationObserved); Assert.IsTrue((bool)result.CommandCompleted);
        }
        [TestMethod]
        public void QueuedNavigationRejectsModeRevisionPositionAndCommandChanges()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4, 5, 6, 7 })
            {
                var f = new EditorDebugFixture(); var cmd = f.Command(1822, "last_position");
                dynamic queued = f.Service.NativeNavigation(Last(f)); string id = queued.OperationId;
                if (scenario == 0) f.Project.Mode = 1;
                if (scenario == 1) f.Module.Code = "changed";
                if (scenario == 2) f.Pane.EndColumn = 2;
                if (scenario == 3) f.Vbe.CommandBars[0].Controls.Clear();
                if (scenario == 4) cmd.Enabled = false;
                if (scenario == 5) cmd.Caption = "changed";
                if (scenario == 6) f.Pane.Window.OnFocus = () => f.Pane.TopLine = 2; // viewport alone is not a navigation-position change
                if (scenario == 7) f.Pane.Window.OnFocus = () => f.Pane.EndColumn = 2;
                context.RunAll(); dynamic result = Status(f,id);
                if (scenario == 6) { Assert.AreEqual("Observing", (string)result.State); Assert.AreEqual(1, cmd.Executions); }
                else { Assert.AreEqual("Failed", (string)result.State); Assert.IsNotNull(result.NativeError); Assert.AreEqual(0, cmd.Executions); }
            }
        }
        [TestMethod]
        public void NavigationPreservesDispatchExecutionAndObservationFailures()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3 })
            {
                context = new NativeNavigationContext { RejectPost = scenario == 0 ? 1 : scenario == 1 ? 2 : 0 };
                SynchronizationContext.SetSynchronizationContext(context);
                var f = new EditorDebugFixture(); f.Command(1822, "last_position", () => {
                    if (scenario == 2) throw new InvalidOperationException("execute failed");
                    if (scenario == 3) f.Pane.OnReadSelection = () => { throw new InvalidOperationException("observation failed"); };
                });
                dynamic queued = f.Service.NativeNavigation(Last(f)); string id = queued.OperationId;
                context.RunAll(); dynamic result = Status(f,id);
                Assert.AreEqual("Failed", (string)result.State); Assert.IsNotNull(result.NativeError);
                Assert.AreEqual(scenario == 1 || scenario == 3, (bool)result.CommandCompleted);
            }
        }
        [TestMethod]
        public void NavigationSnapshotsUnsavedProjectsBrowserVisibilityAndClosedPanes()
        {
            foreach (int scenario in new[] { 0, 1, 2 })
            {
                var f = new EditorDebugFixture(); f.Project.FileName = " "; f.Project.FailPath = scenario == 1;
                f.Vbe.ActiveWindow = null;
                f.Vbe.Windows.Add(new EditorDebugFixture.Window { Type = 2, Visible = false });
                f.Command(1822, "last_position", () => {
                    if (scenario == 2) f.Vbe.ActiveCodePane = null;
                    f.Vbe.Windows.Add(new EditorDebugFixture.Window { Type = 2 });
                });
                dynamic queued = f.Service.NativeNavigation(Last(f)); string id = queued.OperationId;
                Assert.AreEqual(f.Project.Name, (string)queued.Before.Project); Assert.IsNull(queued.Before.ActiveWindow);
                context.RunAll(); dynamic result = Status(f,id);
                Assert.AreEqual("Completed", (string)result.State); Assert.IsTrue((bool)result.NavigationObserved);
                if (scenario == 2) Assert.IsFalse((bool)result.After.CodePaneAvailable);
            }
        }
        [TestMethod]
        public void NavigationHistoryRetainsOnlyTwentyCompletedOperations()
        {
            var f = new EditorDebugFixture(); f.Command(1822, "last_position", () => f.Pane.EndColumn = f.Pane.EndColumn == 1 ? 2 : 1);
            string first = null;
            for (int i=0; i<21; i++)
            {
                dynamic queued = f.Service.NativeNavigation(Last(f)); string id = queued.OperationId; if (i==0) first=id;
                context.RunAll(); Assert.AreEqual("Completed", (string)Status(f,id).State);
            }
            Assert.ThrowsException<InvalidOperationException>(() => Status(f,first));
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativeNavigationTests
    {
        private static dynamic Position()
        {
            dynamic window = new ExpandoObject(); window.Type = 0; window.Caption = "Module (Code)";
            dynamic position = new ExpandoObject(); position.ActiveWindow = window; position.ObjectBrowserVisible = false;
            position.CodePaneAvailable = true; position.Project = "project"; position.Module = "module";
            position.StartLine = 2; position.EndLine = 2; position.StartColumn = 5; position.EndColumn = 8;
            return position;
        }
        [TestMethod]
        public void CaptionAndMainFrameChangesAloneAreNotNavigation()
        {
            dynamic before = Position(), after = Position();
            after.ActiveWindow.Caption = "Workbook - Module (Code)";
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
            after.ActiveWindow.Type = 12;
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
        }
        [TestMethod]
        public void EverySourceCoordinateAndAvailabilityBoundaryHasDistinctSemantics()
        {
            foreach (string field in new[] { "Project", "Module", "StartLine", "StartColumn", "EndLine", "EndColumn" })
            {
                dynamic before = Position(), after = Position();
                var values = (System.Collections.Generic.IDictionary<string, object>)after;
                values[field] = field == "Project" || field == "Module" ? (object)"different" : 99;
                Assert.IsTrue(VbeDebug.NavigationChanged(before,after), field);
            }
            dynamic a = Position(), b = Position();
            a.ActiveWindow = null; b.ActiveWindow.Type = 2; Assert.IsTrue(VbeDebug.NavigationChanged(a,b));
            a = Position(); b = Position(); b.ActiveWindow.Type = 2; Assert.IsTrue(VbeDebug.NavigationChanged(a,b));
            a.ActiveWindow.Type = 2; Assert.IsFalse(VbeDebug.NavigationChanged(a,b));
            b.ActiveWindow = null; Assert.IsFalse(VbeDebug.NavigationChanged(a,b));
            a.CodePaneAvailable = false; Assert.IsFalse(VbeDebug.NavigationChanged(a,b));
            a.CodePaneAvailable = true; b.CodePaneAvailable = false; Assert.IsFalse(VbeDebug.NavigationChanged(a,b));
        }
        [TestMethod]
        public void BrowserAppearanceIsObservedWithoutClaimingItsSelectedMember()
        {
            dynamic before = Position(), after = Position();
            after.ObjectBrowserVisible = true; after.ActiveWindow.Type = 12;
            Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
            before.ObjectBrowserVisible = true;
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
        }
        [TestMethod]
        public void SourcePositionAndModuleChangesAreObserved()
        {
            dynamic before = Position(), after = Position();
            after.Module = "target"; Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
            after.Module = before.Module; after.StartLine = 4;
            Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
        }
    }
}
