using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeCodePaneWorkflowTests
    {
        [TestMethod]
        public void LayoutReportsExpiredEntriesEmptyModulesAndInvalidatesOldTokens()
        {
            var f = new EditorDebugFixture(); var first = f.Inspect();
            f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { ModuleError = unchecked((int)0x80020010) });
            f.AddModule("other", "");
            dynamic layout = f.Service.CodePaneLayout(f.Location());
            Assert.AreEqual(1, ((IEnumerable)layout.Panes).Cast<object>().Count());
            Assert.AreEqual(1, ((IEnumerable)layout.UnavailablePanes).Cast<object>().Count());
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(first));
            var empty = f.Location(); empty.Module = "other";
            dynamic emptyLayout = f.Service.CodePaneLayout(empty);
            dynamic row = ((IEnumerable)emptyLayout.Panes).Cast<object>().Single();
            Assert.AreEqual(0, (int)row.State.LineCount);
            f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window });
            f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CodePaneLayout(f.Location()));
        }
        [TestMethod]
        public void ScrollRejectsMissingStaleForeignAndClosedPanes()
        {
            var f = new EditorDebugFixture(); var request = f.Location();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            request.Pane = "expired"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            request = f.Inspect();
            foreach (int mode in new[] { 0, 3 }) { f.Project.Mode = mode; request.ExpectedMode = mode; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request)); }
            f.Project.Mode = 1; request.ExpectedMode = 2; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            request.ExpectedMode = 1; f.AddModule("other"); request.Module = "other";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            request.Module = "Module1"; f.Vbe.Panes.Remove(f.Pane); Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            f.Vbe.Panes.Add(f.Pane); request.ExpectedWindowVersion = null; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
            request.ExpectedWindowVersion = "stale"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ScrollCodePane(request));
        }
        [TestMethod]
        public void ScrollVerifiesTopLineSelectionHashAndBothErrorChannels()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4, 5 })
            {
                var f = new EditorDebugFixture(); var request = f.Inspect(); request.StartLine = 2;
                f.Pane.OnSetTop = value =>
                {
                    if (scenario == 1) throw new COMException("scroll rejected");
                    if (scenario == 2) f.Pane.OnReadSelection = () => { throw new COMException("readback rejected"); };
                    if (scenario == 3) f.Pane.OnReadSelection = () => { f.Pane.OnSetTop = null; f.Pane.TopLine = 1; };
                    if (scenario == 4) f.Pane.EndColumn = 2;
                    if (scenario == 5) f.Module.Code = "new code";
                };
                dynamic result = f.Service.ScrollCodePane(request);
                Assert.AreEqual(scenario == 0, (bool)result.Verified); Assert.AreEqual(scenario != 0, (bool)result.VerificationPending);
                if (scenario == 1) { Assert.IsNull(result.Applied); Assert.AreEqual("scroll rejected", (string)result.NativeError); }
                if (scenario == 2) { Assert.IsNull(result.After); Assert.IsNull(result.WindowVersion); Assert.AreEqual("readback rejected", (string)result.ReadbackError); }
                if (scenario == 0) Assert.AreEqual(2, (int)result.After.TopLine);
            }
        }
        [TestMethod]
        public void ViewRejectsInvalidActionTokenModeModuleAndTopology()
        {
            var f = new EditorDebugFixture(); var request = f.Location("wrong");
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetCodePaneView(request));
            request.Action = "module"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request.Pane = "expired"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request = f.Inspect("procedure");
            foreach (int mode in new[] { 0, 3 }) { f.Project.Mode = mode; request.ExpectedMode = mode; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request)); }
            f.Project.Mode = 1; request.ExpectedMode = 2; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request.ExpectedMode = 1; f.AddModule("other"); request.Module = "other";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request.Module = "Module1"; f.Vbe.Panes.Remove(f.Pane); Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            f.Vbe.Panes.Add(f.Pane); f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
        }
        [TestMethod]
        public void ViewRequiresTheInspectedVersionOfEverySplitPane()
        {
            var f = new EditorDebugFixture(); var request = f.Inspect("procedure"); request.ExpectedWindowVersion = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request.ExpectedWindowVersion = "stale"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request = f.Inspect("procedure"); var second = new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window };
            f.Vbe.Panes.Add(second); Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            request = f.Inspect("procedure"); second.TopLine = 2; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
        }
        [TestMethod]
        public void ViewDoesNothingWhenEveryPaneAlreadyHasTheRequestedView()
        {
            var f = new EditorDebugFixture(); f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { ModuleError = unchecked((int)0x80020010) }); f.AddModule("other");
            dynamic result = f.Service.SetCodePaneView(f.Inspect("module"));
            Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Changed);
            f.Pane.CodePaneView = 0; Assert.IsFalse((bool)((dynamic)f.Service.SetCodePaneView(f.Inspect("procedure"))).Changed);
        }
        [TestMethod]
        public void ViewRejectsActivePaneAndViewportChangesDuringFocus()
        {
            foreach (int scenario in new[] { 0, 1, 2 })
            {
                var f = new EditorDebugFixture(); var second = new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window };
                f.Vbe.Panes.Add(second); var request = f.Inspect("procedure");
                f.Pane.Window.OnFocus = () => { if (scenario == 0) f.Vbe.ActiveCodePane = second; if (scenario == 1) f.Pane.TopLine = 2; if (scenario == 2) second.TopLine = 2; };
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodePaneView(request));
            }
        }
        [TestMethod]
        public void ViewUsesNativeToolbarAndVerifiesEveryPaneAndSourceHash()
        {
            foreach (int scenario in new[] { 0, 1, 2 })
                using (var scene = new NativeDebugScene())
                {
                    scene.CodeWindow(); var f = new EditorDebugFixture();
                    f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module, Window = f.Pane.Window, CodePaneView = 0 });
                    var request = f.Inspect("procedure");
                    scene.OnClick = () => { if (scenario != 1) f.Pane.CodePaneView = 0; if (scenario == 2) f.Module.Code = "changed"; };
                    dynamic result = f.Service.SetCodePaneView(request);
                    Assert.IsTrue((bool)result.Changed); Assert.AreEqual(scenario == 0, (bool)result.Verified); Assert.AreEqual(scenario != 0, (bool)result.VerificationPending);
                    Assert.AreEqual(2, ((object[])result.AfterPanes).Length); Assert.IsNotNull(result.Native);
                }
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed class VbeCodePaneLifetimeTests
    {
        public sealed class Pane
        {
            public int Error;
            public object Module = new object();
            public object CodeModule { get { if (Error != 0) throw new COMException("native pane error", Error); return Module; } }
        }
        [TestMethod]
        public void LivePanePreservesExactModuleIdentity()
        {
            var pane = new Pane();
            Assert.IsTrue(VbeDebug.TryLivePaneModule(pane, out object module, out string error));
            Assert.AreSame(pane.Module, module);
            Assert.IsNull(error);
        }
        [TestMethod]
        public void DestroyedSplitPaneIsReportedWithoutAUsableModule()
        {
            Assert.IsFalse(VbeDebug.TryLivePaneModule(new Pane { Error = unchecked((int)0x80020010) }, out object module, out string error));
            Assert.IsNull(module);
            Assert.AreEqual("native pane error", error);
        }
        [TestMethod]
        public void OtherNativeFailuresAreNotMisclassifiedAsDestroyedPanes()
        {
            Assert.ThrowsException<COMException>(() => VbeDebug.TryLivePaneModule(new Pane { Error = unchecked((int)0x80004005) }, out object module, out string error));
        }
    }
}
