using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeArrangementWorkflowTests
    {
        private static Request Inspected(EditorDebugFixture f, string action = "tile_vertical")
        {
            var request = f.Location(action); dynamic layout = f.Service.EditorLayout();
            request.ExpectedWindowVersion = layout.WindowVersion; return request;
        }
        [TestMethod]
        public void LayoutReadsExactCodeDesignerAndBrowserIdentitiesAndFiltersOtherWindows()
        {
            var f = new EditorDebugFixture(); var other = f.AddModule("other");
            var designer = new EditorDebugFixture.Window { Type = 1, Caption = "form" };
            var form = new EditorDebugFixture.Component { Name = "Form", HasOpenDesigner = true, Designer = designer };
            f.Project.VBComponents.Add(form); f.Vbe.Windows.Add(designer);
            f.Vbe.Windows.Add(new EditorDebugFixture.Window { Type = 2, Caption = "Browser" });
            foreach (var window in new[] {
                new EditorDebugFixture.Window { Type = -1 }, new EditorDebugFixture.Window { Type = 3 },
                new EditorDebugFixture.Window { Visible = false }, new EditorDebugFixture.Window { LinkedWindowFrame = new object() } }) f.Vbe.Windows.Add(window);
            dynamic layout = f.Service.EditorLayout();
            Assert.AreEqual(4, ((EditorWindowBounds[])layout.Windows).Length);
            foreach (bool unavailable in new[] { false, true })
            {
                f.Project.FileName = " "; f.Project.FailPath = unavailable;
                Assert.AreEqual(4, ((EditorWindowBounds[])((dynamic)f.Service.EditorLayout()).Windows).Length);
            }
            f.Project.FailPath = false; form.HasOpenDesigner = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.EditorLayout());
            form.HasOpenDesigner = true;
            var sameCaption = f.AddModule("duplicate"); sameCaption.CodePane.Window.Caption = f.Pane.Window.Caption;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.EditorLayout());
        }
        [TestMethod]
        public void LayoutRejectsUnknownAndDuplicateNativeDocumentIdentities()
        {
            var f = new EditorDebugFixture(); f.Pane.Window.Caption = "orphan";
            // No matching open native pane: the document must not acquire a guessed identity.
            f.Vbe.Panes.Clear(); Assert.ThrowsException<InvalidOperationException>(() => f.Service.EditorLayout());
            f = new EditorDebugFixture(); f.Vbe.Windows.Add(f.Pane.Window);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.EditorLayout());
            f = new EditorDebugFixture();
            var designer = new EditorDebugFixture.Window { Type = 1, Caption = "form" };
            f.Vbe.Windows.Add(designer);
            f.Project.VBComponents.Add(new EditorDebugFixture.Component { Name = "formA", HasOpenDesigner = true, Designer = designer });
            f.Project.VBComponents.Add(new EditorDebugFixture.Component { Name = "formB", HasOpenDesigner = true, Designer = designer });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.EditorLayout());
        }
        [TestMethod]
        public void ArrangementRequiresSupportedActionTwoTo64WindowsVersionAndExactCommand()
        {
            var f = new EditorDebugFixture(); var request = f.Location("wrong");
            Assert.ThrowsException<ArgumentException>(() => f.Service.ArrangeEditorWindows(request));
            request.Action = "cascade"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            f.Vbe.Windows.Clear(); Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            f = new EditorDebugFixture(); for (int i = 1; i < 65; i++) f.AddModule("m" + i);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(f.Location("cascade")));
            f = new EditorDebugFixture(); f.AddModule("other"); request = f.Location("cascade");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            request.ExpectedWindowVersion = "stale"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            request = Inspected(f, "cascade"); Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            var command = f.Command(10, "cascade"); Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            command.Id = 1826; command.Enabled = false; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            command.Enabled = true; command.Caption = "wrong"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.ArrangeEditorWindows(request));
            Assert.AreEqual(0, command.Executions);
        }
        [TestMethod]
        public void ArrangementVerifiesAllThreeActionsAgainstReadbackGeometry()
        {
            foreach (string action in new[] { "cascade", "tile_vertical", "tile_horizontal" })
            {
                var f = new EditorDebugFixture(); var other = f.AddModule("other");
                int id = action == "cascade" ? 1826 : action == "tile_vertical" ? 2561 : 2562;
                f.Command(id, action, () =>
                {
                    other.CodePane.Window.Left = action == "tile_vertical" ? 300 : action == "cascade" ? 26 : 0;
                    other.CodePane.Window.Top = action == "tile_horizontal" ? 200 : action == "cascade" ? 26 : 0;
                });
                var request = Inspected(f, action); request.ExpectedWindowVersion = request.ExpectedWindowVersion.ToUpperInvariant();
                dynamic result = f.Service.ArrangeEditorWindows(request);
                Assert.IsTrue((bool)result.Verified); Assert.IsTrue((bool)result.Applied); Assert.IsFalse((bool)result.VerificationPending);
                Assert.AreEqual(id, (int)result.CommandId); Assert.IsNotNull(result.WindowVersion);
            }
        }
        [TestMethod]
        public void ArrangementReportsExecutionReadbackTopologyAndGeometryFailures()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4 })
            {
                var f = new EditorDebugFixture(); var other = f.AddModule("other");
                f.Command(2561, "tile_vertical", () =>
                {
                    if (scenario == 1 || scenario == 2) f.Vbe.Panes.Remove(other.CodePane);
                    if (scenario == 3) f.Vbe.Windows.Remove(other.CodePane.Window);
                    if (scenario == 0 || scenario == 2) throw new InvalidOperationException("execute failed");
                });
                dynamic result = f.Service.ArrangeEditorWindows(Inspected(f));
                Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.VerificationPending);
                if (scenario < 3) { Assert.IsNull(result.Applied); StringAssert.Contains((string)result.NativeError, scenario == 1 ? "cannot be mapped" : "execute failed"); }
                else { Assert.IsTrue((bool)result.Applied); Assert.IsNull(result.NativeError); }
                if (scenario == 1 || scenario == 2) { Assert.IsNull(result.After); Assert.IsNull(result.WindowVersion); }
            }
        }
        [TestMethod]
        public void GeometryRejectsInvalidHeightUnequalHeightAndBoundaryCascadeSteps()
        {
            var a = new EditorWindowBounds { Width = 300, Height = 200 };
            var b = new EditorWindowBounds { Left = 26, Top = 26, Width = 300, Height = 0 };
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { a, b }));
            b.Height = 203; Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { a, b }));
            b.Height = 200; b.Top = 200; Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { a, b }));
            b.Top = 26; b.Left = 300; Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { a, b }));
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed class VbeArrangementTests
    {
        private static EditorWindowBounds Rect(int left, int top, int width = 300, int height = 200)
        {
            return new EditorWindowBounds { Left = left, Top = top, Width = width, Height = height };
        }

        [TestMethod]
        public void TilingRequiresTheRequestedOrientation()
        {
            var columns = new[] { Rect(0, 0), Rect(300, 0), Rect(600, 0) };
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_vertical", columns));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_horizontal", columns));
            var lines = new[] { Rect(0, 0), Rect(0, 200), Rect(0, 400) };
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_horizontal", lines));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", lines));
        }

        [TestMethod]
        public void TilingRejectsOverlapsGapsAndUnequalSizes()
        {
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(280, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(500, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(300, 0, 150) }));
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(301, 0, 301) }));
        }

        [TestMethod]
        public void CascadeRequiresOverlappingDiagonalStepsRegardlessOfIdentityOrder()
        {
            Assert.IsTrue(VbeDebug.VerifyArrangement("cascade", new[] { Rect(52, 52), Rect(0, 0), Rect(26, 26) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(26, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(300, 200) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(0, 0) }));
        }

        [TestMethod]
        public void InvalidOrMaximizedWindowsCannotVerifyAnArrangement()
        {
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", null));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0, 0), Rect(300, 0) }));
            var maximized = Rect(0, 0); maximized.State = 2;
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { maximized, Rect(300, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("unknown", new[] { Rect(0, 0), Rect(300, 0) }));
        }
    }
}
