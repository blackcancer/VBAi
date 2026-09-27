using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeEditorWindowsTests
    {
        [TestMethod]
        public void WindowCommandsUseExactCaptionAndReadBackTheResult()
        {
            var host = new FakeVbe();
            var hidden = new FakeWindow(host) { Caption = "Immediate", Type = 3, Visible = false };
            host.Windows.Add(hidden);
            var editor = new VbeEditorWindows(host);
            Assert.ThrowsException<ArgumentException>(() => editor.FocusWindow("", 3));
            Assert.ThrowsException<InvalidOperationException>(() => editor.FocusWindow("Immediate", 3));
            Assert.ThrowsException<InvalidOperationException>(() => editor.ShowWindow("immediate", 3));
            dynamic shown = editor.ShowWindow("Immediate", 3);
            Assert.IsFalse((bool)shown.WasVisible);
            Assert.IsTrue((bool)shown.Visible);
            Assert.IsTrue((bool)shown.FocusVerified);
            dynamic focused = editor.FocusWindow("Immediate", 3);
            Assert.AreEqual("ActiveWindowReadback", (string)focused.Verification);
            Assert.AreEqual(2, hidden.FocusCount);
            dynamic closed = editor.CloseWindow("Immediate", 3);
            Assert.AreEqual("HiddenInWindows", (string)closed.Verification);
            Assert.IsFalse(hidden.Visible);
            Assert.ThrowsException<InvalidOperationException>(() => editor.CloseWindow("Immediate", 3));
            Assert.ThrowsException<InvalidOperationException>(() => editor.CloseWindow("CodexVBE", 3));
            host.Windows.Add(new FakeWindow(host) { Caption = "Immediate", Type = 3, Visible = true });
            Assert.ThrowsException<InvalidOperationException>(() => editor.ShowWindow("Immediate", 3));
        }

        [TestMethod]
        public void SnapshotsAndEnvironmentReportCollectionsAndLinkedWindows()
        {
            var host = new FakeVbe();
            var window = new FakeWindow(host) { Caption = "Module1", Type = 0, Visible = true,
                LinkedWindowFrame = new FakeFrame { Caption = "Frame", LinkedWindows = new List<FakeWindow>() } };
            window.LinkedWindowFrame.LinkedWindows.Add(window);
            host.Windows.Add(window);
            host.ActiveWindow = window;
            host.VBProjects.Add(new FakeProject { Name = "Projet" });
            host.ActiveVBProject = host.VBProjects[0];
            host.AddIns.Add(new FakeAddIn { ProgId = "CodexVBE.AddIn", Guid = "test", Description = "Test", Connect = true });
            var editor = new VbeEditorWindows(host);
            dynamic snapshot = editor.Windows();
            Assert.AreEqual(1, (int)snapshot.Windows.Count);
            Assert.AreEqual("Module1", (string)snapshot.Windows[0].Properties["Caption"]);
            Assert.AreEqual("Module1", (string)snapshot.ActiveWindow.Properties["Caption"]);
            dynamic environment = editor.Environment();
            Assert.AreEqual("7.1", (string)environment.Properties["Version"]);
            Assert.AreEqual(1, (int)environment.Properties["ProjectCount"]);
            Assert.AreEqual(1, (int)environment.Properties["AddInCount"]);
            dynamic addIns = editor.AddIns();
            Assert.AreEqual(1, (int)addIns.Count);
            Assert.AreEqual("CodexVBE.AddIn", (string)addIns.AddIns[0].Properties["ProgId"]);
            dynamic linkage = editor.WindowLinkage("Module1", 0);
            Assert.AreEqual(true, (bool)linkage.Properties["IsLinked"]);
            Assert.AreEqual("Frame", (string)linkage.Properties["FrameCaption"]);
            Assert.AreEqual(1, (int)linkage.Properties["LinkedWindows"].Count);
        }

        [TestMethod]
        public void CodePaneSnapshotReadsProjectModuleWindowAndSelectionWithoutOpeningPane()
        {
            var host = new FakeVbe();
            var pane = new FakePane { CodeModule = new FakePaneModule {
                Parent = new FakePaneComponent { Name = "Module1", Collection = new FakePaneCollection {
                    Parent = new FakePaneProject { Name = "Projet", FileName = @"C:\Temp\Projet.xlsm" } } } },
                CodePaneView = 1, TopLine = 9, CountOfVisibleLines = 30,
                Window = new FakePaneWindow { Caption = "Module1 (Code)" } };
            host.CodePanes.Add(pane);
            host.ActiveCodePane = pane;
            dynamic snapshot = new VbeEditorWindows(host).CodePanes();
            Assert.AreEqual(1, (int)snapshot.CodePanes.Count);
            dynamic first = snapshot.CodePanes[0];
            Assert.AreEqual(1, (int)first.Index);
            Assert.AreEqual("Projet", (string)first.Properties["Project"]);
            Assert.AreEqual("Module1", (string)first.Properties["Module"]);
            Assert.AreEqual("Module1 (Code)", (string)first.Properties["WindowCaption"]);
            Assert.AreEqual(9, (int)first.Properties["TopLine"]);
            Assert.AreEqual(2, (int)((dynamic)first.Properties["Selection"]).StartLine);
            Assert.AreEqual(2, pane.SelectionReads);
            Assert.AreEqual(0, (int)first.Errors.Count);
        }

        [TestMethod]
        public void CodePaneSnapshotReportsUnavailableMembersIndividually()
        {
            var host = new FakeVbe();
            host.CodePanes.Add(new FakePane { SelectionThrows = true });
            dynamic snapshot = new VbeEditorWindows(host).CodePanes();
            dynamic first = snapshot.CodePanes[0];
            Assert.IsTrue(first.Errors.ContainsKey("Module"));
            Assert.IsTrue(first.Errors.ContainsKey("WindowCaption"));
            Assert.IsTrue(first.Errors.ContainsKey("Selection"));
            Assert.AreEqual(1, (int)first.Properties["CodePaneView"]);
            Assert.IsNull(snapshot.ActiveCodePane);
        }

        public sealed class FakeVbe
        {
            public string Version { get; set; } = "7.1";
            public List<FakeWindow> Windows { get; } = new List<FakeWindow>();
            public List<object> CodePanes { get; } = new List<object>();
            public List<FakeAddIn> AddIns { get; } = new List<FakeAddIn>();
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
            public FakeProject ActiveVBProject { get; set; }
            public FakeWindow ActiveWindow { get; set; }
            public object ActiveCodePane { get; set; }
        }
        public sealed class FakeProject { public string Name { get; set; } }
        public sealed class FakePaneProject { public string Name { get; set; } public string FileName { get; set; } }
        public sealed class FakePaneCollection { public FakePaneProject Parent { get; set; } }
        public sealed class FakePaneComponent { public string Name { get; set; } public FakePaneCollection Collection { get; set; } }
        public sealed class FakePaneModule { public FakePaneComponent Parent { get; set; } }
        public sealed class FakePaneWindow { public string Caption { get; set; } }
        public sealed class FakePane
        {
            public FakePaneModule CodeModule { get; set; }
            public int CodePaneView { get; set; } = 1;
            public int TopLine { get; set; }
            public int CountOfVisibleLines { get; set; }
            public FakePaneWindow Window { get; set; }
            public bool SelectionThrows { get; set; }
            public int SelectionReads { get; private set; }
            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                SelectionReads++;
                if (SelectionThrows) throw new InvalidOperationException("Selection unavailable");
                startLine = 2; startColumn = 1; endLine = 4; endColumn = 8;
            }
        }
        public sealed class FakeAddIn
        {
            public string ProgId { get; set; }
            public string Guid { get; set; }
            public string Description { get; set; }
            public bool Connect { get; set; }
        }
        public sealed class FakeFrame
        {
            public string Caption { get; set; }
            public List<FakeWindow> LinkedWindows { get; set; }
        }
        public sealed class FakeWindow
        {
            private readonly FakeVbe host;
            public FakeWindow(FakeVbe host) { this.host = host; }
            public string Caption { get; set; }
            public int Type { get; set; }
            public bool Visible { get; set; }
            public int WindowState { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; } = 500;
            public int Height { get; set; } = 300;
            public FakeFrame LinkedWindowFrame { get; set; }
            public int FocusCount { get; private set; }
            public void SetFocus() { FocusCount++; host.ActiveWindow = this; }
            public void Close() { Visible = false; }
        }
    }
}
