namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeEditorWindowsTests
    {
        [TestMethod]
        public void WindowCommandsUseExactCaptionAndReadBackTheResult()
        {
            var host = new FakeVbe();
            var hidden = new FakeWindow(host)
            {
                Caption = "Immediate",
                Type = 3,
                Visible = false
            };
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
            var window = new FakeWindow(host)
            {
                Caption = "Module1",
                Type = 0,
                Visible = true,
                LinkedWindowFrame = new FakeFrame
                {
                    Caption = "Frame",
                    LinkedWindows = new List<FakeWindow>()
                }
            };
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
            var pane = new FakePane
            {
                CodeModule = new FakePaneModule
                {
                    Parent = new FakePaneComponent
                    {
                        Name = "Module1",
                        Collection = new FakePaneCollection
                        {
                            Parent = new FakePaneProject
                            {
                                Name = "Projet",
                                FileName = @"C:\Temp\Projet.xlsm"
                            }
                        }
                    }
                },
                CodePaneView = 1,
                TopLine = 9,
                CountOfVisibleLines = 30,
                Window = new FakePaneWindow
                {
                    Caption = "Module1 (Code)"
                }
            };
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
    }
}
