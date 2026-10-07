namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using VBAi;

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
            Assert.ThrowsException<InvalidOperationException>(() => editor.CloseWindow("VBAi", 3));
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
            host.AddIns.Add(new FakeAddIn { ProgId = "VBAi.AddIn", Guid = "test", Description = "Test", Connect = true });
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
            Assert.AreEqual("VBAi.AddIn", (string)addIns.AddIns[0].Properties["ProgId"]);
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
        [TestMethod]
        public void FocusAndShowMatrixReportNullMismatchedAndSuccessfulActiveWindowReadback()
        {
            var host = new FakeVbe();
            var target = new FakeWindow(host) { Caption = "Target", Type = 3, Visible = true, SuppressFocus = true }; host.Windows.Add(target);
            var editor = new VbeEditorWindows(host);
            Assert.ThrowsException<ArgumentException>(() => editor.ShowWindow("Target", -1));
            host.Windows.Add(new FakeWindow(host) { Caption = "Other", Type = 3, Visible = true });
            host.Windows.Add(new FakeWindow(host) { Caption = "Target", Type = 4, Visible = true });
            foreach (FakeWindow active in new[] { null, host.Windows[1], host.Windows[2], target })
            {
                target.FocusTarget = active;
                dynamic focus = editor.FocusWindow("Target", 3);
                Assert.AreEqual(active == target ? "ActiveWindowReadback" : "Unverified", (string)focus.Verification);
                Assert.AreEqual(active == null, focus.ActiveWindow == null);
                dynamic shown = editor.ShowWindow("Target", 3);
                Assert.AreEqual(active == target, (bool)shown.FocusVerified);
                Assert.AreEqual(active == null, shown.ActiveWindow == null);
                Assert.IsTrue((bool)shown.WasVisible);
            }
            target.Visible = false; target.IgnoreVisibilityChanges = true;
            Assert.ThrowsException<InvalidOperationException>(() => editor.ShowWindow("Target", 3));
        }

        [TestMethod]
        public void CloseMatrixDistinguishesRemovedHiddenStillVisibleAndDuplicateWindows()
        {
            var host = new FakeVbe(); var editor = new VbeEditorWindows(host);
            var target = new FakeWindow(host) { Caption = "Target", Type = 3, Visible = true }; host.Windows.Add(target);
            host.Windows.Add(new FakeWindow(host) { Caption = "Other", Type = 3 });
            host.Windows.Add(new FakeWindow(host) { Caption = "Target", Type = 4 });
            target.OnClose = () => { };
            dynamic unchanged = editor.CloseWindow("Target", 3); Assert.AreEqual("Unverified", (string)unchanged.Verification); Assert.IsTrue((bool)unchanged.RemainingVisible);
            target.OnClose = () => { target.Visible = false; host.Windows.Add(new FakeWindow(host) { Caption = "Target", Type = 3 }); };
            dynamic duplicate = editor.CloseWindow("Target", 3); Assert.AreEqual("Unverified", (string)duplicate.Verification); Assert.AreEqual(2, (int)duplicate.RemainingMatches);
            host.Windows.RemoveAt(host.Windows.Count - 1); target.Visible = true;
            target.OnClose = () => host.Windows.Remove(target);
            dynamic removed = editor.CloseWindow("Target", 3); Assert.AreEqual("RemovedFromWindows", (string)removed.Verification); Assert.AreEqual(0, (int)removed.RemainingMatches);
        }

        [TestMethod]
        public void SnapshotMatrixReportsUnavailableActiveComObjectsAndLinkageErrors()
        {
            var host = new FakeVbe(); var editor = new VbeEditorWindows(host);
            dynamic empty = editor.Windows(); Assert.IsNull(empty.ActiveWindow); Assert.AreEqual(0, (int)empty.Windows.Count);
            host.ActiveWindowThrows = true; host.ActiveCodePaneThrows = true;
            dynamic windows = editor.Windows(); StringAssert.Contains((string)windows.ActiveWindow.Error, "unavailable");
            dynamic panes = editor.CodePanes(); StringAssert.Contains((string)panes.ActiveCodePane.Error, "unavailable");
            host.ActiveWindowThrows = false; host.ActiveCodePaneThrows = false;
            var target = new FakeWindow(host) { Caption = "Target", Type = 0, Visible = true }; host.Windows.Add(target);
            dynamic linkage = editor.WindowLinkage("Target", 0); Assert.AreEqual(false, (bool)linkage.Properties["IsLinked"]);
            target.LinkageThrows = true; linkage = editor.WindowLinkage("Target", 0); StringAssert.Contains((string)linkage.Errors["LinkedWindowFrame"], "unavailable");
            dynamic environment = editor.Environment(); Assert.IsTrue(environment.Errors.ContainsKey("ActiveProject"));
        }

    }
}
