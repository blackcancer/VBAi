using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed class WorkspaceContractHost
    {
        public object MainWindow { get; set; }
        public Func<object> ReadActive { get; set; }
        public object ActiveWindow => ReadActive();
        public object[] Windows { get; set; } = new object[0];
    }

    public sealed class WorkspaceDocumentWindow
    {
        public int Type { get; set; }
        public long HWnd { get; set; }
        public string Caption { get; set; }
    }

    [TestClass, TestCategory("Unit")]
    public sealed class EditorWorkspaceHostTests
    {
        [DllImport("kernel32.dll", SetLastError = true)] private static extern void SetLastError(uint error);
        [DllImport("user32.dll")] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);

        [STATestMethod]
        public void NewlyOpenedNativeCodePaneStaysBehindMonacoWithoutCoveringObjectBrowser()
        {
            using (var f = new AddInModernEditorFixture())
            using (var code = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None })
            {
                var editor = f.Get();
                var mdi = f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single();
                SetParent(code.Handle, mdi.Handle); code.Show();
                Assert.IsTrue(SetWindowPos(code.Handle, IntPtr.Zero, 0, 0, 300, 250, 0x0010));
                Assert.AreEqual(code.Handle, GetWindow(editor.Handle, 3));
                var workspace = LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace");
                f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 0 };
                LlmBoundaryScope.Call(workspace, "Resize");
                Assert.AreEqual(editor.Handle, GetWindow(code.Handle, 3));
                Assert.IsTrue(code.Visible); Assert.IsTrue(editor.Visible);
                f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 2 };
                LlmBoundaryScope.Call(workspace, "Resize");
                Assert.IsFalse(editor.Visible); Assert.IsTrue(code.Visible);
                f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 0 };
                LlmBoundaryScope.Call(workspace, "Resize");
                Assert.IsTrue(editor.Visible); Assert.AreEqual(editor.Handle, GetWindow(code.Handle, 3));
            }
        }

        [STATestMethod]
        public void NativeInspectionTimerUpdatesGeometryWithoutReadingOrChangingTheSelectedDocument()
        {
            foreach (bool visible in new[] { true, false })
                using (var f = new AddInModernEditorFixture())
                using (var code = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None })
                using (var input = new TextBox())
                {
                    var editor = f.Get();
                    LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace").Dispose();
                    var owner = f.Scope.Host.Owner;
                    var mdi = owner.Controls.OfType<MdiClient>().Single();
                    code.Controls.Add(input);
                    SetParent(code.Handle, mdi.Handle); code.Show();
                    Assert.IsTrue(SetWindowPos(code.Handle, IntPtr.Zero, 0, 0, 300, 250, 0x0010));
                    if (visible) editor.Show(); else editor.Hide();
                    SetFocus(input.Handle);
                    int activeReads = 0;
                    var host = new WorkspaceContractHost
                    {
                        MainWindow = f.Scope.Host.MainWindow,
                        ReadActive = () =>
                        {
                            Assert.IsFalse(VbeDebugInspection.IsActive, "A timer must not inspect or activate documents during native dispatch.");
                            activeReads++;
                            return new AddInEditorActiveWindow { Type = visible ? 2 : 0 };
                        }
                    };
                    using (var workspace = new EditorWorkspaceHost(host, editor))
                    using (new VbeDebugInspection())
                    {
                        var initialBounds = editor.Bounds;
                        owner.ClientSize = new System.Drawing.Size(owner.ClientSize.Width + 120, owner.ClientSize.Height + 80);
                        owner.PerformLayout();
                        var expectedBounds = mdi.ClientRectangle;
                        Assert.AreNotEqual(initialBounds, expectedBounds);
                        var previous = GetWindow(code.Handle, 3);
                        var next = GetWindow(code.Handle, 2);
                        var focus = GetFocus();
                        Assert.AreEqual(input.Handle, focus);
                        var codeBounds = code.Bounds;
                        for (int tick = 0; tick < 3; tick++) LlmBoundaryScope.Call(workspace, "Resize");
                        Assert.AreEqual(expectedBounds, editor.Bounds);
                        Assert.AreEqual(visible, editor.Visible, "Reservation must preserve both hidden and visible editor states.");
                        Assert.AreEqual(0, activeReads);
                        Assert.AreEqual(previous, GetWindow(code.Handle, 3));
                        Assert.AreEqual(next, GetWindow(code.Handle, 2));
                        Assert.AreEqual(focus, GetFocus());
                        Assert.AreEqual(codeBounds, code.Bounds);
                        Assert.IsTrue(code.Visible);
                    }
                    Assert.IsFalse(VbeDebugInspection.IsActive);
                    using (var resumed = new EditorWorkspaceHost(host, editor))
                    {
                        LlmBoundaryScope.Call(resumed, "Resize");
                        Assert.AreEqual(1, activeReads, "Ordinary workspace observation must resume after reservation disposal.");
                        Assert.AreEqual(!visible, editor.Visible);
                    }
                }
        }
        [STATestMethod]
        public void TimerResizePreservesNativePaneZOrderDuringFrameDocking()
        {
            using (var f = new AddInModernEditorFixture())
            using (var browser = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None })
            {
                var editor = f.Get();
                var mdi = f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single();
                SetParent(browser.Handle, mdi.Handle); browser.Show();
                Assert.IsTrue(SetWindowPos(browser.Handle, IntPtr.Zero, 0, 0, 300, 250, 0x0010));
                Assert.AreEqual(browser.Handle, GetWindow(editor.Handle, 3));
                f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 11 };
                var workspace = LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace");
                for (int pass = 0; pass < 3; pass++) LlmBoundaryScope.Call(workspace, "Resize");
                Assert.AreEqual(browser.Handle, GetWindow(editor.Handle, 3), "The resize timer must preserve the native pane above Monaco.");
                Assert.AreEqual(mdi.Handle, OwnedMdiWorkspace.GetParent(browser.Handle));
                Assert.AreEqual(new System.Drawing.Size(300, 250), browser.Size);
                Assert.IsTrue(browser.Visible);
                f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 2 };
                LlmBoundaryScope.Call(workspace, "Resize");
                Assert.IsFalse(editor.Visible); Assert.IsTrue(browser.Visible);
            }
        }

        /// <summary>Hidden Monaco keeps workspace geometry without moving or activating the native document.</summary>
        [STATestMethod]
        public void HiddenEditorTracksWorkspaceResizeWithoutChangingNativeDocumentBoundsZOrderOrFocus()
        {
            foreach (int type in new[] { 1, 2 })
                using (var f = new AddInModernEditorFixture())
                using (var native = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None })
                using (var input = new TextBox())
                {
                    var editor = f.Get();
                    var owner = f.Scope.Host.Owner;
                    var mdi = owner.Controls.OfType<MdiClient>().Single();
                    var workspace = LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace");
                    native.Controls.Add(input);
                    SetParent(native.Handle, mdi.Handle); native.Show();
                    Assert.IsTrue(SetWindowPos(native.Handle, IntPtr.Zero, 20, 25, 300, 250, 0x0010));
                    f.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = type };
                    LlmBoundaryScope.Call(workspace, "Resize");
                    Assert.IsFalse(editor.Visible);
                    SetFocus(input.Handle);
                    Assert.AreEqual(input.Handle, GetFocus());

                    var previousEditorBounds = editor.Bounds;
                    var nativeBounds = native.Bounds;
                    owner.ClientSize = new System.Drawing.Size(owner.ClientSize.Width + 120, owner.ClientSize.Height + 80);
                    owner.PerformLayout();
                    var expectedEditorBounds = mdi.ClientRectangle;
                    Assert.AreNotEqual(previousEditorBounds, expectedEditorBounds, "The fixture must actually resize the workspace.");
                    var nativePrevious = GetWindow(native.Handle, 3);
                    var nativeNext = GetWindow(native.Handle, 2);
                    var topChild = GetWindow(mdi.Handle, 5);
                    var focus = GetFocus();
                    Assert.AreEqual(input.Handle, focus, "Resizing the owner must leave the native document focused.");

                    LlmBoundaryScope.Call(workspace, "Resize");

                    Assert.AreEqual(expectedEditorBounds, editor.Bounds, "Hidden Monaco must follow the resized MDI client for native document type " + type + ".");
                    Assert.IsFalse(editor.Visible, "Geometry updates must keep Monaco hidden.");
                    Assert.IsTrue(native.Visible);
                    Assert.AreEqual(nativeBounds, native.Bounds, "The native document geometry must be preserved.");
                    Assert.AreEqual(mdi.Handle, OwnedMdiWorkspace.GetParent(native.Handle));
                    Assert.AreEqual(nativePrevious, GetWindow(native.Handle, 3));
                    Assert.AreEqual(nativeNext, GetWindow(native.Handle, 2));
                    Assert.AreEqual(topChild, GetWindow(mdi.Handle, 5), "Resizing hidden Monaco must preserve native child ordering.");
                    Assert.AreEqual(focus, GetFocus(), "Resizing hidden Monaco must not take keyboard focus.");
                }
        }

        [STATestMethod]
        public void FocusedToolWindowStillUsesTheActiveMdiCodeChildWithoutTakingKeyboardFocus()
        {
            using (var f = new AddInModernEditorFixture())
            using (var code = new Form { Text = "Owned code", MdiParent = f.Scope.Host.Owner })
            using (var tool = new TextBox())
            {
                var editor = f.Get();
                LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace").Dispose();
                var mdi = f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single();
                code.Show(); code.Activate();
                Assert.IsTrue(SetWindowPos(code.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013));
                f.Scope.Host.Owner.Controls.Add(tool);
                tool.Show(); SetFocus(tool.Handle);
                var host = new WorkspaceContractHost
                {
                    MainWindow = f.Scope.Host.MainWindow,
                    ReadActive = () => new AddInEditorActiveWindow { Type = 15 },
                    Windows = new object[] { new WorkspaceDocumentWindow { Type = 0, HWnd = code.Handle.ToInt64(), Caption = code.Text } }
                };
                using (var workspace = new EditorWorkspaceHost(host, editor))
                {
                    Assert.IsTrue(SetWindowPos(code.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013));
                    SetFocus(tool.Handle);
                    Assert.AreEqual(tool.Handle, GetFocus());
                    Assert.AreEqual(code.Handle, GetWindow(mdi.Handle, 5));
                    LlmBoundaryScope.Call(workspace, "Resize");
                    Assert.AreEqual(editor.Handle, GetWindow(mdi.Handle, 5));
                    Assert.AreEqual(code, f.Scope.Host.Owner.ActiveMdiChild);
                    Assert.AreEqual(tool.Handle, GetFocus(), "Raising Monaco must not activate it or change keyboard focus.");
                    Assert.IsTrue(code.Visible); Assert.IsTrue(editor.Visible);
                }
            }
        }

        [STATestMethod]
        public void FocusedToolWindowLeavesTheActiveMdiObjectBrowserVisible()
        {
            using (var f = new AddInModernEditorFixture())
            using (var browser = new Form { Text = "Owned object browser", MdiParent = f.Scope.Host.Owner })
            {
                var editor = f.Get();
                LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace").Dispose();
                browser.Show(); browser.Activate();
                var host = new WorkspaceContractHost
                {
                    MainWindow = f.Scope.Host.MainWindow,
                    ReadActive = () => new AddInEditorActiveWindow { Type = 15 },
                    Windows = new object[] { new WorkspaceDocumentWindow { Type = 2, HWnd = browser.Handle.ToInt64(), Caption = browser.Text } }
                };
                using (var workspace = new EditorWorkspaceHost(host, editor))
                {
                    LlmBoundaryScope.Call(workspace, "Resize");
                    Assert.IsFalse(editor.Visible);
                    Assert.IsTrue(browser.Visible);
                    Assert.AreEqual(browser, f.Scope.Host.Owner.ActiveMdiChild);
                }
            }
        }

        [STATestMethod]
        public void ZeroComHandleNeedsAnExactUniqueCaptionBeforeRaisingTheActiveMdiChild()
        {
            using (var f = new AddInModernEditorFixture())
            using (var code = new Form { Text = "Owned code", MdiParent = f.Scope.Host.Owner })
            using (var ambiguous = new Form { Text = "Owned code", MdiParent = f.Scope.Host.Owner })
            {
                var editor = f.Get();
                LlmBoundaryScope.Get<EditorWorkspaceHost>(f.Instance, "editorWorkspace").Dispose();
                var mdi = f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single();
                var host = new WorkspaceContractHost
                {
                    MainWindow = f.Scope.Host.MainWindow,
                    ReadActive = () => new AddInEditorActiveWindow { Type = 15 },
                    Windows = new object[] { new WorkspaceDocumentWindow { Type = 0, HWnd = 0, Caption = code.Text } }
                };
                using (var workspace = new EditorWorkspaceHost(host, editor))
                {
                    code.Show(); code.Activate();
                    Assert.IsTrue(SetWindowPos(code.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013));
                    Assert.AreEqual(code.Handle, GetWindow(mdi.Handle, 5));
                    LlmBoundaryScope.Call(workspace, "Resize");
                    Assert.AreEqual(editor.Handle, GetWindow(mdi.Handle, 5));

                    ambiguous.Show(); ambiguous.Activate();
                    Assert.IsTrue(SetWindowPos(ambiguous.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013));
                    host.Windows = new object[] {
                        new WorkspaceDocumentWindow { Type = 0, HWnd = 0, Caption = ambiguous.Text },
                        new WorkspaceDocumentWindow { Type = 0, HWnd = 0, Caption = ambiguous.Text }
                    };
                    Assert.AreEqual(ambiguous.Handle, GetWindow(mdi.Handle, 5));
                    LlmBoundaryScope.Call(workspace, "Resize");
                    Assert.AreEqual(ambiguous.Handle, GetWindow(mdi.Handle, 5), "Ambiguous captions must not raise Monaco over a native document.");
                    Assert.IsTrue(ambiguous.Visible);
                }
            }
        }

        [STATestMethod]
        public void NativeParentResultsPreserveZeroSuccessAndRejectReportedFailure()
        {
            var native = EditorWorkspaceHost.ChangeParent;
            try
            {
                foreach (bool rejected in new[] { false, true })
                    using (var f = new AddInModernEditorFixture())
                    {
                        EditorWorkspaceHost.ChangeParent = (child, parent) =>
                        {
                            if (!rejected) native(child, parent);
                            SetLastError(rejected ? 5u : 0u);
                            return IntPtr.Zero;
                        };
                        if (rejected)
                        {
                            var failure = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => f.Get());
                            Assert.IsInstanceOfType(failure.InnerException, typeof(System.ComponentModel.Win32Exception));
                            Assert.AreEqual(5, ((System.ComponentModel.Win32Exception)failure.InnerException).NativeErrorCode);
                            Assert.IsTrue(f.Editors.Single().IsDisposed);
                        }
                        else
                        {
                            var editor = f.Get();
                            Assert.IsTrue(editor.Visible);
                            Assert.AreEqual(f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single().Handle, OwnedMdiWorkspace.GetParent(editor.Handle));
                        }
                    }
            }
            finally { EditorWorkspaceHost.ChangeParent = native; }
        }

        [STATestMethod]
        public void WorkspaceLifetimeAndLateNativeWindowLossStopResizingSafely()
        {
            foreach (int outcome in new[] { 0, 1, 2, 3 })
                using (var f = new AddInModernEditorFixture())
                {
                    var editor = f.Get();
                    var host = new WorkspaceContractHost { MainWindow = f.Scope.Host.MainWindow, ReadActive = () => null };
                    using (var workspace = new EditorWorkspaceHost(host, editor))
                    {
                        if (outcome == 0) editor.Dispose();
                        if (outcome == 1) f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single().Dispose();
                        if (outcome == 2) host.ReadActive = () => throw new COMException("Owned active-window failure");
                        if (outcome == 3) host.ReadActive = () => { f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single().Dispose(); return null; };
                        LlmBoundaryScope.Call(workspace, "Resize");
                        if (outcome < 2) Assert.IsFalse(LlmBoundaryScope.Get<Timer>(workspace, "timer").Enabled);
                        if (outcome == 2) Assert.IsTrue(editor.Visible);
                    }
                }
        }

        [STATestMethod]
        public void ChildEnumerationAndBrowserFocusUseTheOwnedDocumentWorkspace()
        {
            using (var f = new AddInModernEditorFixture())
            using (var child = new TextBox())
            {
                f.Scope.Host.Owner.Controls.Add(child); child.BringToFront();
                var editor = f.Get();
                var browser = new Microsoft.Web.WebView2.WinForms.WebView2();
                typeof(ModernEditorWindow).GetProperty("Browser", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(editor, browser);
                try
                {
                    editor.Controls.Add(browser);
                    f.Get();
                    Assert.AreEqual(f.Scope.Host.Owner.Controls.OfType<MdiClient>().Single().Handle, OwnedMdiWorkspace.GetParent(editor.Handle));
                    Assert.IsTrue(editor.Visible);
                }
                finally { browser.Dispose(); }
            }
        }
    }
}
