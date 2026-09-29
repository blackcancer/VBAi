using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed class WorkspaceContractHost
    {
        public object MainWindow { get; set; }
        public Func<object> ReadActive { get; set; }
        public object ActiveWindow => ReadActive();
    }

    [TestClass, TestCategory("Unit")]
    public sealed class EditorWorkspaceHostTests
    {
        [DllImport("kernel32.dll", SetLastError = true)] private static extern void SetLastError(uint error);
        [DllImport("user32.dll")] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);

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
