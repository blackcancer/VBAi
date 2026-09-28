using System.ComponentModel;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorWindowTests
    {
        [STATestMethod]
        public void DesignerAndConstructionNeverLaunchWebViewOrReadVba()
        {
            using (var window = new ModernEditorWindow()) { Assert.IsNull(window.Browser); Assert.IsFalse(window.Ready); }
            using (var window = (ModernEditorWindow)LicenseManager.CreateWithContext(typeof(ModernEditorWindow), new DesignContext()))
            { Assert.IsNull(window.Browser); Assert.AreEqual("VBAi editor", window.Text); }
        }
        [DataTestMethod]
        [DataRow("Clear All Breakpoints")]
        [DataRow("Effacer tous les points d'arrêt")]
        public void ToggleNeverAcceptsClearAllBreakpoints(string caption)
        { Assert.IsFalse(VbeDebug.IsAllowed("toggle_breakpoint", caption, 2)); }
        [STATestMethod]
        public void WorkspaceCannotBeClosedByANativeCloseCommand()
        {
            using (var window = new ModernEditorWindow())
            {
                window.WorkspaceHosted = true;
                var args = new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false);
                UiInvoke.Call(typeof(ModernEditorWindow), "ClosingWindow", window, window, args);
                Assert.IsTrue(args.Cancel);
                Assert.IsFalse(window.ControlBox);
                Assert.AreEqual(System.Windows.Forms.FormBorderStyle.None, window.FormBorderStyle);
                Assert.IsTrue(UiInvoke.Field<ThemedTabControl>(window, "tabs").ShowCloseButtons);
            }
        }
        [DataTestMethod]
        [DataRow("https://editor.vbai.local/index.html", true)]
        [DataRow("https://editor.vbai.local/index.html?external", false)]
        [DataRow("https://editor.vbai.local.evil.test/index.html", false)]
        [DataRow("http://editor.vbai.local/index.html", false)]
        [DataRow("https://user@editor.vbai.local/index.html", false)]
        public void BridgeAcceptsOnlyTheBundledDocument(string source, bool allowed)
        { Assert.AreEqual(allowed, ModernEditorWindow.Trusted(source)); }
    }
}
