using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises actual SDK and WinForms disposal with owned raw-controller boundaries, without launching Chromium.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ModernEditorWindowDesignerTests
    {
        [STATestMethod]
        [DataRow("normal")]
        [DataRow("detached-getter")]
        [DataRow("detached-close")]
        public void ControllerDisposalCompletesManagedLifetimePreservesDraftsAndDoesNotReplayNativeClose(string boundary)
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                ModernEditorDebugFixture.Wait(f.Initialize());
                var browser = f.Window.Browser;
                var profile = f.Profiles.Single();
                var components = f.Editor.Base.Get<IContainer>("components");
                var native = f.Editor.Base.Native;
                native.Adapter.RetainNativeWindow();
                var draft = f.Editor.Base.Document;
                draft.Edit(draft.Text + "\n' unsynchronized shutdown draft");
                string original = native.Original.CodeModule.Raw;
                int getterFaults = 0, closeAttempts = 0;
                f.Controller.BeforeCall = name =>
                {
                    if (name == "Close") closeAttempts++;
                    if ((boundary == "detached-getter" && name == "get_CoreWebView2") ||
                        (boundary == "detached-close" && name == "Close"))
                    {
                        getterFaults++;
                        throw new InvalidComObjectException("Owned controller already released");
                    }
                };

                f.Window.Dispose();

                Assert.IsTrue(f.Window.IsDisposed);
                Assert.IsTrue(browser.IsDisposed, "SDK managed disposal must finish after its controller is released.");
                Assert.AreEqual(0, components.Components.Count);
                Assert.AreEqual(1, native.Original.CodeModule.CodePane.Window.Closes);
                Assert.AreEqual(original, native.Original.CodeModule.Raw);
                Assert.AreEqual(draft.Text, f.Window.Drafts.Recover(draft.RecoveryKey).Text);
                Assert.AreEqual(boundary == "normal" ? 0 : 1, getterFaults, "A detached raw boundary must not be replayed.");
                Assert.AreEqual(boundary == "detached-getter" ? 0 : 1, closeAttempts);
                Assert.AreEqual(boundary == "normal" ? 1 : 0, f.NativeCloses);
                Assert.IsTrue(Directory.Exists(profile.Path), "A controller fault is not proof of browser-process exit.");
                f.Window.Dispose();
                Assert.AreEqual(1, native.Original.CodeModule.CodePane.Window.Closes);
                Assert.AreEqual(boundary == "normal" ? 0 : 1, getterFaults);
                Assert.AreEqual(boundary == "detached-getter" ? 0 : 1, closeAttempts);
                f.BrowserExited(Editor.ModernEditorBrowserFixture.BrowserProcessId);
                ModernEditorDebugFixture.Wait(profile.Cleanup);
                Assert.IsFalse(Directory.Exists(profile.Path));
            }
        }

        private sealed class FailingBrowser : WebView2
        {
            internal readonly IOException Failure = new IOException("Unrelated managed browser disposal failure");
            internal int Attempts;
            internal bool Fail = true;
            protected override void Dispose(bool disposing)
            {
                if (disposing) { Attempts++; if (Fail) throw Failure; }
                base.Dispose(disposing);
            }
        }

        [STATestMethod]
        public void UnrelatedBrowserFailurePropagatesWithoutRepeatingDisposalOrPretendingTheFormIsClosed()
        {
            using (var f = new ModernEditorDebugFixture())
            using (var browser = new FailingBrowser())
            {
                typeof(ModernEditorWindow).GetProperty("Browser", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(f.Window, browser);
                f.Get<System.Windows.Forms.Panel>("surface").Controls.Add(browser);
                try
                {
                    Assert.AreSame(browser.Failure, Assert.ThrowsException<IOException>(() => f.Window.Dispose()));
                    Assert.AreEqual(1, browser.Attempts);
                    Assert.IsFalse(f.Window.IsDisposed);
                    Assert.IsFalse(browser.IsDisposed);
                }
                finally { browser.Fail = false; }
            }
        }
    }
}
