using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed partial class AddInShutdownObservationTests
    {
        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ActualChatCleanupKeepsItsExactNestedParentAndReturnsBeforeDispatcherCleanup(bool begin)
        {
            using (var host = new HostUiScope())
            using (var scope = new ObservationScope())
            {
                var instance = host.Connected();
                var chat = (ChatWindow)Get(instance, "chat");
                var dispatcher = (Control)Get(instance, "dispatcher");
                try
                {
                    Shutdown(instance, begin);
                    Assert.IsTrue(chat.IsDisposed); Assert.IsTrue(dispatcher.IsDisposed);
                    Assert.IsTrue(scope.Roots.All(root => !root.DiagnosticFailed));
                    var invocation = scope.Events.Single(row => Is(row, "Entry", "chat.Dispose"));
                    var child = scope.Events.Single(row => Is(row, "Entry", "chat.Runtime.contextMonitor.Stop"));
                    Assert.AreNotEqual(invocation["InvocationId"], child["InvocationId"]);
                    var childEntry = scope.Events.Single(row => Is(row, "Entry", null) && row["InvocationId"].Equals(child["InvocationId"]));
                    Assert.AreEqual(invocation["InvocationId"], childEntry["ParentInvocationId"]);
                    Before(scope, "Returned", "chat.Base.Dispose", "Returned", "chat.Dispose");
                    Before(scope, "Returned", "chat.Dispose", "Entry", "dispatcher.Dispose");
                    Assert.IsFalse(scope.Events.Any(row => Is(row, "Fault", null)));
                }
                finally { AddIn.CreateShutdownDiagnostic = entry => null; host.Close(instance); }
            }
        }

        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ReleasedEditorControllerDoesNotAbortTheRemainingOwnedShutdownResources(bool begin)
        {
            using (var host = new HostUiScope())
            using (var scope = new ObservationScope())
            using (var editor = new Editor.ModernEditorBrowserFixture())
            {
                var instance = host.Connected();
                ModernEditorDebugFixture.Wait(editor.Initialize());
                Set(instance, "modernEditor", editor.Window);
                int controllerCalls = 0;
                editor.Controller.BeforeCall = name =>
                {
                    if (name == "get_CoreWebView2") { controllerCalls++; throw new InvalidComObjectException("Released shutdown controller"); }
                };
                var chat = (ChatWindow)Get(instance, "chat");
                var dispatcher = (Control)Get(instance, "dispatcher");
                try
                {
                    Shutdown(instance, begin);
                    Assert.AreEqual(1, controllerCalls);
                    Assert.IsTrue(editor.Window.IsDisposed);
                    Assert.IsTrue(chat.IsDisposed);
                    Assert.IsTrue(dispatcher.IsDisposed);
                    foreach (string field in new[] { "modernEditor", "chat", "nativeChatControl", "nativeChatWindow", "server", "dispatcher", "vbe" })
                        Assert.IsNull(Get(instance, field), field);
                    Before(scope, "Returned", "modernEditor.Dispose", "Entry", "chat.Dispose");
                    Assert.IsFalse(scope.Events.Any(row => Is(row, "Fault", null)));
                    Shutdown(instance, !begin);
                    Assert.AreEqual(1, controllerCalls);
                }
                finally { AddIn.CreateShutdownDiagnostic = entry => null; host.Close(instance); }
            }
        }

        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ReleasedNativeTestWindowIsAttemptedOnceAndDoesNotPreventDispatcherCleanup(bool begin)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var instance = new AddIn();
                var native = new ObservedNativeWindow { Error = new InvalidComObjectException("Released VBE test site") };
                Set(instance, "nativeTestWindow", native); Set(instance, "dispatcher", dispatcher);
                Shutdown(instance, begin);
                Assert.AreEqual(1, native.Calls);
                Assert.IsTrue(dispatcher.IsDisposed);
                Assert.IsNull(Get(instance, "nativeTestWindow"));
                Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Fault", "nativeTestWindow.Close")));
                Assert.IsFalse(scope.Events.Any(row => Is(row, "Fault", null)));
                Shutdown(instance, !begin);
                Assert.AreEqual(1, native.Calls);
            }
        }

        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void EmbeddedTestShutdownDoesNotRestoreFloatingStylesOrCreateNewHandles(bool begin)
        {
            using (var scope = new ObservationScope())
            using (var site = new ChatToolWindow())
            using (var window = new TestExplorerWindow())
            {
                window.TopLevel = false; window.FormBorderStyle = FormBorderStyle.None; window.Dock = DockStyle.Fill;
                site.Controls.Add(window);
                int created = 0;
                window.HandleCreated += (sender, args) => created++;
                window.Disposed += (sender, args) =>
                {
                    Assert.IsFalse(window.TopLevel);
                    Assert.AreEqual(FormBorderStyle.None, window.FormBorderStyle);
                    Assert.AreEqual(DockStyle.Fill, window.Dock);
                    Assert.IsNull(window.Parent);
                };
                var instance = new AddIn();
                Set(instance, "nativeTestControl", site); Set(instance, "testExplorerWindow", window);
                scope.BeforePublish = row =>
                {
                    if (Is(row, "Returned", "nativeTestControl.Detach"))
                    {
                        Assert.AreSame(site, window.Parent, "The live child must be disposed in place, without notifying a released site by reparenting it.");
                        Assert.IsFalse(window.IsDisposed);
                    }
                };
                Shutdown(instance, begin);
                Assert.AreEqual(0, created);
                Assert.IsTrue(window.IsDisposed);
                Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Returned", "nativeTestControl.Detach")));
            }
        }
    }
}
