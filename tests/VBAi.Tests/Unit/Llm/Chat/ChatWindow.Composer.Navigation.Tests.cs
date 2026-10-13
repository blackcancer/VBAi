using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual project picker and draft continuity without a native VBE host.</summary>
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void ProjectPickerNavigatesChosenModuleWithoutChangingDraftSelectionOrAttachments()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var prompt = Get<TextBox>(window, "prompt");
                prompt.Text = "Please review #P and keep my pending instructions."; prompt.Select(7, 6);
                string draft = prompt.Text; int start = prompt.SelectionStart, length = prompt.SelectionLength;
                var selected = Get<List<VbeChatReference>>(window, "selectedReferences");
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                var attachment = new ChatAttachment { Label = "Pending selection" }; attachments.Add(attachment);
                var project = new VbeChatReference { Project = "P", Kind = "Projet" }; selected.Add(project);
                var index = Get<VbeChatReferences>(window, "referenceIndex");
                var requests = new List<Request>();
                index.Execute = request =>
                {
                    requests.Add(request);
                    if (request.Command == "list_modules") return Response.Success(new object[] { new { Name = "M" }, new { Name = "C" } });
                    if (request.Command == "read_module") return Response.Success(new { Code = "Option Explicit", Sha256 = "live" });
                    return Response.Success(new { Selected = true });
                };
                Call(window, "NavigateReference", project);
                var popup = Get<Popup>(window, "referencePopup");
                Assert.IsTrue(popup.IsOpen); Assert.AreEqual(draft, prompt.Text);
                Assert.AreEqual(start, prompt.SelectionStart); Assert.AreEqual(length, prompt.SelectionLength);
                Call(window, "UpdateReferences"); // A background index notification must not replace the navigation list.
                var list = Get<System.Windows.Forms.ListBox>(window, "referenceList");
                CollectionAssert.AreEqual(new[] { "C", "M" }, list.Items.Cast<VbeChatReference>().Select(item => item.Module).ToArray());
                list.SelectedIndex = 1; Call(window, "AcceptReference");
                Assert.IsFalse(popup.IsOpen); Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                Assert.AreEqual(draft, prompt.Text); Assert.AreEqual(start, prompt.SelectionStart); Assert.AreEqual(length, prompt.SelectionLength);
                Assert.AreSame(project, selected.Single());
                Assert.AreSame(attachment, attachments.Single());
                CollectionAssert.AreEqual(new[] { "list_modules", "read_module", "select_code" }, requests.Select(request => request.Command).ToArray());
                Assert.AreEqual("M", requests.Last().Module); Assert.AreEqual("live", requests.Last().ExpectedSha256);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void EmptyPickerDoesNotSendAndDismissalOrTypingRestoresOrdinaryComposerBehavior()
        {
            var previous = ChatWindow.ReadModifiers;
            try
            {
                ChatWindow.ReadModifiers = () => ModifierKeys.None;
                using (var runtime = new RuntimeScope())
                using (var window = LoadedWindow(runtime.Session))
                {
                    var prompt = Get<TextBox>(window, "prompt"); prompt.Text = "Pending draft"; prompt.Select(3, 4);
                    var index = Get<VbeChatReferences>(window, "referenceIndex");
                    index.Execute = request => Response.Success(new object[0]);
                    var project = new VbeChatReference { Project = "P" };
                    Call(window, "NavigateReference", project);
                    Assert.IsTrue(RunKey(window, Key.Enter).Handled); Assert.IsTrue(RunKey(window, Key.Tab).Handled);
                    Assert.AreEqual("Pending draft", prompt.Text); Assert.AreEqual(3, prompt.SelectionStart); Assert.AreEqual(4, prompt.SelectionLength);
                    Call(window, "AcceptReference"); Assert.IsNotNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                    Assert.IsTrue(RunKey(window, Key.Escape).Handled); Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                    Call(window, "NavigateReference", project);
                    Get<Popup>(window, "referencePopup").IsOpen = false; System.Windows.Forms.Application.DoEvents();
                    Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                    Call(window, "NavigateReference", project); prompt.Text = "Edited draft";
                    Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                    Assert.IsFalse(Get<Popup>(window, "referencePopup").IsOpen);
                    Assert.AreEqual("Edited draft", prompt.Text);
                }
            }
            finally { ChatWindow.ReadModifiers = previous; }
        }

        [STATestMethod, TestCategory("Unit")]
        public void NavigationReportsBridgeRefusalsAndExceptionsWithoutTouchingTheDraft()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var prompt = Get<TextBox>(window, "prompt"); prompt.Text = "Keep this draft"; prompt.Select(1, 5);
                var index = Get<VbeChatReferences>(window, "referenceIndex");
                foreach (string kind in new[] { "project missing", "module missing", "exception", "null" })
                {
                    int calls = 0;
                    index.Execute = request => { calls++; if (kind == "exception") throw new InvalidOperationException("Native refusal"); return Response.Failure(kind); };
                    VbeChatReference target = kind == "null" ? null : new VbeChatReference { Project = "P", Module = kind == "project missing" ? null : "M" };
                    Call(window, "NavigateReference", target);
                    Assert.AreEqual("Keep this draft", prompt.Text); Assert.AreEqual(1, prompt.SelectionStart); Assert.AreEqual(5, prompt.SelectionLength);
                    Assert.AreEqual(kind == "null" ? 0 : 1, calls);
                    StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text,
                        kind == "exception" ? "Native refusal" : kind == "null" ? "reference" : kind);
                    Assert.IsFalse(Get<Popup>(window, "referencePopup").IsOpen);
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DelayedCloseCannotClearAReopenedProjectPickerAndDisposalDetachesItsObserver()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var prompt = Get<TextBox>(window, "prompt"); prompt.Text = "Unchanged draft"; prompt.Select(2, 3);
                Get<VbeChatReferences>(window, "referenceIndex").Execute = request => Response.Success(new object[0]);
                var popup = Get<Popup>(window, "referencePopup");
                for (int cycle = 0; cycle < 12; cycle++)
                {
                    var first = new VbeChatReference { Project = "First" };
                    var second = new VbeChatReference { Project = "Second" };
                    Call(window, "NavigateReference", first);
                    popup.IsOpen = false;
                    Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"), "Close admission must change synchronously.");
                    Call(window, "NavigateReference", second);
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(new Action(() => { }), System.Windows.Threading.DispatcherPriority.Background);
                    Assert.IsTrue(popup.IsOpen); Assert.AreSame(second, Get<VbeChatReference>(window, "referenceNavigationProject"));
                    Assert.AreEqual("Unchanged draft", prompt.Text); Assert.AreEqual(2, prompt.SelectionStart); Assert.AreEqual(3, prompt.SelectionLength);
                }
                window.Dispose();
                Assert.IsNull(Get<object>(window, "referencePopupOpenDescriptor"));
                Assert.IsNull(Get<VbeChatReference>(window, "referenceNavigationProject"));
                Assert.IsFalse(popup.IsOpen);
            }
        }
    }
}
