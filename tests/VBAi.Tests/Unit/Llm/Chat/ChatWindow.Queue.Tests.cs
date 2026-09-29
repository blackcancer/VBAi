using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void QueuedRequestDefersWhileBusyAndDiscardsAResponseReceivedAfterInterruption()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var state = Get<ChatSessionState>(window, "currentSession");
                var item = new QueuedChatMessage { Text = "interrupt queued request" };
                state.PendingMessages.Add(item); int requests = 0;
                window.CodexTurnOverride = (text, model, effort) => {
                    requests++; Set(window, "stopRequested", true); return Task.FromResult("completed after stop");
                };
                Set(window, "busy", true);
                CompleteOnSta((Task)Call(window, "SendRequestAsync", item));
                Assert.AreEqual(0, requests); Assert.AreEqual(1, state.PendingMessages.Count);
                Set(window, "busy", false);
                CompleteOnSta((Task)Call(window, "SendRequestAsync", item));
                Assert.AreEqual(1, requests); Assert.AreEqual(0, state.PendingMessages.Count);
                Assert.IsTrue(state.Entries.Any(entry => entry.Text == UiText.Get("Response interrupted. Changes already applied can still be undone in the chat.")));
                Assert.IsFalse(state.Entries.Any(entry => entry.Text == "completed after stop"));
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void PendingQueueGuardsPreserveUnknownItemsAndDetachedSessions()
        {
            var state = new ChatSessionState { Scope = "temporary:test", PendingMessages = null };
            using (var window = ReadyCodexWindow(state))
            {
                var property = typeof(ChatWindow).GetProperty("PendingMessages", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var queue = (List<QueuedChatMessage>)property.GetValue(window);
                Assert.AreSame(queue, property.GetValue(window)); Assert.AreSame(queue, state.PendingMessages);
                var unknown = new QueuedChatMessage { Text = "unknown" };
                Call(window, "DeletePendingMessage", unknown); Call(window, "EditPendingMessage", unknown);
                CompleteOnSta((Task)Call(window, "SendPendingNowAsync", unknown));
                var panel = Get<FlowLayoutPanel>(window, "pendingMessagesPanel");
                Set(window, "pendingMessagesPanel", null); Call(window, "RefreshPendingMessages"); Set(window, "pendingMessagesPanel", panel);
                Question(window, " "); Call(window, "QueueComposerMessage"); Assert.AreEqual(0, queue.Count);
                Set(window, "currentSession", null);
                Assert.IsNull(property.GetValue(window));
                Question(window, ""); CompleteOnSta((Task)Call(window, "SendRequestAsync", (object)null));
                Call(window, "QueueComposerMessage"); Call(window, "RefreshPendingMessages");
                Call(window, "DeletePendingMessage", unknown); Call(window, "EditPendingMessage", unknown);
                CompleteOnSta((Task)Call(window, "SendPendingNowAsync", unknown));
                CompleteOnSta((Task)Call(window, "DispatchPendingAsync", true));
                Assert.AreEqual(0, queue.Count); Assert.IsFalse(panel.Visible);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void QueueCapturesTheSelectedMemoryAndRefusesAClosedScopeWithoutClearingDraft()
        {
            foreach (string memory in new[] { null, "captured draft memory" })
            foreach (bool include in new[] { false, true })
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                Set(window, "projectMemory", "project memory"); Set(window, "queuedDraftMemory", memory);
                Get<CheckBox>(window, "attachMemory").Checked = include;
                Question(window, " queued question "); Call(window, "QueueComposerMessage");
                var state = Get<ChatSessionState>(window, "currentSession");
                Assert.AreEqual(include ? memory ?? "project memory" : null, state.PendingMessages.Single().Memory);
                Assert.AreEqual("queued question", state.PendingMessages.Single().Text);
                Assert.IsNull(Get<string>(window, "queuedDraftMemory"));
                Set(window, "scopeSession", new VbeSession(new object())); Question(window, "keep this draft");
                Call(window, "QueueComposerMessage");
                Assert.AreEqual(1, state.PendingMessages.Count); Assert.AreEqual("keep this draft", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                Set(window, "scopeSession", null); Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void QueueEditingPreservesEveryKindOfDraftAndRestoresNullableCapturedContext()
        {
            foreach (string draft in new[] { "text", "attachment", "reference", "memory", "empty" })
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var state = Get<ChatSessionState>(window, "currentSession");
                var item = new QueuedChatMessage { Text = "queued", References = null, Attachments = null, Memory = null };
                state.PendingMessages.Add(item); Set(window, "immediateMessageId", item.Id);
                if (draft == "text") Question(window, "unsent text");
                if (draft == "attachment") Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "draft", Text = "code" });
                if (draft == "reference") Get<List<VbeChatReference>>(window, "selectedReferences").Add(new VbeChatReference());
                if (draft == "memory") Get<CheckBox>(window, "attachMemory").Checked = true;
                Call(window, "EditPendingMessage", item);
                Assert.AreEqual(draft == "empty" ? 0 : 1, state.PendingMessages.Count);
                if (draft == "empty") { Assert.AreEqual("queued", state.Draft); Assert.IsNull(Get<string>(window, "immediateMessageId")); }
                else { Assert.AreEqual(item.Id, Get<string>(window, "immediateMessageId")); Call(window, "DeletePendingMessage", item); Assert.IsNull(Get<string>(window, "immediateMessageId")); }
                Set(window, "currentSession", null);
            }
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var state = Get<ChatSessionState>(window, "currentSession");
                var item = new QueuedChatMessage { Text = "restore", References = new VbeChatReference[0], Attachments = new ChatAttachment[0], Memory = "captured" };
                state.PendingMessages.Add(item); Set(window, "immediateMessageId", "different");
                Call(window, "EditPendingMessage", item);
                Assert.AreEqual("captured", Get<string>(window, "queuedDraftMemory")); Assert.IsTrue(Get<CheckBox>(window, "attachMemory").Checked);
                Assert.AreEqual("different", Get<string>(window, "immediateMessageId")); Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void QueuedRequestsUseCapturedNullablePayloadWhilePreservingTheCurrentComposer()
        {
            foreach (bool nullable in new[] { false, true })
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var state = Get<ChatSessionState>(window, "currentSession");
                var item = new QueuedChatMessage { Text = "queued request", References = nullable ? null : new VbeChatReference[0], Attachments = nullable ? null : new ChatAttachment[0], Memory = nullable ? null : "captured memory" };
                state.PendingMessages.Add(item); string request = null;
                if (!nullable) { var effort = Get<ComboBox>(window, "effortPicker"); effort.Items.Add(new LlmEffortOption("high", "High")); effort.SelectedIndex = 0; }
                window.CodexTurnOverride = (text, model, effort) => { request = text; return Task.FromResult("done"); };
                Question(window, "keep composer"); Set(window, "queuedDraftMemory", "unsent memory");
                CompleteOnSta((Task)Call(window, "SendPendingNowAsync", item));
                Assert.AreEqual(0, state.PendingMessages.Count); Assert.AreEqual("keep composer", state.Draft);
                Assert.AreEqual("unsent memory", Get<string>(window, "queuedDraftMemory"));
                StringAssert.Contains(request, "queued request"); Assert.AreEqual(!nullable, request.Contains("captured memory"));
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DispatchGuardsAndBudgetButtonsRespectBusyStopPauseAndDraftCombinations()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var state = Get<ChatSessionState>(window, "currentSession"); int requests = 0;
                window.CodexTurnOverride = (text, model, effort) => { requests++; return Task.FromResult("done"); };
                foreach (bool busy in new[] { false, true }) foreach (bool stop in new[] { false, true })
                foreach (bool paused in new[] { false, true }) foreach (bool draft in new[] { false, true })
                {
                    Set(window, "busy", busy); Set(window, "stopRequested", stop); state.BudgetPaused = paused;
                    Question(window, draft ? "draft" : ""); Call(window, "UpdateBudgetControls");
                    Assert.AreEqual(!busy || !stop || draft, Get<Button>(window, "send").Enabled);
                }
                Set(window, "busy", false); Set(window, "stopRequested", false); state.BudgetPaused = false;
                var prompt = Get<object>(window, "prompt"); Set(window, "prompt", null); Call(window, "UpdateBudgetControls"); Set(window, "prompt", prompt);
                var item = new QueuedChatMessage { Text = "waiting" }; state.PendingMessages.Add(item);
                Set(window, "busy", true); CompleteOnSta((Task)Call(window, "DispatchPendingAsync", true));
                Set(window, "stopRequested", true); CompleteOnSta((Task)Call(window, "SendPendingNowAsync", item));
                Assert.IsNull(Get<string>(window, "immediateMessageId"));
                Set(window, "busy", false); CompleteOnSta((Task)Call(window, "DispatchPendingAsync", true));
                Set(window, "stopRequested", false); state.BudgetPaused = true; CompleteOnSta((Task)Call(window, "DispatchPendingAsync", true));
                state.BudgetPaused = false; CompleteOnSta((Task)Call(window, "DispatchPendingAsync", false));
                Set(window, "immediateMessageId", "removed"); CompleteOnSta((Task)Call(window, "DispatchPendingAsync", false));
                Assert.IsNull(Get<string>(window, "immediateMessageId")); Assert.AreEqual(0, requests); Assert.AreEqual(1, state.PendingMessages.Count);
                Set(window, "currentSession", null); window.Dispose(); CompleteOnSta((Task)Call(window, "DispatchPendingAsync", true));
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void QueueRowsDispatchTheirActualActionsAndFollowTheirContainerWidth()
        {
            foreach (string action in new[] { "sendNow", "edit", "delete" })
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                int requests = 0; window.CodexTurnOverride = (text, model, effort) => { requests++; return Task.FromResult("done"); };
                var state = Get<ChatSessionState>(window, "currentSession");
                state.PendingMessages.Add(new QueuedChatMessage { Text = "action target" }); Call(window, "RefreshPendingMessages");
                var panel = Get<FlowLayoutPanel>(window, "pendingMessagesPanel"); var row = panel.Controls[0];
                panel.Width = 640; Assert.AreEqual(Math.Max(200, panel.ClientSize.Width - 24), row.Width);
                var button = (Button)row.GetType().GetField(action, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(row);
                typeof(Button).GetMethod("OnClick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    .Invoke(button, new object[] { EventArgs.Empty });
                Assert.AreEqual(0, state.PendingMessages.Count); Assert.AreEqual(action == "sendNow" ? 1 : 0, requests);
                if (action == "edit") Assert.AreEqual("action target", state.Draft);
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void QueuedMessagesRunInOrderAndPreserveUnsentDraft()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var pending = new TaskCompletionSource<string>();
                var requests = new List<string>();
                window.CodexTurnOverride = (text, model, effort) => { requests.Add(text); return requests.Count == 1 ? pending.Task : Task.FromResult("done"); };
                Question(window, "first"); var run = (Task)Call(window, "SendAsync");
                StringAssert.Contains(Get<Button>(window, "send").Text, UiText.Get("Stop ■"));
                Question(window, "second"); StringAssert.Contains(Get<Button>(window, "send").Text, UiText.Get("Queue ↑"));
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "snippet", Text = "captured code" });
                CompleteOnSta((Task)Call(window, "SendAsync"));
                Question(window, "third"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                Assert.AreEqual(2, state.PendingMessages.Count); Assert.AreEqual(1, requests.Count);
                Assert.AreEqual(2, Get<FlowLayoutPanel>(window, "pendingMessagesPanel").Controls.Count);
                var restored = new JavaScriptSerializer().Deserialize<ChatSessionState>(new JavaScriptSerializer().Serialize(state));
                Assert.AreEqual("captured code", restored.PendingMessages[0].Attachments[0].Text);
                Question(window, "unsent draft"); pending.SetResult("first done"); CompleteOnSta(run);
                Assert.AreEqual(3, requests.Count); StringAssert.Contains(requests[1], "second"); StringAssert.Contains(requests[1], "captured code"); StringAssert.Contains(requests[2], "third");
                Assert.AreEqual("unsent draft", state.Draft); Assert.AreEqual(0, state.PendingMessages.Count);
                Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void SendNowWaitsForInterruptionAndPrioritizesSelectedMessage()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var pending = new TaskCompletionSource<string>(); int requests = 0; bool interrupted = false;
                window.CodexTurnOverride = (text, model, effort) => {
                    requests++; if (requests == 1) return pending.Task;
                    Assert.IsTrue(interrupted); if (requests == 2) StringAssert.Contains(text, "priority");
                    return Task.FromResult("done");
                };
                window.CodexInterruptOverride = () => { interrupted = true; pending.SetException(new OperationCanceledException()); return Task.CompletedTask; };
                Question(window, "running"); var run = (Task)Call(window, "SendAsync");
                Question(window, "later"); CompleteOnSta((Task)Call(window, "SendAsync"));
                Question(window, "priority"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                CompleteOnSta((Task)Call(window, "SendPendingNowAsync", state.PendingMessages[1])); CompleteOnSta(run);
                Assert.AreEqual(3, requests); Assert.AreEqual(0, state.PendingMessages.Count);
                Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void StopKeepsQueueAndEditDeleteNeverOverwriteCurrentDraft()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var pending = new TaskCompletionSource<string>();
                window.CodexTurnOverride = (text, model, effort) => pending.Task;
                window.CodexInterruptOverride = () => { pending.SetException(new OperationCanceledException()); return Task.CompletedTask; };
                Question(window, "running"); var run = (Task)Call(window, "SendAsync");
                Question(window, "queued"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession"); var item = state.PendingMessages[0];
                Question(window, "draft"); Call(window, "EditPendingMessage", item);
                Assert.AreEqual(1, state.PendingMessages.Count);
                Question(window, ""); CompleteOnSta((Task)Call(window, "SendAsync")); CompleteOnSta(run);
                Assert.AreEqual(1, state.PendingMessages.Count); Assert.IsFalse(Get<bool>(window, "busy"));
                Call(window, "EditPendingMessage", item); Assert.AreEqual("queued", state.Draft); Assert.AreEqual(0, state.PendingMessages.Count);
                // Requeue while busy then delete, without sending to the provider.
                Call(window, "SetBusy", true); CompleteOnSta((Task)Call(window, "SendAsync"));
                Call(window, "DeletePendingMessage", state.PendingMessages[0]); Assert.AreEqual(0, state.PendingMessages.Count);
                Call(window, "SetBusy", false); Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void ProgressingHttpWorkContinuesBeyondEightRounds()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                int actions = 0; var json = new JavaScriptSerializer();
                var replies = Enumerable.Range(0, 10).Select(i => json.Serialize(new { choices = new[] { new { message = new { role = "assistant", tool_calls = new[] { new { id = "step-" + i, type = "function", function = new { name = "status", arguments = "{}" } } } } } } })).Concat(new[] { "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"finished\"}}]}" }).ToArray();
                window.HttpHandlerOverride = () => new ChatResponseHandler(replies);
                ChatWindow.InvokeTool = (t,n,a) => Task.FromResult(json.Serialize(Response.Success(new { Step = ++actions })));
                Question(window, "long work"); CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(10, actions); Assert.IsFalse(Get<ChatSessionState>(window, "currentSession").BudgetPaused);
                Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(e => e.Text == "finished"));
                Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void FailedImmediateStopKeepsMessageQueuedUntilTheActiveTurnEnds()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                var pending = new TaskCompletionSource<string>(); int calls = 0;
                window.CodexTurnOverride = (text, model, effort) => ++calls == 1 ? pending.Task : Task.FromResult("done");
                window.CodexInterruptOverride = () => Task.FromException(new InvalidOperationException("cannot interrupt"));
                Question(window, "running"); var run = (Task)Call(window, "SendAsync");
                Question(window, "pending"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                CompleteOnSta((Task)Call(window, "SendPendingNowAsync", state.PendingMessages[0]));
                Assert.AreEqual(1, calls); Assert.AreEqual(1, state.PendingMessages.Count); Assert.IsNull(Get<string>(window, "immediateMessageId"));
                pending.SetResult("finished"); CompleteOnSta(run); Assert.AreEqual(2, calls);
                Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void ContinuouslyChangingHttpLoopStillHasAHardSafetyCeiling()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                int actions = 0; var json = new JavaScriptSerializer();
                var replies = Enumerable.Range(0, 64).Select(i => json.Serialize(new { choices = new[] { new { message = new { role = "assistant", tool_calls = new[] { new { id = "bounded-" + i, type = "function", function = new { name = "status", arguments = "{}" } } } } } } })).ToArray();
                var handler = new ChatResponseHandler(replies); window.HttpHandlerOverride = () => handler;
                ChatWindow.InvokeTool = (t,n,a) => Task.FromResult(json.Serialize(Response.Success(new { Step = ++actions })));
                Question(window, "bounded work"); CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(64, actions); Assert.AreEqual(64, handler.Requests.Count);
                Assert.IsTrue(Get<ChatSessionState>(window, "currentSession").BudgetPaused); Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }
    }
}
