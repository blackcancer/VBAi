using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
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
