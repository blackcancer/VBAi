using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void EightRoundsPauseAndResumeKeepsResultsWithoutReplayingActions()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                int actions = 0; var json = new JavaScriptSerializer();
                var replies = Enumerable.Range(0, 8).Select(i => json.Serialize(new { choices = new[] { new { message = new { role = "assistant", tool_calls = new[] { new { id = "action-" + i, type = "function", function = new { name = "status", arguments = "{}" } } } } } } })).ToArray();
                var first = new ChatResponseHandler(replies);
                window.HttpHandlerOverride = () => first;
                ChatWindow.InvokeTool = (t, n, a) => { actions++; return Task.FromResult(json.Serialize(Response.Success(new { Action = actions }))); };
                Question(window, "perform work"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                Assert.IsTrue(state.BudgetPaused); Assert.AreEqual(8, actions);
                string pausedTurn = state.PausedTurnId;
                Assert.AreEqual(8, state.CompletedToolActions.Count);
                Assert.IsFalse(Get<bool>(window, "busy"));
                var history = Get<List<object>>(window, "messages");
                Assert.AreEqual(8, history.Count(m => json.Serialize(m).Contains("\"role\":\"tool\"")));
                var restored = json.Deserialize<ChatSessionState>(json.Serialize(state));
                Assert.IsTrue(restored.BudgetPaused); Assert.AreEqual(state.PausedTurnId, restored.PausedTurnId);
                var resume = new ChatResponseHandler(replies[7], "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"done\"}}]}");
                window.HttpHandlerOverride = () => resume;
                var mode = state.Mode; state.Mode = ChatMode.Discussion;
                CompleteOnSta((Task)Call(window, "ResumeBudgetAsync")); Assert.IsTrue(state.BudgetPaused); Assert.AreEqual(0, resume.Requests.Count);
                state.Mode = mode;
                CompleteOnSta((Task)Call(window, "ResumeBudgetAsync"));
                Assert.IsFalse(state.BudgetPaused); Assert.AreEqual(8, actions);
                Assert.AreEqual(2, resume.Requests.Count);
                StringAssert.Contains(resume.Requests[0], "action-7");
                StringAssert.Contains(resume.Requests[1], "No action was replayed");
                Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(e => e.Text == "done" && e.TurnId == pausedTurn));
                Assert.AreEqual(1, history.Count(m => json.Serialize(m).Contains("\"role\":\"user\"")));
                Set(window, "currentSession", null);
            }
        }
    }
}
