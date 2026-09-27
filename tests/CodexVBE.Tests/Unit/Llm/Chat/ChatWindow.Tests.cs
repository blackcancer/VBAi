namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void ProviderEffortChoicesFollowModelAndBusyState()
        {
            using (var window = Surfaces())
            {
                Set(window, "currentSession", new ChatSessionState { Effort = "high" });
                var codexModel = new LlmModelOption("gpt-test", "Test", false, "medium", new[] { new LlmEffortOption("medium", "Medium"), new LlmEffortOption("high", "High") });
                Call(window, "UpdateEfforts", LlmProvider.All[0], codexModel);
                var effort = Get<ComboBox>(window, "effortPicker");
                Assert.AreEqual(2, effort.Items.Count);
                Assert.AreEqual("high", ((LlmEffortOption)effort.SelectedItem).Id);
                Assert.IsTrue(effort.Enabled);
                string sendBefore = Get<Button>(window, "send").Text;
                Call(window, "SetBusy", true);
                Assert.IsFalse(effort.Enabled);
                Assert.AreNotEqual(sendBefore, Get<Button>(window, "send").Text);
                Call(window, "SetBusy", false);
                Assert.AreNotEqual("", Get<Button>(window, "send").Text);
                Call(window, "UpdateEfforts", LlmProvider.All[1], codexModel);
                Assert.AreEqual(0, effort.Items.Count);
                Assert.IsFalse(effort.Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CodexSendRecordsQuestionAnswerAndStructuredRequestWithoutNetwork()
        {
            var session = new ChatSessionState
            {
                Scope = "temporary:test",
                ResumeContext = "Earlier branch"
            };
            using (var window = ReadyCodexWindow(session))
            {
                string request = null;
                window.CodexTurnOverride = (text, model, effort) =>
                {
                    request = text;
                    Assert.AreEqual("gpt-test", model);
                    Assert.IsNull(effort);
                    return Task.FromResult("Réponse complète");
                };
                Question(window, "Corrige la procédure");
                ((Task)Call(window, "SendAsync")).GetAwaiter().GetResult();
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(2, entries.Count);
                Assert.AreEqual("Vous", entries[0].Speaker);
                Assert.AreEqual("Réponse complète", entries[1].Text);
                StringAssert.Contains(request, "Corrige la procédure");
                StringAssert.Contains(request, "Earlier branch");
                Assert.IsNull(session.ResumeContext);
                Assert.IsFalse(Get<bool>(window, "busy"));
                Assert.AreEqual("", session.Draft);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void FailedCodexSendRestoresDraftAndContextForRetry()
        {
            var session = new ChatSessionState
            {
                Scope = "temporary:test"
            };
            using (var window = ReadyCodexWindow(session))
            {
                var drafts = Get<List<ChatAttachment>>(window, "draftAttachments");
                drafts.Add(new ChatAttachment { Label = "Extrait", Text = "Sub Écrire()" });
                window.CodexTurnOverride = (text, model, effort) => Task.FromException<string>(new InvalidOperationException("provider failed"));
                Question(window, "Explique le code");
                ((Task)Call(window, "SendAsync")).GetAwaiter().GetResult();
                Assert.AreEqual("Explique le code", session.Draft);
                Assert.AreEqual(1, drafts.Count);
                Assert.AreEqual("Extrait", drafts[0].Label);
                Assert.AreEqual("Erreur", Get<List<ChatEntry>>(window, "transcriptEntries")[1].Speaker);
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void StopCodexSendRecordsInterruptionAndDoesNotRestoreDraft()
        {
            var session = new ChatSessionState
            {
                Scope = "temporary:test"
            };
            using (var window = ReadyCodexWindow(session))
            {
                var pending = new TaskCompletionSource<string>();
                window.CodexTurnOverride = (text, model, effort) => pending.Task;
                window.CodexInterruptOverride = () =>
                {
                    pending.SetException(new OperationCanceledException());
                    return Task.FromResult(0);
                };
                Question(window, "Arrête cette réponse");
                var send = (Task)Call(window, "SendAsync");
                Assert.IsTrue(Get<bool>(window, "busy"));
                CompleteOnSta((Task)Call(window, "StopTurnAsync"));
                CompleteOnSta(send);
                Assert.AreEqual("Assistant", Get<List<ChatEntry>>(window, "transcriptEntries")[1].Speaker);
                Assert.AreEqual("", session.Draft);
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void StopFailureReenablesControlForAnotherAttempt()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                Call(window, "SetBusy", true);
                window.CodexInterruptOverride = () => Task.FromException(new InvalidOperationException("stop failed"));
                ((Task)Call(window, "StopTurnAsync")).GetAwaiter().GetResult();
                Assert.IsFalse(Get<bool>(window, "stopRequested"));
                Assert.IsTrue(Get<Button>(window, "send").Enabled);
                Assert.IsTrue(Get<bool>(window, "busy"));
                Call(window, "SetBusy", false);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void SendPreflightRejectsMissingModelAndOversizedContext()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                bool invoked = false;
                window.CodexTurnOverride = (text, model, effort) =>
                {
                    invoked = true;
                    return Task.FromResult("unexpected");
                };
                var modelPicker = Get<ComboBox>(window, "modelPicker");
                modelPicker.SelectedIndex = -1;
                Question(window, "Question");
                ((Task)Call(window, "SendAsync")).GetAwaiter().GetResult();
                Assert.IsFalse(invoked);
                Assert.AreEqual(0, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                modelPicker.SelectedIndex = 0;
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "Très long", Text = new string ('x', 48001) });
                ((Task)Call(window, "SendAsync")).GetAwaiter().GetResult();
                Assert.IsFalse(invoked);
                Assert.AreEqual("Question", Get<object>(window, "prompt").GetType().GetProperty("Text").GetValue(Get<object>(window, "prompt"), null));
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModelCatalogueSelectsSavedModelAndRejectsLateStaleResult()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test", Model = "saved" }))
            {
                var first = new TaskCompletionSource<LlmModelOption[]>();
                int calls = 0;
                window.ModelCatalogueOverride = provider => ++calls == 1 ? first.Task : Task.FromResult(new[] { new LlmModelOption("other", "Other"), new LlmModelOption("saved", "Saved") });
                var stale = (Task)Call(window, "LoadModelsAsync");
                CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                first.SetResult(new[] { new LlmModelOption("stale", "Stale") });
                CompleteOnSta(stale);
                var picker = Get<ComboBox>(window, "modelPicker");
                Assert.AreEqual(2, picker.Items.Count);
                Assert.AreEqual("saved", ((LlmModelOption)picker.SelectedItem).Id);
                Assert.IsTrue(Get<ToolStripMenuItem>(window, "refreshModels").Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModelCatalogueFailureLeavesRefreshAvailable()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                window.ModelCatalogueOverride = provider => Task.FromException<LlmModelOption[]>(new InvalidOperationException("catalogue failed"));
                CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                Assert.AreEqual(0, Get<ComboBox>(window, "modelPicker").Items.Count);
                Assert.IsTrue(Get<ToolStripMenuItem>(window, "refreshModels").Enabled);
                Assert.IsFalse(Get<ComboBox>(window, "modelPicker").Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void HttpProviderCompletesToolCallAndAnswerWithoutNetwork()
        {
            var session = new ChatSessionState
            {
                Scope = "temporary:test",
                Provider = "Ollama"
            };
            using (var window = ReadyHttpWindow(session))
            {
                var handler = new ChatResponseHandler("{\"choices\":[{\"finish_reason\":\"tool_calls\",\"message\":{\"role\":\"assistant\",\"content\":null,\"tool_calls\":[{\"id\":\"call-1\",\"type\":\"function\",\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}]}}]}", "{\"choices\":[{\"finish_reason\":\"stop\",\"message\":{\"role\":\"assistant\",\"content\":\"Terminé\"}}]}");
                window.HttpHandlerOverride = () => handler;
                Question(window, "Quel est le statut ?");
                CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(2, handler.Requests.Count);
                StringAssert.Contains(handler.Requests[1], "call-1");
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual("Outil", entries[1].Speaker);
                Assert.AreEqual("Terminé", entries[2].Text);
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void HttpProviderFailureRepairsConversationHistoryForRetry()
        {
            var session = new ChatSessionState
            {
                Scope = "temporary:test",
                Provider = "Ollama"
            };
            using (var window = ReadyHttpWindow(session))
            {
                var handler = new ChatResponseHandler();
                window.HttpHandlerOverride = () => handler;
                Question(window, "Réessaie ensuite");
                CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(1, handler.Requests.Count);
                Assert.AreEqual(3, Get<List<object>>(window, "messages").Count);
                Assert.AreEqual("Erreur", Get<List<ChatEntry>>(window, "transcriptEntries")[1].Speaker);
                Assert.AreEqual("Réessaie ensuite", session.Draft);
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }
    }
}
