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
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "Très long", Text = new string('x', 48001) });
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
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeConstructorProviderModelEffortEventsAndSettingsDialogsUseNativeContracts()
        {
            Assert.AreEqual("Save", ChatWindow.WriteSettings.Method.Name); Assert.AreEqual("ShowDialog", ChatWindow.ShowModal.Method.Name); Assert.IsTrue(ChatWindow.HistoryPath().EndsWith("chat.db")); using (var native = ChatWindow.TransportFactory()) Assert.IsFalse(native.IsRunning);
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                CompleteOnSta((Task)Call(window, "LoadModelsAsync")); var provider = Get<ComboBox>(window, "providerPicker"); var model = Get<ComboBox>(window, "modelPicker"); var effort = Get<ComboBox>(window, "effortPicker"); var current = Get<ChatSessionState>(window, "currentSession");
                var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "before", "new", "after", 1); var tool = Get<LlmVbeTools>(window, "tools"); ((Action<CodeChange>)typeof(LlmVbeTools).GetField("CodeEdited", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(tool))(change); Assert.AreEqual(1, Get<List<CodeChange>>(window, "codeChanges").Count); Click(Get<Button>(window, "changes")); Get<Button>(window, "changes").ContextMenuStrip.Close();
                Set(window, "busy", true); provider.SelectedIndex = 1; Set(window, "busy", false); provider.SelectedIndex = 0; Set(window, "loadingSession", true); provider.SelectedIndex = 1; Set(window, "loadingSession", false); provider.SelectedIndex = -1; provider.SelectedIndex = 0;
                CompleteOnSta((Task)Call(window, "LoadModelsAsync")); current = Get<ChatSessionState>(window, "currentSession"); model.Items.Add(new LlmModelOption("other", "Other", false, "missing", new[] { new LlmEffortOption("low", "Low"), new LlmEffortOption("high", "High") })); current.Effort = "high"; model.SelectedIndex = model.Items.Count - 1; Assert.AreEqual("other", current.Model); effort.SelectedIndex = 0; Assert.AreEqual("low", current.Effort);
                ChatWindow.WriteSettings = s => { throw new InvalidOperationException("settings save failed"); }; model.SelectedIndex = 0; provider.SelectedIndex = 1; provider.SelectedIndex = 0;
                Set(window, "restoringSelection", true); model.SelectedIndex = 0; effort.SelectedIndex = 0; Set(window, "restoringSelection", false);
                current.Effort = "missing"; Call(window, "UpdateEfforts", LlmProvider.All[0], new LlmModelOption("x", "X", false, "missing", new[] { new LlmEffortOption("low", "Low") })); Assert.AreEqual(0, effort.SelectedIndex); Call(window, "UpdateEfforts", LlmProvider.All[0], new LlmModelOption("x", "X")); Assert.IsFalse(effort.Enabled);
                Set(window, "busy", true); window.ShowSettings(); Set(window, "busy", false); window.ShowSettings();
                foreach (var mutation in new Action<LlmSettings>[] { s => { }, s => s.OpenAiEndpoint += "/changed", s => s.OllamaEndpoint += "/changed", s => s.EncryptedOpenAiKey = "changed", s => s.ManualModelLists["Ollama"] = "manual", s => s.ProviderName = "Ollama", s => s.ProviderName = "Unknown" }) { ChatWindow.ShowModal = (d, o) => { mutation(runtime.Settings); return DialogResult.OK; }; window.ShowSettings(window); }
                ChatWindow.WriteSettings = s => runtime.Saves++; Call(window, "ResetProviderConnection");
                var handle = window.Handle; var thread = new Thread(() => Call(window, "SetStatus", "worker update")); thread.Start(); Assert.IsTrue(thread.Join(5000)); Application.DoEvents(); Assert.AreEqual("worker update", Get<Label>(window, "status").Text); Set(window, "storageFailed", true); Call(window, "SetStatus", "storage"); StringAssert.Contains(Get<Label>(window, "status").Text, UiText.Get("History not saved"));
            }
            using (var runtime = new RuntimeScope()) { ChatWindow.ReadSettings = () => { throw new InvalidOperationException("settings load failed"); }; runtime.Host = r => Response.Success(new object[0]); using (var window = new ChatWindow(runtime.Session)) { Assert.AreEqual(0, Get<ComboBox>(window, "providerPicker").SelectedIndex); } }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeModelCatalogueHandlesMissingEmptyDefaultsFallbacksAndLateDisposedRequests()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var picker = Get<ComboBox>(window, "providerPicker"); var model = Get<ComboBox>(window, "modelPicker");
                picker.SelectedIndex = -1; CompleteOnSta((Task)Call(window, "LoadModelsAsync")); Assert.AreEqual(0, model.Items.Count);
                foreach (var provider in LlmProvider.All) { picker.SelectedItem = provider; CompleteOnSta((Task)Call(window, "LoadModelsAsync")); Assert.AreEqual(provider.Available, model.Items.Count > 0); }
                picker.SelectedIndex = 0;
                foreach (var catalogue in new[] { new LlmModelOption[0], new[] { new LlmModelOption("first", "First") }, new[] { new LlmModelOption("first", "First"), new LlmModelOption("default", "Default", true) } }) { window.ModelCatalogueOverride = p => Task.FromResult(catalogue); CompleteOnSta((Task)Call(window, "LoadModelsAsync")); Assert.AreEqual(catalogue.Length, model.Items.Count); }
                var pending = new TaskCompletionSource<LlmModelOption[]>(); window.ModelCatalogueOverride = p => pending.Task; var load = (Task)Call(window, "LoadModelsAsync"); picker.SelectedIndex = 1; pending.TrySetResult(new LlmModelOption[0]); CompleteOnSta(load);
                window.ModelCatalogueOverride = p => Task.FromException<LlmModelOption[]>(new InvalidOperationException("failed")); CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                var disposed = new TaskCompletionSource<LlmModelOption[]>(); window.ModelCatalogueOverride = p => disposed.Task; var last = (Task)Call(window, "LoadModelsAsync"); window.Dispose(); disposed.SetResult(new LlmModelOption[0]); CompleteOnSta(last);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeCodexTransportSendsStreamsThreadReadyAndInterruptionWithoutAProcess()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                runtime.Transport.BeforeComplete = () => { var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "before", "new", "after", 1); var tools = Get<LlmVbeTools>(window, "tools"); ((Action<CodeChange>)typeof(LlmVbeTools).GetField("CodeEdited", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(tools))(change); }; Get<CheckBox>(window, "verifyAfterEdit").Checked = true; Question(window, "request"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.AreEqual("thread", Get<ChatSessionState>(window, "currentSession").CodexThreadId); Assert.IsFalse(Get<bool>(window, "busy"));
                window.Show(); System.Threading.SynchronizationContext.SetSynchronizationContext(new System.Windows.Forms.WindowsFormsSynchronizationContext()); runtime.Transport.Complete = false; Question(window, "stop request"); var task = (Task)Call(window, "SendAsync"); CompleteOnSta((Task)Call(window, "StopTurnAsync")); CompleteOnSta(task); Assert.IsTrue(Get<bool>(window, "stopRequested")); CompleteOnSta((Task)Call(window, "StopTurnAsync"));
                System.Threading.SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext()); runtime.Transport.Complete = true; runtime.Transport.FailTurn = true; Question(window, "failure"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.AreEqual("failure", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                window.CodexTurnOverride = (t, m, e) => Task.FromResult("override"); Get<ChatSessionState>(window, "currentSession").Mode = ChatMode.Discussion; Set(window, "projectMemory", "notes"); Get<CheckBox>(window, "attachMemory").Checked = true; Question(window, "analysis"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.AreEqual("notes", Get<List<ChatEntry>>(window, "transcriptEntries").Last(x => x.Speaker == "Vous").AttachedMemory);
                Question(window, ""); CompleteOnSta((Task)Call(window, "SendAsync")); Call(window, "SetBusy", true); Question(window, "busy"); CompleteOnSta((Task)Call(window, "SendAsync")); Call(window, "SetBusy", false);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeHttpResponsesValidateToolsRepairFailuresAndRespectCallLimit()
        {
            foreach (var response in new[] { "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"\"}}]}", "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"tool_calls\":[]}}]}", "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"tool_calls\":[null]}}]}", "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"tool_calls\":[{\"id\":\"bad\"}]}}]}", "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"tool_calls\":[{\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}]}}]}" })
                using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
                { window.HttpHandlerOverride = () => new ChatResponseHandler(response); Question(window, "request"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsFalse(Get<bool>(window, "busy")); Set(window, "currentSession", null); }
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                var response = "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":null,\"tool_calls\":[{\"id\":\"call\",\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}]}}]}";
                var handler = new ChatResponseHandler(Enumerable.Range(0, 9).Select(i => response.Replace("\"call\"", "\"call-" + i + "\"")).ToArray()); window.HttpHandlerOverride = () => handler; Question(window, "limit"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.AreEqual(9, handler.Requests.Count); Assert.IsTrue(Get<ChatSessionState>(window, "currentSession").BudgetPaused); Assert.AreEqual(9, Get<ChatSessionState>(window, "currentSession").CompletedToolActions.Count); Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeProviderCallbacksAndLateClientNotificationsRespectCurrentSession()
        {
            using (var runtime = new RuntimeScope())
            {
                runtime.Host = r => Response.Success(new object[0]);
                using (var window = new ChatWindow(runtime.Session))
                {
                    var providers = Get<ComboBox>(window, "providerPicker"); providers.SelectedIndex = 2; providers.SelectedIndex = 0; CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                    var model = Get<ComboBox>(window, "modelPicker"); var effort = Get<ComboBox>(window, "effortPicker"); providers.SelectedIndex = -1; model.Items.Add(new LlmModelOption("other", "Other")); model.SelectedIndex = model.Items.Count - 1; effort.Items.Add(new LlmEffortOption("other", "Other")); effort.SelectedIndex = effort.Items.Count - 1; providers.SelectedIndex = 0; CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                    model.SelectedIndex = -1; effort.Items.Add(new LlmEffortOption("extra", "Extra")); effort.SelectedIndex = effort.Items.Count - 1; model.SelectedIndex = 0; effort.SelectedIndex = -1;
                    var client = Get<CodexAppServerClient>(window, "codex"); var chat = (Action<string, string, string, bool>)typeof(CodexAppServerClient).GetField("ChatUpdate", Fields).GetValue(client); var ready = (Action<string>)typeof(CodexAppServerClient).GetField("ThreadReady", Fields).GetValue(client); ready("no-session"); Assert.IsNull(Get<ChatSessionState>(window, "currentSession")); chat("final", "late", "pending", true);
                    var state = new ChatSessionState { Scope = "temporary:P" }; Set(window, "currentSession", state); chat("final", "old", "old session", true); ready("old thread"); Assert.IsNull(state.CodexThreadId); Set(window, "currentSession", null); window.Dispose(); chat("final", "disposed", "ignored", true); ready("disposed");
                }
                using (var design = new ChatWindow()) { Call(design, "SetBusy", true); design.PrepareEditorAction("/corriger"); Assert.IsTrue(Get<bool>(design, "busy")); }
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeModelFailuresStaleExceptionsAndChangeMenusPreserveUiState()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                window.Show(); CompleteOnSta((Task)Call(window, "LoadModelsAsync")); runtime.Transport.FailModels = true; CompleteOnSta((Task)Call(window, "LoadModelsAsync")); Assert.IsNull(Get<CodexAppServerClient>(window, "codex")); runtime.Transport.FailModels = false;
                var pending = new TaskCompletionSource<LlmModelOption[]>(); window.ModelCatalogueOverride = p => pending.Task; var stale = (Task)Call(window, "LoadModelsAsync"); window.ModelCatalogueOverride = p => Task.FromResult(new[] { new LlmModelOption("fresh", "Fresh") }); CompleteOnSta((Task)Call(window, "LoadModelsAsync")); pending.SetException(new InvalidOperationException("stale failure")); CompleteOnSta(stale); Assert.AreEqual("fresh", ((LlmModelOption)Get<ComboBox>(window, "modelPicker").SelectedItem).Id);
                var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "a", "new", "b", 1); Call(window, "AddEntry", new ChatEntry { Speaker = "Code", Change = change }); Get<List<CodeChange>>(window, "codeChanges").Add(change); for (int i = 0; i < 90; i++) Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "line " + i }); Call(window, "RefreshTranscriptWindow", 80); Call(window, "ShowCodeChanges", change); Assert.IsFalse(Get<bool>(window, "followConversation")); Assert.AreEqual(0, Get<int>(window, "firstLoadedEntry"));
                Call(window, "ShowCodeChanges", new object[] { null }); Call(window, "ShowCodeChanges", new object[] { null }); var item = Get<Button>(window, "changes").ContextMenuStrip.Items[0]; item.PerformClick(); Call(window, "NewSession", (object)null); item.PerformClick(); Assert.AreEqual(0, Get<List<CodeChange>>(window, "codeChanges").Count);
                ChatWindow.WriteSettings = s => { throw new InvalidOperationException("reset persistence failure"); }; Call(window, "ResetProviderConnection");
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeHttpStreamingAndCancellationCallbacksUseOnlyMemoryTransport()
        {
            foreach (bool stop in new[] { false, true })
                using (var runtime = new RuntimeScope())
                using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
                {
                    var handler = new RuntimeHttpHandler { Streaming = true, Body = "data: {\"choices\":[{\"delta\":{\"content\":\"streamed\"},\"finish_reason\":\"stop\"}]}\n\ndata: [DONE]\n\n" }; window.CodexInterruptOverride = () => Task.CompletedTask; handler.BeforeResponse = () => { if (stop) CompleteOnSta((Task)Call(window, "StopTurnAsync")); }; window.HttpHandlerOverride = () => handler; Question(window, "request"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsFalse(Get<bool>(window, "busy")); Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(x => x.Text.Contains(stop ? UiText.Get("Response interrupted") : "streamed"))); Set(window, "currentSession", null);
                }
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                var handler = new RuntimeHttpHandler(); handler.BeforeResponse = () => { var client = Get<LlmChatClient>(window, "activeHttpClient"); Assert.AreEqual("tool result", client.ToolHandler("status", "{}").GetAwaiter().GetResult()); window.CodexInterruptOverride = () => Task.CompletedTask; CompleteOnSta((Task)Call(window, "StopTurnAsync")); Assert.ThrowsException<OperationCanceledException>(() => client.ToolHandler("status", "{}").GetAwaiter().GetResult()); client.TextDelta("late fragment"); }; ChatWindow.InvokeTool = (t, n, a) => Task.FromResult("tool result"); window.HttpHandlerOverride = () => handler; Question(window, "request"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsTrue(Get<bool>(window, "stopRequested")); Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeLateFailuresDisposedCallbacksAndPartialCleanupRemainSafe()
        {
            using (var runtime = new RuntimeScope())
            {
                using (var window = new ChatWindow(runtime.Session))
                {
                    CompleteOnSta((Task)Call(window, "LoadModelsAsync")); var effort = Get<ComboBox>(window, "effortPicker"); effort.Items.Add(new LlmEffortOption("high", "High")); ChatWindow.WriteSettings = s => { throw new IOException("effort persistence unavailable"); }; effort.SelectedIndex = effort.Items.Count - 1; ChatWindow.WriteSettings = s => runtime.Saves++;
                    Call(window, "ResetProviderConnection"); effort.SelectedIndex = -1; Question(window, "reconnect without effort"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.AreEqual("thread", Get<ChatSessionState>(window, "currentSession").CodexThreadId);
                    var context = SynchronizationContext.Current; CodexAppServerClient fallback; try { SynchronizationContext.SetSynchronizationContext(null); fallback = (CodexAppServerClient)Call(window, "CreateCodexClient"); } finally { SynchronizationContext.SetSynchronizationContext(context); }
                    fallback.Dispose();
                    var client = Get<CodexAppServerClient>(window, "codex"); var chat = (Action<string, string, string, bool>)typeof(CodexAppServerClient).GetField("ChatUpdate", Fields).GetValue(client); var ready = (Action<string>)typeof(CodexAppServerClient).GetField("ThreadReady", Fields).GetValue(client);
                    var pending = new TaskCompletionSource<LlmModelOption[]>(); window.ModelCatalogueOverride = p => pending.Task; var load = (Task)Call(window, "LoadModelsAsync"); window.Dispose(); pending.SetException(new InvalidOperationException("late disposed catalogue failure")); CompleteOnSta(load); var count = Get<List<ChatEntry>>(window, "transcriptEntries").Count; chat("final", "late", "ignored", true); ready("ignored"); Assert.AreEqual(count, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                }
                using (var design = new ChatWindow())
                {
                    var changes = Get<Button>(design, "changes"); try { Set(design, "changes", null); Call(design, "DisposeRuntime"); Assert.IsNull(Get<CodexAppServerClient>(design, "codex")); } finally { Set(design, "changes", changes); }
                    Call(design, "SetBusy", true); CompleteOnSta((Task)Call(design, "StopTurnAsync")); Assert.IsTrue(Get<bool>(design, "stopRequested")); Call(design, "SetBusy", false);
                }
                runtime.Host = r => Response.Success(new object[0]); using (var window = new ChatWindow(runtime.Session))
                {
                    var unavailable = (LlmProvider)Activator.CreateInstance(typeof(LlmProvider), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "Future provider", false, false, null, null, null }, null); var providers = Get<ComboBox>(window, "providerPicker"); providers.Items.Add(unavailable); providers.SelectedItem = unavailable; Call(window, "ResetProviderConnection"); StringAssert.Contains(Get<Label>(window, "status").Text, UiText.Get("not implemented yet"));
                    var pending = new TaskCompletionSource<LlmModelOption[]>(); providers.SelectedIndex = 0; window.ModelCatalogueOverride = p => pending.Task; var stale = (Task)Call(window, "LoadModelsAsync"); window.ModelCatalogueOverride = p => Task.FromResult(new LlmModelOption[0]); CompleteOnSta((Task)Call(window, "LoadModelsAsync")); pending.SetException(new InvalidOperationException("stale codex failure")); CompleteOnSta(stale);
                    Click(Get<Button>(window, "send")); Call(window, "SetBusy", true); window.CodexInterruptOverride = () => Task.CompletedTask; Set(window, "stopRequested", false); Click(Get<Button>(window, "send")); Assert.IsTrue(Get<bool>(window, "stopRequested")); Call(window, "SetBusy", false);
                }
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeHttpStreamErrorsAndConcurrentToolStopRetainUndoHistory()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                window.HttpHandlerOverride = () => new RuntimeHttpHandler { Streaming = true, Body = "data: {\"choices\":[{\"delta\":{\"content\":\"partial\"}}]}\n\ndata: {\"error\":{\"message\":\"failed\"}}\n\n" }; Question(window, "request"); CompleteOnSta((Task)Call(window, "SendAsync")); var live = Get<Dictionary<string, ChatEntry>>(window, "liveEntries"); Assert.AreEqual(1, live.Count); Assert.IsTrue(Get<HashSet<string>>(window, "completedStreams").Contains(live.Keys.Single())); Assert.AreEqual("partial", live.Values.Single().Text);
                string call = "{\"id\":\"one\",\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}"; window.CodexInterruptOverride = () => Task.CompletedTask; window.HttpHandlerOverride = () => new RuntimeHttpHandler { Body = "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"tool_calls\":[" + call + "," + call.Replace("one", "two") + "]}}]}" }; ChatWindow.InvokeTool = (t, n, a) => { CompleteOnSta((Task)Call(window, "StopTurnAsync")); return Task.FromResult("stopped after first tool"); }; Question(window, "tools"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsTrue(Get<bool>(window, "stopRequested"));
                window.HttpHandlerOverride = () => { CompleteOnSta((Task)Call(window, "StopTurnAsync")); return new RuntimeHttpHandler(); }; Question(window, "stop before HTTP loop"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsFalse(Get<bool>(window, "busy")); Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeDisposalDuringHttpAndDeferredVerificationCompleteWithoutExternalEffects()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                var handler = new RuntimeHttpHandler(); handler.BeforeResponse = () => { var client = Get<LlmChatClient>(window, "activeHttpClient"); window.Dispose(); client.TextDelta("queued late fragment"); }; window.HttpHandlerOverride = () => handler; Question(window, "dispose during request"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsTrue(window.IsDisposed); Set(window, "currentSession", null);
            }
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                CompleteOnSta((Task)Call(window, "LoadModelsAsync")); var pending = new TaskCompletionSource<string>(); bool verifying = false; ChatWindow.InvokeTool = (t, n, a) => { verifying = true; return pending.Task; }; Get<CheckBox>(window, "verifyAfterEdit").Checked = true; runtime.Transport.BeforeComplete = () => { var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "before", "new", "after", 1); var tools = Get<LlmVbeTools>(window, "tools"); ((Action<CodeChange>)typeof(LlmVbeTools).GetField("CodeEdited", Fields).GetValue(tools))(change); }; Question(window, "edit and verify"); var send = (Task)Call(window, "SendAsync"); var deadline = DateTime.UtcNow.AddSeconds(5); while (!verifying && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
                Assert.IsTrue(verifying); Assert.IsFalse(send.IsCompleted); pending.SetResult(new JavaScriptSerializer().Serialize(Response.Success(new { Compiled = true }))); CompleteOnSta(send); Assert.IsFalse(Get<bool>(window, "busy")); Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(x => x.Speaker == "Vérification"));
            }
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                var handler = new RuntimeHttpHandler(); handler.BeforeResponse = () => CompleteOnSta((Task)Call(window, "StopTurnAsync")); window.HttpHandlerOverride = () => handler; Question(window, "native HTTP cancellation"); CompleteOnSta((Task)Call(window, "SendAsync")); Assert.IsTrue(Get<bool>(window, "stopRequested")); Assert.IsFalse(Get<bool>(window, "busy")); Set(window, "currentSession", null);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void RuntimeUnknownProviderInvalidNativeEndpointAndSubscriberFailureAreSafe()
        {
            using (var runtime = new RuntimeScope()) { runtime.Settings.ProviderName = "Unknown provider"; using (var window = new ChatWindow(runtime.Session)) { Assert.AreEqual("Codex", ((LlmProvider)Get<ComboBox>(window, "providerPicker").SelectedItem).Name); } }
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                Get<LlmSettings>(window, "settings").OllamaEndpoint = "file:///invalid-chat-endpoint"; window.HttpHandlerOverride = null; Get<CheckBox>(window,"verifyAfterEdit").Checked=false; Question(window, "invalid endpoint"); CompleteOnSta((Task)Call(window, "SendAsync")); StringAssert.Contains(Get<List<ChatEntry>>(window, "transcriptEntries").Last().Text, "HTTPS"); Set(window, "currentSession", null);
            }
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                window.HttpHandlerOverride = () => new ChatResponseHandler(); Question(window, "request"); var prompt = Get<System.Windows.Controls.TextBox>(window, "prompt"); System.Windows.Controls.TextChangedEventHandler failedView = (s, e) => { if (prompt.Text == "request" && Get<bool>(window, "busy")) throw new IOException("view subscriber failed while restoring draft"); }; prompt.TextChanged += failedView;
                try { Assert.ThrowsException<IOException>(() => CompleteOnSta((Task)Call(window, "SendAsync"))); Assert.IsFalse(Get<bool>(window, "busy")); Assert.AreEqual("request", Get<ChatSessionState>(window, "currentSession").Draft); }
                finally { prompt.TextChanged -= failedView; Set(window, "currentSession", null); }
            }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void AssistantIdentityIsEmbeddedAndPartialCleanupToleratesAMissingAboutControl()
        {
            using (var resource = typeof(ChatWindow).Assembly.GetManifestResourceStream("CodexVBE.Icons.assistant.ico"))
            {
                Assert.IsNotNull(resource);
                Assert.IsTrue(resource.Length > 0);
            }
            using (var window = new ChatWindow())
            {
                var about = Get<ToolStripMenuItem>(window, "about");
                Assert.IsNotNull(about.Image);
                Assert.IsTrue(about.Image.Width > 0);
                try
                {
                    Set(window, "about", null);
                    Call(window, "DisposeRuntime");
                    Assert.IsNotNull(about.Image);
                }
                finally { Set(window, "about", about); }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void NativeActivitiesAreDeliveredOnlyToTheOwningLiveConversation()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyCodexWindow(new ChatSessionState()))
            using (var client = (CodexAppServerClient)Call(window, "CreateCodexClient"))
            {
                var owner = Get<ChatSessionState>(window, "currentSession");
                var update = (Action<CodexAgentActivity>)typeof(CodexAppServerClient).GetField("ActivityUpdate", Fields).GetValue(client);
                update(new CodexAgentActivity { Id = "owned", Kind = "dynamicToolCall", Title = "read_module", Status = "inProgress" });
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("owned", entries[0].Activity.Id);
                Set(window, "currentSession", new ChatSessionState());
                update(new CodexAgentActivity { Id = "foreign", Kind = "dynamicToolCall", Status = "completed" });
                Assert.AreEqual(1, entries.Count);
                Set(window, "currentSession", owner);
                window.Dispose();
                update(new CodexAgentActivity { Id = "late", Kind = "dynamicToolCall", Status = "completed" });
                Assert.AreEqual(1, entries.Count);
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void EmptySendResumesThePausedHttpBudgetWithoutCreatingANewQuestion()
        {
            var state = new ChatSessionState { BudgetPaused = true, PausedProvider = "Ollama", PausedModel = "local-test", PausedMode = ChatMode.Agent, Mode = ChatMode.Agent, PausedTurnId = "paused" };
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(state))
            {
                var handler = new ChatResponseHandler("{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"resumed answer\"}}]}");
                window.HttpHandlerOverride = () => handler;
                CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(1, handler.Requests.Count);
                Assert.IsFalse(state.BudgetPaused);
                Assert.IsFalse(Get<bool>(window, "busy"));
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.IsFalse(entries.Exists(entry => entry.Speaker == "Vous"));
                Assert.IsTrue(entries.Exists(entry => entry.Text == "resumed answer"));
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void SendRepropagatesARecoveryFailureAfterTheWindowClosesDuringTheProviderReply()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState()))
            {
                var messages = Get<List<object>>(window, "messages");
                window.HttpHandlerOverride = () => new RuntimeHttpHandler { BeforeResponse = () => {
                    window.Dispose();
                    messages.Add(new { role = "assistant", content = new string('x', 10 * 1024 * 1024 + 1) });
                    throw new System.IO.IOException("reply interrupted after closure");
                } };
                Question(window, "continue");
                var task = (Task)Call(window, "SendAsync");
                var error = Assert.ThrowsException<InvalidOperationException>(() => CompleteOnSta(task));
                StringAssert.Contains(error.StackTrace, "CompletePendingToolResponses");
                Assert.IsTrue(task.IsFaulted);
                Assert.IsTrue(window.IsDisposed);
                Assert.IsNull(Get<LlmChatClient>(window, "activeHttpClient"));
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void SendFinalizationReleasesTheTurnWhenTheSessionDisappearsAtTheProviderFailureBoundary()
        {
            foreach (string state in new[] { "absent", "paused", "unpaused" })
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState()))
            {
                var session = Get<ChatSessionState>(window, "currentSession");
                int replies = 0, compilations = 0;
                Get<CheckBox>(window, "verifyAfterEdit").Checked = true;
                ChatWindow.InvokeTool = (tools, name, arguments) => { compilations++; throw new AssertFailedException("A turn without applied changes must not compile."); };
                window.HttpHandlerOverride = () => new RuntimeHttpHandler { BeforeResponse = () => {
                    replies++;
                    if (state == "absent") Set(window, "currentSession", null);
                    else session.BudgetPaused = state == "paused";
                    throw new System.IO.IOException("owned provider reply failed");
                } };
                Question(window, "owned retry request");
                CompleteOnSta((Task)Call(window, "SendAsync"));
                Assert.AreEqual(1, replies); Assert.AreEqual(0, compilations);
                Assert.IsFalse(window.IsDisposed); Assert.IsFalse(Get<bool>(window, "busy"));
                Assert.IsNull(Get<LlmChatClient>(window, "activeHttpClient")); Assert.IsNull(Get<string>(window, "activeTurnId"));
                Assert.AreEqual("owned retry request", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Exists(entry => entry.Speaker == "Erreur" && entry.Text == "owned provider reply failed"));
                if (state == "absent") Assert.IsNull(Get<ChatSessionState>(window, "currentSession"));
                else { Assert.AreSame(session, Get<ChatSessionState>(window, "currentSession")); Assert.AreEqual(state == "paused", session.BudgetPaused); }
                Set(window, "currentSession", null);
            }
        }
    }
}
