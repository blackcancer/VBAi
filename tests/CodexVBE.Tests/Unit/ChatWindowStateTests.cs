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

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class ChatWindowStateTests
    {
        private const BindingFlags Fields = BindingFlags.Instance | BindingFlags.NonPublic;
        private const BindingFlags Methods = BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic;

        private static T Get<T>(ChatWindow window, string field)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            return (T)info.GetValue(window);
        }

        private static void Set(ChatWindow window, string field, object value)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            info.SetValue(window, value);
        }

        private static object Call(ChatWindow window, string method, params object[] args)
        {
            var info = typeof(ChatWindow).GetMethod(method, Methods);
            Assert.IsNotNull(info, "Missing ChatWindow method " + method);
            return info.Invoke(window, args);
        }

        private static ChatWindow Surfaces()
        {
            var window = new ChatWindow();
            Call(window, "InitializeShell");
            Call(window, "InitializeComposer", new object[] { null });
            Call(window, "InitializeTranscript");
            return window;
        }

        private static ChatWindow ReadyCodexWindow(ChatSessionState session)
        {
            var window = Surfaces();
            var settings = new LlmSettings();
            Set(window, "settings", settings);
            Set(window, "tools", new LlmVbeTools(null, window, settings));
            Set(window, "currentSession", session);
            var providers = Get<ComboBox>(window, "providerPicker");
            providers.Items.Add(LlmProvider.All[0]);
            providers.SelectedIndex = 0;
            var models = Get<ComboBox>(window, "modelPicker");
            models.Items.Add(new LlmModelOption("gpt-test", "Test", true, null, new LlmEffortOption[0]));
            models.SelectedIndex = 0;
            return window;
        }

        private static ChatWindow ReadyHttpWindow(ChatSessionState session)
        {
            var window = Surfaces();
            var settings = new LlmSettings();
            Set(window, "settings", settings);
            Set(window, "tools", new LlmVbeTools(new VbeSession(new object()), window, settings));
            Set(window, "currentSession", session);
            var providers = Get<ComboBox>(window, "providerPicker");
            providers.Items.Add(LlmProvider.All[2]);
            providers.SelectedIndex = 0;
            var models = Get<ComboBox>(window, "modelPicker");
            models.Items.Add(new LlmModelOption("local-test", "Local test"));
            models.SelectedIndex = 0;
            return window;
        }

        private sealed class ChatResponseHandler : HttpMessageHandler
        {
            private readonly Queue<string> responses = new Queue<string>();
            public readonly List<string> Requests = new List<string>();
            public ChatResponseHandler(params string[] bodies) { foreach (string body in bodies) responses.Enqueue(body); }
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                Requests.Add(await request.Content.ReadAsStringAsync());
                if (responses.Count == 0) throw new InvalidOperationException("Unexpected provider request.");
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(responses.Dequeue()) };
            }
        }

        private static void Question(ChatWindow window, string text)
        {
            var prompt = Get<object>(window, "prompt");
            prompt.GetType().GetProperty("Text").SetValue(prompt, text, null);
        }

        private static object AddScope(ChatWindow window, string key)
        {
            var type = typeof(ChatWindow).GetNestedType("MacroScope", BindingFlags.NonPublic);
            var scope = Activator.CreateInstance(type, true);
            type.GetField("Key").SetValue(scope, key);
            type.GetField("Project").SetValue(scope, key);
            type.GetField("Name").SetValue(scope, key);
            type.GetField("Label").SetValue(scope, key);
            Get<ComboBox>(window, "scopePicker").Items.Add(scope);
            return scope;
        }

        private static void CompleteOnSta(Task task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!task.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(1);
            }
            Assert.IsTrue(task.IsCompleted, "The chat operation did not complete on the STA thread.");
            task.GetAwaiter().GetResult();
        }

        [TestMethod]
        [STATestMethod]
        public void ConstructorAndLocalSurfacesCreateComposerTranscriptAndModes()
        {
            using (var window = Surfaces())
            {
                Assert.AreEqual(3, Get<ComboBox>(window, "modePicker").Items.Count);
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
                Assert.IsNotNull(Get<object>(window, "prompt"));
                Assert.IsNotNull(Get<object>(window, "conversationScroll"));
                Assert.IsFalse(Get<Panel>(window, "historyPanel").Visible);
                Assert.IsFalse(Get<ProgressBar>(window, "activityBar").Visible);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ProviderEffortChoicesFollowModelAndBusyState()
        {
            using (var window = Surfaces())
            {
                Set(window, "currentSession", new ChatSessionState { Effort = "high" });
                var codexModel = new LlmModelOption("gpt-test", "Test", false, "medium", new[] {
                    new LlmEffortOption("medium", "Medium"), new LlmEffortOption("high", "High") });
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
        public void HistoryFiltersArchivedSessionsAndOrdersPinnedFirst()
        {
            using (var window = Surfaces())
            {
                var sessions = Get<List<ChatSessionState>>(window, "scopeSessions");
                var normal = new ChatSessionState { Title = "Alpha" };
                var pinned = new ChatSessionState { Title = "Alpha pinned", Pinned = true };
                var archived = new ChatSessionState { Title = "Alpha archived", Archived = true };
                sessions.Add(normal); sessions.Add(pinned); sessions.Add(archived);
                Get<TextBox>(window, "historySearch").Text = "Alpha";
                Call(window, "RefreshHistory");
                var list = Get<ListBox>(window, "sessionList");
                Assert.AreEqual(2, list.Items.Count);
                Assert.AreSame(pinned, list.Items[0]);
                Get<CheckBox>(window, "showArchived").Checked = true;
                Call(window, "RefreshHistory");
                Assert.AreEqual(3, list.Items.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void SessionTitleAndDraftAreSavedWithoutExternalStore()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState { Scope = "temporary:test" };
                Set(window, "currentSession", session);
                Call(window, "RenameFromQuestion", "First line\nSecond line");
                Assert.AreEqual("First line Second line", session.Title);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                var prompt = Get<object>(window, "prompt");
                prompt.GetType().GetProperty("Text").SetValue(prompt, "unsent draft", null);
                Get<List<object>>(window, "messages").Add(new Dictionary<string, object> {
                    ["role"] = "user", ["content"] = "hello" });
                Call(window, "SaveCurrentSession");
                Assert.AreEqual("unsent draft", session.Draft);
                StringAssert.Contains(session.MessagesJson, "hello");
                Assert.AreEqual(0, session.DraftAttachments.Length);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ManualRenameTrimsAndCapsTitle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState { Scope = "temporary:test" };
                Set(window, "currentSession", session);
                Get<TextBox>(window, "chatTitleEditor").Text = "  " + new string('a', 130) + "  ";
                Call(window, "RenameCurrentChat");
                Assert.AreEqual(120, session.Title.Length);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void InterruptedToolHistoryDropsOnlyUnfinishedAssistantTail()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "system", ["content"] = "rules" });
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] =
                    new object[] { new Dictionary<string, object> { ["id"] = "pending" } } });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
                var repaired = new JavaScriptSerializer().Serialize(messages[2]);
                StringAssert.Contains(repaired, "assistant");
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CompletedToolHistoryIsPreserved()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] =
                    new object[] { new Dictionary<string, object> { ["id"] = "done" } } });
                messages.Add(new Dictionary<string, object> { ["role"] = "tool", ["tool_call_id"] = "done" });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void TranscriptStreamsUpdateInPlaceAndAvoidDuplicateFinalMessage()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "empty", " ", false);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(0, entries.Count);
                Call(window, "ReceiveChatUpdate", "final", "answer", "Hel", false);
                Call(window, "ReceiveChatUpdate", "final", "answer", "lo", true);
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("lo", entries[0].Text);
                Call(window, "CompleteAssistantResponse", "lo");
                Assert.AreEqual(1, entries.Count);
                Call(window, "CompleteAssistantResponse", "different");
                Assert.AreEqual(2, entries.Count);
                Call(window, "ClearTranscript");
                Assert.AreEqual(0, entries.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ComposerTokenBoundaryAndDraftContextLimitAreEnforced()
        {
            using (var window = Surfaces())
            {
                var token = "#Project.Module";
                Assert.AreEqual(true, Call(window, "ContainsToken", "Use " + token + " now", token));
                Assert.AreEqual(false, Call(window, "ContainsToken", "x" + token + "suffix", token));
                Assert.AreEqual(true, Call(window, "IsReferenceChar", '_'));
                Assert.AreEqual(false, Call(window, "IsReferenceChar", '-'));
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "selection", Text = "code" });
                var prepared = (ChatAttachment[])Call(window, "PrepareAttachments", "question");
                Assert.AreEqual(1, prepared.Length);
                Assert.AreEqual("selection", prepared[0].Label);
                attachments[0].Text = new string('x', 48001);
                var error = Assert.ThrowsException<TargetInvocationException>(() =>
                    Call(window, "PrepareAttachments", "question"));
                Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            }
        }

        [TestMethod]
        [STATestMethod]
        public void EditorActionSeedsCommandWhileBusyStateBlocksIt()
        {
            using (var window = Surfaces())
            {
                Call(window, "SetBusy", true);
                window.PrepareEditorAction("/corriger");
                var prompt = Get<object>(window, "prompt");
                Assert.AreEqual("", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Call(window, "SetBusy", false);
                window.PrepareEditorAction("/corriger");
                Assert.AreEqual("/corriger ", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CurrentReferencesKeepLatestExactTokenAndIgnoreEmbeddedMatches()
        {
            using (var window = Surfaces())
            {
                var selected = Get<List<VbeChatReference>>(window, "selectedReferences");
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "old" });
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "new" });
                var exact = (VbeChatReference[])Call(window, "CurrentReferences", "Use #P.M for this change");
                Assert.AreEqual(1, exact.Length);
                Assert.AreEqual("new", exact[0].Sha256);
                var embedded = (VbeChatReference[])Call(window, "CurrentReferences", "prefix#P.Msuffix");
                Assert.AreEqual(0, embedded.Length);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ContextChipsRepresentMemoryAndDraftAttachmentsAndCanRemoveDraft()
        {
            using (var window = Surfaces())
            {
                Set(window, "projectMemory", "local note");
                Get<CheckBox>(window, "attachMemory").Checked = true;
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "Selected code", Text = "Sub A()" });
                Call(window, "RefreshContextChips");
                var chips = Get<FlowLayoutPanel>(window, "contextChips");
                Assert.AreEqual(2, chips.Controls.Count);
                typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(chips.Controls[1], new object[] { EventArgs.Empty });
                Assert.AreEqual(0, attachments.Count);
                Assert.AreEqual(1, chips.Controls.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModeSelectionUpdatesSessionOnlyWhenIdle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState { Mode = ChatMode.Agent };
                Set(window, "currentSession", session);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Plan;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Call(window, "SetBusy", true);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Discussion;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void MarkdownTranscriptAndContextPreviewRetainStructuredEntries()
        {
            using (var window = Surfaces())
            {
                Call(window, "AddTranscriptMessage", "Assistant", "# Heading\n- item\n```vba\nDebug.Print 1\n```");
                Assert.AreEqual(1, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Assert.AreEqual(1, ((IDictionary)Get<object>(window, "entryViews")).Count);
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(
                    new ChatAttachment { Label = "Selection", Text = "VBA code" });
                Call(window, "RefreshContextPreview");
                Assert.AreEqual(2, Get<FlowLayoutPanel>(window, "contextPreview").Controls.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CodexSendRecordsQuestionAnswerAndStructuredRequestWithoutNetwork()
        {
            var session = new ChatSessionState { Scope = "temporary:test", ResumeContext = "Earlier branch" };
            using (var window = ReadyCodexWindow(session))
            {
                string request = null;
                window.CodexTurnOverride = (text, model, effort) => {
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
            var session = new ChatSessionState { Scope = "temporary:test" };
            using (var window = ReadyCodexWindow(session))
            {
                var drafts = Get<List<ChatAttachment>>(window, "draftAttachments");
                drafts.Add(new ChatAttachment { Label = "Extrait", Text = "Sub Écrire()" });
                window.CodexTurnOverride = (text, model, effort) =>
                    Task.FromException<string>(new InvalidOperationException("provider failed"));
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
            var session = new ChatSessionState { Scope = "temporary:test" };
            using (var window = ReadyCodexWindow(session))
            {
                var pending = new TaskCompletionSource<string>();
                window.CodexTurnOverride = (text, model, effort) => pending.Task;
                window.CodexInterruptOverride = () => {
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
                window.CodexTurnOverride = (text, model, effort) => {
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
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(
                    new ChatAttachment { Label = "Très long", Text = new string('x', 48001) });
                ((Task)Call(window, "SendAsync")).GetAwaiter().GetResult();
                Assert.IsFalse(invoked);
                Assert.AreEqual("Question", Get<object>(window, "prompt").GetType().GetProperty("Text").GetValue(Get<object>(window, "prompt"), null));
                Assert.IsFalse(Get<bool>(window, "busy"));
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CurrentSessionPersistsDraftMessagesAndEntriesInIsolatedStore()
        {
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-Chat-" + Guid.NewGuid().ToString("N"), "chat.db");
            try
            {
                using (var store = new ChatSessionStore(path))
                using (var window = Surfaces())
                {
                    var session = new ChatSessionState { Scope = "temporary:test", Title = "Essai éè" };
                    Set(window, "sessionStore", store);
                    Set(window, "currentSession", session);
                    Question(window, "Brouillon éè");
                    Get<List<object>>(window, "messages").Add(new Dictionary<string, object> {
                        ["role"] = "user", ["content"] = "Question éè" });
                    Call(window, "AddTranscriptMessage", "Assistant", "Réponse éè");
                    Call(window, "SaveCurrentSession");
                    var restored = store.List(session.Scope);
                    Assert.AreEqual(1, restored.Count);
                    Assert.AreEqual("Essai éè", restored[0].Title);
                    Assert.AreEqual("Brouillon éè", restored[0].Draft);
                    StringAssert.Contains(restored[0].MessagesJson, "Question éè");
                    Assert.AreEqual("Réponse éè", restored[0].Entries[0].Text);
                    Set(window, "currentSession", null);
                    Set(window, "sessionStore", null);
                }
            }
            finally
            {
                string directory = Path.GetDirectoryName(path);
                if (Directory.Exists(directory))
                {
                    foreach (string file in Directory.GetFiles(directory)) File.Delete(file);
                    Directory.Delete(directory, false);
                }
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
                window.ModelCatalogueOverride = provider => ++calls == 1 ? first.Task : Task.FromResult(new[] {
                    new LlmModelOption("other", "Other"), new LlmModelOption("saved", "Saved") });
                var stale = (Task)Call(window, "LoadModelsAsync");
                CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                first.SetResult(new[] { new LlmModelOption("stale", "Stale") });
                CompleteOnSta(stale);
                var picker = Get<ComboBox>(window, "modelPicker");
                Assert.AreEqual(2, picker.Items.Count);
                Assert.AreEqual("saved", ((LlmModelOption)picker.SelectedItem).Id);
                Assert.IsTrue(Get<Button>(window, "refreshModels").Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModelCatalogueFailureLeavesRefreshAvailable()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test" }))
            {
                window.ModelCatalogueOverride = provider => Task.FromException<LlmModelOption[]>(
                    new InvalidOperationException("catalogue failed"));
                CompleteOnSta((Task)Call(window, "LoadModelsAsync"));
                Assert.AreEqual(0, Get<ComboBox>(window, "modelPicker").Items.Count);
                Assert.IsTrue(Get<Button>(window, "refreshModels").Enabled);
                Assert.IsFalse(Get<ComboBox>(window, "modelPicker").Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void HttpProviderCompletesToolCallAndAnswerWithoutNetwork()
        {
            var session = new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" };
            using (var window = ReadyHttpWindow(session))
            {
                var handler = new ChatResponseHandler(
                    "{\"choices\":[{\"finish_reason\":\"tool_calls\",\"message\":{\"role\":\"assistant\",\"content\":null,\"tool_calls\":[{\"id\":\"call-1\",\"type\":\"function\",\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}]}}]}",
                    "{\"choices\":[{\"finish_reason\":\"stop\",\"message\":{\"role\":\"assistant\",\"content\":\"Terminé\"}}]}");
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
            var session = new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" };
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

        [TestMethod]
        [STATestMethod]
        public void ScopeSwitchLoadsSavedConversationAndMemoryThenRestoresCachedScope()
        {
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-Scope-" + Guid.NewGuid().ToString("N"), "chat.db");
            try
            {
                using (var store = new ChatSessionStore(path))
                using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:A", Title = "A" }))
                {
                    window.ModelCatalogueOverride = provider => Task.FromResult(new LlmModelOption[0]);
                    var original = Get<ChatSessionState>(window, "currentSession");
                    Get<List<ChatSessionState>>(window, "scopeSessions").Add(original);
                    var scopes = Get<ComboBox>(window, "scopePicker");
                    AddScope(window, "temporary:A");
                    AddScope(window, "temporary:B");
                    store.Save(new ChatSessionState { Scope = "temporary:B", Title = "B", Draft = "Brouillon B" });
                    store.SaveMemory("temporary:B", "Mémoire B");
                    Set(window, "sessionStore", store);
                    scopes.SelectedIndex = 1;
                    Call(window, "ChangeScope");
                    Assert.AreEqual("B", Get<ChatSessionState>(window, "currentSession").Title);
                    Assert.AreEqual("Brouillon B", Get<object>(window, "prompt").GetType().GetProperty("Text").GetValue(Get<object>(window, "prompt"), null));
                    Assert.AreEqual("Mémoire B", Get<TextBox>(window, "memoryEditor").Text);
                    scopes.SelectedIndex = 0;
                    Call(window, "ChangeScope");
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                    Set(window, "currentSession", null);
                    Set(window, "sessionStore", null);
                }
            }
            finally
            {
                string directory = Path.GetDirectoryName(path);
                if (Directory.Exists(directory))
                {
                    foreach (string file in Directory.GetFiles(directory)) File.Delete(file);
                    Directory.Delete(directory, false);
                }
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ForkKeepsOnlySelectedConversationPrefixAndBuildsResumeContext()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test", Title = "Original" }))
            {
                window.ModelCatalogueOverride = provider => Task.FromResult(new LlmModelOption[0]);
                var original = Get<ChatSessionState>(window, "currentSession");
                Get<List<ChatSessionState>>(window, "scopeSessions").Add(original);
                var first = new ChatEntry { Speaker = "Vous", Text = "Question initiale" };
                var reply = new ChatEntry { Speaker = "Assistant", Text = "Réponse initiale" };
                Call(window, "AddEntry", first);
                Call(window, "AddEntry", reply);
                Call(window, "AddEntry", new ChatEntry { Speaker = "Vous", Text = "Suite exclue" });
                Call(window, "ForkChat", reply);
                var fork = Get<ChatSessionState>(window, "currentSession");
                Assert.AreNotSame(original, fork);
                Assert.AreEqual(2, fork.Entries.Count);
                StringAssert.Contains(fork.MessagesJson, "Question initiale");
                Assert.IsFalse(fork.MessagesJson.Contains("Suite exclue"));
                StringAssert.Contains(fork.ResumeContext, "Réponse initiale");
                Assert.AreEqual(3, original.Entries.Count);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void StreamingSummaryAndToolEntriesCompleteInPlace()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "Réflexion ", false);
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "terminée", true);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", "Lecture", false);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", null, true);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(2, entries.Count);
                Assert.AreEqual("Réflexion", entries[0].Speaker);
                Assert.AreEqual("terminée", entries[0].Text);
                Assert.AreEqual("Outil", entries[1].Speaker);
                Assert.AreEqual("Lecture", entries[1].Text);
                Assert.AreEqual(2, Get<HashSet<string>>(window, "completedStreams").Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void WorkflowRejectsStaleModuleAttachmentBeforeProviderCall()
        {
            var host = new VbeSessionTests.FakeVbe();
            var project = new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("Sub Test()\r\nEnd Sub") });
            host.VBProjects.Add(project);
            using (var window = Surfaces())
            {
                Set(window, "scopeSession", new VbeSession(host));
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment {
                    Label = "Sélection", Text = "Sub Test()", Project = "P", Module = "Module1", Sha256 = "stale" });
                var error = Assert.ThrowsException<TargetInvocationException>(() =>
                    Call(window, "PrepareAttachments", "Question"));
                StringAssert.Contains(error.InnerException.Message, "Sélection");
            }
        }

        [TestMethod]
        [STATestMethod]
        public void VerificationReportsUnverifiedWhenNoProjectIsConnected()
        {
            using (var window = Surfaces())
            {
                CompleteOnSta((Task)Call(window, "VerifyProjectAsync"));
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("Vérification", entries[0].Speaker);
                Assert.IsFalse(string.IsNullOrWhiteSpace(entries[0].Text));
            }
        }
    }
}
