using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class ChatWindow : Form
    {
        private readonly LlmSettings settings;
        private readonly LlmVbeTools tools;
        private CodexAppServerClient codex;
        private readonly List<object> messages = new List<object>();
        private readonly List<CodeChange> codeChanges = new List<CodeChange>();
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private bool busy;
        private int catalogueVersion;
        private bool restoringSelection;
        private bool stopRequested;
        private LlmChatClient activeHttpClient;
        internal Func<string, string, string, Task<string>> CodexTurnOverride;
        internal Func<Task> CodexInterruptOverride;
        internal Func<LlmProvider, Task<LlmModelOption[]>> ModelCatalogueOverride;
        internal Func<HttpMessageHandler> HttpHandlerOverride;

        public ChatWindow()
        {
            InitializeComponent();
            Icon = VbeWindowIcons.Icon("assistant");
            github.Image = VbeWindowIcons.Image("github");
            configure.Image = VbeWindowIcons.Image("settings");
            UiText.Apply(this, components);
        }

        public ChatWindow(VbeSession session) : this()
        {
            InitializeShell();
            InitializeComposer(session);
            InitializeTranscript();
            foreach (var item in LlmProvider.All) providerPicker.Items.Add(item);

            try { settings = ReadSettings(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
            tools = new LlmVbeTools(session, this, settings);
            tools.ValidateScope = EnsureCurrentScope;
            tools.CodeEdited += change =>
            {
                change.TurnId = activeTurnId;
                codeChanges.Add(change);
                changes.Enabled = true;
                changes.Text = UiText.Get("Changes · ") + codeChanges.Count;
                AddCodeChangeCard(change);
                RefreshCodeChangeCards();
            };
            changes.Click += (sender, args) => ShowCodeChanges();
            providerPicker.SelectedIndexChanged += (sender, args) =>
            {
                if (loadingSession || busy) return;
                var provider = providerPicker.SelectedItem as LlmProvider;
                if (provider == null) return;
                settings.ProviderName = provider.Name;
                try { WriteSettings(settings); } catch (Exception ex) { LoadLog.Write("Provider selection save failed: " + ex.Message); }
                if (scopePicker.SelectedItem == null)
                {
                    ResetProviderConnection();
                    _ = LoadModelsAsync();
                }
                else NewSession(provider.Name);
            };
            configure.Click += (sender, args) => ShowSettings();
            refreshModels.Click += async (sender, args) => await LoadModelsAsync();
            modelPicker.SelectedIndexChanged += (sender, args) =>
            {
                var selectedModel = modelPicker.SelectedItem as LlmModelOption;
                var selectedProvider = providerPicker.SelectedItem as LlmProvider;
                if (selectedModel == null || selectedProvider == null) return;
                if (currentSession != null)
                {
                    if (!restoringSelection && currentSession.Model != selectedModel.Id) currentSession.Effort = null;
                    currentSession.Model = selectedModel.Id;
                }
                if (!restoringSelection) settings.SetSelectedModel(selectedProvider, selectedModel.Id);
                UpdateEfforts(selectedProvider, selectedModel);
                ScheduleSessionSave();
                if (!restoringSelection)
                    try { WriteSettings(settings); } catch (Exception ex) { LoadLog.Write("Model selection save failed: " + ex.Message); }
            };
            effortPicker.SelectedIndexChanged += (sender, args) =>
            {
                var selectedProvider = providerPicker.SelectedItem as LlmProvider;
                var selectedModel = modelPicker.SelectedItem as LlmModelOption;
                var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
                if (selectedProvider == null || selectedModel == null || selectedEffort == null) return;
                if (currentSession != null) currentSession.Effort = selectedEffort.Id;
                if (!restoringSelection) settings.SetReasoningEffort(selectedProvider, selectedModel.Id, selectedEffort.Id);
                ScheduleSessionSave();
                if (!restoringSelection)
                    try { WriteSettings(settings); } catch (Exception ex) { LoadLog.Write("Reasoning effort save failed: " + ex.Message); }
            };
            send.Click += async (sender, args) => { if (busy) await StopTurnAsync(); else await SendAsync(); };
            int defaultProvider = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            loadingSession = true;
            try { providerPicker.SelectedIndex = defaultProvider < 0 ? 0 : defaultProvider; }
            finally { loadingSession = false; }
            InitializeSessions(session);
            if (scopePicker.SelectedItem == null) _ = LoadModelsAsync();
        }

        private void UpdateEfforts(LlmProvider provider, LlmModelOption model)
        {
            effortPicker.Items.Clear();
            effortPicker.Enabled = false;
            if (!provider.IsCodex || model.Efforts.Length == 0) return;
            foreach (var item in model.Efforts) effortPicker.Items.Add(item);
            string selected = currentSession?.Effort ?? settings.GetReasoningEffort(provider, model.Id);
            int index = Array.FindIndex(model.Efforts, item => item.Id == selected);
            if (index < 0) index = Array.FindIndex(model.Efforts, item => item.Id == model.DefaultEffort);
            effortPicker.SelectedIndex = index < 0 ? 0 : index;
            effortPicker.Enabled = !busy;
        }

        private void ResetProviderConnection()
        {
            codex?.Dispose();
            codex = null;
            status.Text = ((LlmProvider)providerPicker.SelectedItem).Available ?
                UiText.Get("Loading provider models…") : UiText.Get("This provider is not implemented yet.");
            settings.ProviderName = ((LlmProvider)providerPicker.SelectedItem).Name;
            try { WriteSettings(settings); }
            catch (Exception saveError) { LoadLog.Write("LLM settings save failed: " + saveError.Message); }
        }

        private async Task LoadModelsAsync()
        {
            int version = ++catalogueVersion;
            var provider = providerPicker.SelectedItem as LlmProvider;
            modelPicker.Items.Clear();
            modelPicker.Enabled = false;
            effortPicker.Items.Clear();
            effortPicker.Enabled = false;
            refreshModels.Enabled = false;
            if (provider == null || !provider.Available) { refreshModels.Enabled = true; return; }
            try
            {
                LlmModelOption[] models;
                if (ModelCatalogueOverride != null)
                    models = await ModelCatalogueOverride(provider);
                else if (provider.IsCodex)
                {
                    if (codex == null)
                        codex = CreateCodexClient();
                    models = await codex.ListModelsAsync();
                }
                else models = await ReadModelCatalogue(provider, settings);
                if (IsDisposed || version != catalogueVersion || provider != providerPicker.SelectedItem) return;
                foreach (var item in models) modelPicker.Items.Add(item);
                LoadLog.Write("Chat model catalogue: " + provider.Name + " count=" + models.Length);
                string saved = currentSession?.Model ?? settings.GetSelectedModel(provider);
                int selected = Array.FindIndex(models, item => item.Id == saved);
                if (selected < 0) selected = Array.FindIndex(models, item => item.IsDefault);
                if (selected < 0 && models.Length > 0) selected = 0;
                if (selected >= 0)
                {
                    restoringSelection = true;
                    try { modelPicker.SelectedIndex = selected; }
                    finally { restoringSelection = false; }
                }
                SetStatus(models.Length == 0 ? UiText.Get("No models available for this provider.") : provider.Name + UiText.Get(" — ready"));
            }
            catch (Exception ex)
            {
                if (!IsDisposed && version == catalogueVersion)
                    SetStatus(UiText.Get("Model list unavailable: ") + ex.Message);
                LoadLog.Write("Chat model catalogue failed: " + provider.Name + " " + ex.ToString());
                if (provider.IsCodex && version == catalogueVersion) { codex?.Dispose(); codex = null; }
            }
            finally
            {
                if (!IsDisposed && version == catalogueVersion)
                {
                    modelPicker.Enabled = modelPicker.Items.Count > 0;
                    refreshModels.Enabled = true;
                }
            }
        }

        public void ShowSettings(IWin32Window owner = null)
        {
            if (busy)
            {
                ShowNotice(owner ?? this, UiText.Get("Wait for the response to finish or stop the agent before changing settings."),
                    "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string previousProvider = settings.ProviderName;
            string previousOpenAiEndpoint = settings.OpenAiEndpoint;
            string previousOllamaEndpoint = settings.OllamaEndpoint;
            string previousKey = settings.EncryptedOpenAiKey;
            string previousProviders = json.Serialize(new { settings.ProviderEndpoints, settings.EncryptedProviderKeys, settings.AzureUseEntraToken, settings.ManualModelLists });
            using (var dialog = new LlmSettingsWindow(settings))
            {
                if (ShowModal(dialog, owner ?? this) != DialogResult.OK) return;
                int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
                if (selected >= 0 && providerPicker.SelectedIndex != selected) providerPicker.SelectedIndex = selected;
                else
                {
                    bool connectionChanged = previousProvider != settings.ProviderName ||
                        previousOpenAiEndpoint != settings.OpenAiEndpoint ||
                        previousOllamaEndpoint != settings.OllamaEndpoint || previousKey != settings.EncryptedOpenAiKey ||
                        previousProviders != json.Serialize(new { settings.ProviderEndpoints, settings.EncryptedProviderKeys, settings.AzureUseEntraToken, settings.ManualModelLists });
                    if (connectionChanged) ResetProviderConnection();
                    _ = LoadModelsAsync();
                }
            }
        }

        private void Append(string speaker, string content)
        {
            if (IsDisposed) return;
            AddTranscriptMessage(speaker, content);
        }

        private void ShowCodeChanges(CodeChange selected = null)
        {
            if (selected == null)
            {
                var menu = new ContextMenuStrip();
                for (int i = codeChanges.Count - 1; i >= 0; i--)
                {
                    var change = codeChanges[i];
                    var item = new ToolStripMenuItem(change.Label);
                    item.Click += (sender, args) => ShowCodeChanges(change);
                    menu.Items.Add(item);
                }
                changes.ContextMenuStrip?.Dispose();
                changes.ContextMenuStrip = menu;
                menu.Show(changes, 0, changes.Height);
                return;
            }
            var entry = transcriptEntries.Find(x => x.Change == selected);
            if (entry != null)
            {
                followConversation = false;
                int index = transcriptEntries.IndexOf(entry);
                if (index < firstLoadedEntry) RefreshTranscriptWindow(index);
                conversationItems.ScrollIntoView(entry);
            }
        }

        private void SetStatus(string text)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetStatus(text))); return; }
            status.Text = storageFailed ? text + UiText.Get(" · History not saved") : text;
            toolTips.SetToolTip(status, status.Text);
        }

        private CodexAppServerClient CreateCodexClient()
        {
            var ownerSession = currentSession;
            var client = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                tools, SetStatus, settings, currentSession?.CodexThreadId, TransportFactory());
            client.ChatUpdate += (kind, id, text, complete) =>
            {
                if (currentSession == ownerSession && !IsDisposed) ReceiveChatUpdate(kind, id, text, complete);
            };
            client.ThreadReady += id =>
            {
                if (currentSession == ownerSession && currentSession != null && !IsDisposed)
                { currentSession.CodexThreadId = id; SaveCurrentSession(); }
            };
            return client;
        }

        private void SetBusy(bool value)
        {
            busy = value;
            modePicker.Enabled = !value;
            send.Text = value ? UiText.Get("Stop ■") : UiText.Get("Send ↑");
            toolTips.SetToolTip(send, value ? UiText.Get("Stop the current response. Changes already applied can still be undone in the chat.") : UiText.Get("Send the message and its context to the agent."));
            send.Enabled = true;
            newChat.Enabled = scopePicker.Enabled = sessionList.Enabled =
                providerPicker.Enabled = refreshModels.Enabled = configure.Enabled = !value;
            modelPicker.Enabled = !value && modelPicker.Items.Count > 0;
            effortPicker.Enabled = !value && effortPicker.Items.Count > 0;
            activityBar.Visible = value;
            RefreshCodeChangeCards();
        }

        private async Task StopTurnAsync()
        {
            if (!busy || stopRequested) return;
            stopRequested = true;
            send.Enabled = false;
            SetStatus(UiText.Get("Stopping…"));
            try
            {
                if (CodexInterruptOverride != null) await CodexInterruptOverride();
                else if (codex != null) await codex.InterruptAsync();
                else activeHttpClient?.Dispose();
            }
            catch (Exception ex) { SetStatus(UiText.Get("Unable to stop: ") + ex.Message); stopRequested = false; send.Enabled = true; }
        }

        private async Task SendAsync()
        {
            string question = prompt.Text.Trim();
            if (busy || question.Length == 0) return;
            var selectedModel = modelPicker.SelectedItem as LlmModelOption;
            if (selectedModel == null) { SetStatus(UiText.Get("Choose an available model before sending.")); return; }
            string requestText;
            ChatAttachment[] attachments;
            try
            {
                EnsureCurrentScope(); attachments = PrepareAttachments(question);
                requestText = ChatCommand.Expand(question);
                foreach (var attachment in attachments) requestText += "\n\n<context label=\"" + attachment.Label + "\">\n" + attachment.Text + "\n</context>";
            }
            catch (Exception ex) { SetStatus(UiText.Get("Context: ") + ex.Message); return; }
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope != null) requestText = UiText.Get("VBA project for this conversation: ") + scope.Label +
                UiText.Get("\nProject identifier to use in tools: ") + scope.Project + "\n\n" + requestText;
            string attachedMemory = attachMemory.Checked == true ? projectMemory : null;
            if (!string.IsNullOrWhiteSpace(attachedMemory))
                requestText += "\n\n<memoire-document>\n" + attachedMemory + "\n</memoire-document>";
            var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
            var attachedReferences = CurrentReferences(question);
            var pendingDraftAttachments = draftAttachments.ToArray();
            stopRequested = false;
            activeTurnId = Guid.NewGuid().ToString("N");
            tools.Mode = currentSession.Mode;
            requestText = UiText.Get("Mode for this request: ") + currentSession.Mode + (currentSession.Mode == ChatMode.Agent ? ".\n" : UiText.Get(". Analysis only; no edits or macro execution.\n")) + requestText;
            requestText = "<vbe-encoding-context>\n" + LlmVbeContext.EncodingInstructions +
                "\n</vbe-encoding-context>\n\n" + requestText;
            if (!string.IsNullOrEmpty(currentSession.ResumeContext)) requestText = UiText.Get("Branch history (context only; read the live code again):\n") + currentSession.ResumeContext + "\n\n" + requestText;
            streamedFinalText = null;
            RenameFromQuestion(question);
            SetBusy(true);
            prompt.Clear();
            HideReferences();
            selectedReferences.Clear();
            draftAttachments.Clear();
            attachMemory.Checked = false;
            followConversation = true;
            AddEntry(new ChatEntry { Speaker = "Vous", Text = question, References = attachedReferences, AttachedMemory = attachedMemory, Attachments = attachments, TurnId = activeTurnId });
            tools.NoteUserRequest(question);
            // Refresh the provider's system message even when a saved chat resumes on another ANSI code page.
            var firstMessage = messages.Count == 0 ? null :
                json.DeserializeObject(json.Serialize(messages[0])) as IDictionary<string, object>;
            if (firstMessage != null && firstMessage.ContainsKey("role") &&
                string.Equals(Convert.ToString(firstMessage["role"]), "system", StringComparison.OrdinalIgnoreCase))
                messages.RemoveAt(0);
            messages.Insert(0, new { role = "system", content = LlmVbeContext.DeveloperInstructions });
            int checkpoint = messages.Count;
            var provider = (LlmProvider)providerPicker.SelectedItem;
            tools.CurrentProviderName = provider.Name;
            if (!provider.IsCodex) messages.Add(new { role = "user", content = requestText });
            SaveCurrentSession();
            string providerStreamId = null;
            try
            {
                if (provider.IsCodex)
                {
                    if (codex == null && CodexTurnOverride == null)
                        codex = CreateCodexClient();
                    SetStatus(UiText.Get("Codex — working"));
                    CompleteAssistantResponse(await (CodexTurnOverride != null
                        ? CodexTurnOverride(requestText, selectedModel.Id, selectedEffort == null ? null : selectedEffort.Id)
                        : codex.TurnAsync(requestText, selectedModel.Id, selectedEffort == null ? null : selectedEffort.Id)));
                    currentSession.ResumeContext = null;
                    SetStatus(UiText.Get("Codex — ready"));
                    return;
                }
                using (var client = new LlmChatClient((LlmProvider)providerPicker.SelectedItem, settings,
                    selectedModel.Id, HttpHandlerOverride?.Invoke()))
                {
                    activeHttpClient = client;
                    client.ToolHandler = async (name, arguments) =>
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        Append("Outil", name);
                        return await InvokeTool(tools, name, arguments);
                    };
                    SetStatus(client.DisplayName + UiText.Get(" — working"));
                    for (int turn = 0; turn < 8; turn++)
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        providerStreamId = "http-" + Guid.NewGuid().ToString("N");
                        string streamId = providerStreamId;
                        bool receivedText = false;
                        client.TextDelta = fragment =>
                        {
                            if (stopRequested || IsDisposed) return;
                            receivedText = true;
                            ReceiveChatUpdate("final", streamId, fragment, false);
                        };
                        var message = await client.CompleteAsync(messages, LlmVbeTools.Definitions);
                        if (stopRequested) throw new OperationCanceledException();
                        if (receivedText) ReceiveChatUpdate("final", streamId, Convert.ToString(message["content"]), true);
                        providerStreamId = null;
                        messages.Add(message);
                        object rawCalls;
                        var calls = message.TryGetValue("tool_calls", out rawCalls) ? rawCalls as object[] : null;
                        if (calls == null || calls.Length == 0)
                        {
                            string answer = message.ContainsKey("content") ? Convert.ToString(message["content"]) : "";
                            CompleteAssistantResponse(string.IsNullOrWhiteSpace(answer) ? UiText.Get("No text response.") : answer);
                            currentSession.ResumeContext = null;
                            SetStatus(client.DisplayName + UiText.Get(" — ready"));
                            return;
                        }
                        foreach (object rawCall in calls)
                        {
                            if (stopRequested) throw new OperationCanceledException();
                            var call = rawCall as IDictionary<string, object>;
                            var function = call != null && call.ContainsKey("function") ? call["function"] as IDictionary<string, object> : null;
                            if (function == null || !call.ContainsKey("id")) throw new InvalidOperationException("Invalid tool call.");
                            string name = Convert.ToString(function["name"]);
                            string arguments = Convert.ToString(function["arguments"]);
                            Append("Outil", name);
                            string result = await InvokeTool(tools, name, arguments);
                            messages.Add(new { role = "tool", tool_call_id = Convert.ToString(call["id"]), content = result });
                        }
                    }
                    throw new InvalidOperationException("The assistant exceeded the tool-call limit.");
                }
            }
            catch (Exception ex)
            {
                if (providerStreamId != null && liveEntries.ContainsKey(providerStreamId)) ReceiveChatUpdate("tool", providerStreamId, null, true);
                if (provider.IsCodex && !stopRequested) { codex?.Dispose(); codex = null; }
                // A failed request must not leave an orphaned tool call in the next API request.
                messages.RemoveRange(checkpoint, messages.Count - checkpoint);
                if (!provider.IsCodex)
                {
                    messages.Add(new { role = "user", content = requestText });
                    messages.Add(new { role = "assistant", content = UiText.Get("The response did not complete. Actions may already have been applied; read the live code again before continuing.") });
                }
                Append(stopRequested ? "Assistant" : "Erreur", stopRequested ? UiText.Get("Response interrupted. Changes already applied can still be undone in the chat.") : ex.Message);
                SetStatus(stopRequested ? UiText.Get("Stopped") : UiText.Get("Error — you can resume the conversation"));
                if (!stopRequested && prompt.Text.Length == 0)
                {
                    selectedReferences.AddRange(attachedReferences);
                    draftAttachments.AddRange(pendingDraftAttachments);
                    attachMemory.Checked = !string.IsNullOrEmpty(attachedMemory);
                    prompt.Text = question;
                    RefreshContextChips();
                }
            }
            finally
            {
                activeHttpClient = null;
                if (!IsDisposed)
                {
                    var intervention = codeChanges.FindAll(x => x.TurnId == activeTurnId);
                    if (intervention.Count > 0) AddEntry(new ChatEntry
                    {
                        Speaker = "Intervention",
                        TurnId = activeTurnId,
                        Text = intervention.Count + UiText.Get(" change(s) applied. Review the files and undo this turn below.")
                    });
                    if (verifyAfterEdit.Checked == true && codeChanges.Exists(x => x.TurnId == activeTurnId)) await VerifyProjectAsync();
                    activeTurnId = null; SetBusy(false); SaveCurrentSession();
                }
            }
        }

        private void DisposeRuntime()
        {
            SaveCurrentSession(); saveTimer?.Stop(); projectRetryTimer?.Stop();
            sessionStore?.Dispose(); sessionStore = null;
            activeHttpClient?.Dispose(); codex?.Dispose(); codex = null;
            changes?.ContextMenuStrip?.Dispose();
            DisposeComposer();
        }
    }
}


namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        // Native defaults; tests may substitute only the external effects of a chat.
        internal static Func<LlmSettings> ReadSettings = LlmSettings.Load;
        internal static Action<LlmSettings> WriteSettings = (Action<LlmSettings>)Delegate.CreateDelegate(typeof(Action<LlmSettings>), typeof(LlmSettings).GetMethod("Save"));
        internal static Func<string> HistoryPath = DefaultHistoryPath;
        internal static Func<string, ChatSessionStore> OpenHistory = OpenHistoryNative;
        internal static Func<Form, IWin32Window, DialogResult> ShowModal = (Func<Form, IWin32Window, DialogResult>)Delegate.CreateDelegate(typeof(Func<Form, IWin32Window, DialogResult>), typeof(Form).GetMethod("ShowDialog", new[] { typeof(IWin32Window) }));
        internal static Func<CommonDialog, IWin32Window, DialogResult> ShowSaveDialog = (Func<CommonDialog, IWin32Window, DialogResult>)Delegate.CreateDelegate(typeof(Func<CommonDialog, IWin32Window, DialogResult>), typeof(CommonDialog).GetMethod("ShowDialog", new[] { typeof(IWin32Window) }));
        internal static Func<IWin32Window, string, string, MessageBoxButtons, MessageBoxIcon, DialogResult> ShowNotice = MessageBox.Show;
        internal static Func<System.Windows.Input.ModifierKeys> ReadModifiers = (Func<System.Windows.Input.ModifierKeys>)Delegate.CreateDelegate(typeof(Func<System.Windows.Input.ModifierKeys>), typeof(System.Windows.Input.Keyboard).GetProperty("Modifiers").GetGetMethod());
        internal static Action<string> WriteClipboard = System.Windows.Clipboard.SetText;
        internal static Func<ICodexAppServerTransport> TransportFactory = CreateNativeTransport;
        internal static Func<LlmProvider, LlmSettings, Task<LlmModelOption[]>> ReadModelCatalogue = LlmChatClient.ListModelsAsync;
        internal static Func<VbeSession, Request, Response> ReadHost = ReadHostNative;
        internal static Func<LlmVbeTools, string, string, Task<string>> InvokeTool = InvokeToolNative;
        private static string DefaultHistoryPath() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CodexVBE", "chat.db"); }
        private static ChatSessionStore OpenHistoryNative(string path) { return new ChatSessionStore(path); }
        private static ICodexAppServerTransport CreateNativeTransport() { return new CodexProcessTransport(); }
        private static Response ReadHostNative(VbeSession session, Request request) { return session.Execute(request); }
        private static Task<string> InvokeToolNative(LlmVbeTools tools, string name, string arguments) { return tools.InvokeAsync(name, arguments); }
    }
}
