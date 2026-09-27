using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
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

        public ChatWindow()
        {
            InitializeComponent();
        }

        public ChatWindow(VbeSession session) : this()
        {
            InitializeShell();
            InitializeComposer(session);
            InitializeTranscript();
            foreach (var item in LlmProvider.All) providerPicker.Items.Add(item);

            try { settings = LlmSettings.Load(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
            tools = new LlmVbeTools(session, this, settings);
            tools.ValidateScope = EnsureCurrentScope;
            tools.CodeEdited += change => {
                change.TurnId = activeTurnId;
                codeChanges.Add(change);
                changes.Enabled = true;
                changes.Text = "Modifications · " + codeChanges.Count;
                AddCodeChangeCard(change);
                RefreshCodeChangeCards();
            };
            changes.Click += (sender, args) => ShowCodeChanges();
            providerPicker.SelectedIndexChanged += (sender, args) => {
                if (loadingSession || busy) return;
                var provider = providerPicker.SelectedItem as LlmProvider;
                if (provider == null) return;
                settings.ProviderName = provider.Name;
                try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Provider selection save failed: " + ex.Message); }
                if (scopePicker.SelectedItem == null)
                {
                    ResetProviderConnection();
                    _ = LoadModelsAsync();
                }
                else NewSession(provider.Name);
            };
            configure.Click += (sender, args) => ShowSettings();
            refreshModels.Click += async (sender, args) => await LoadModelsAsync();
            modelPicker.SelectedIndexChanged += (sender, args) => {
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
                    try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Model selection save failed: " + ex.Message); }
            };
            effortPicker.SelectedIndexChanged += (sender, args) => {
                var selectedProvider = providerPicker.SelectedItem as LlmProvider;
                var selectedModel = modelPicker.SelectedItem as LlmModelOption;
                var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
                if (selectedProvider == null || selectedModel == null || selectedEffort == null) return;
                if (currentSession != null) currentSession.Effort = selectedEffort.Id;
                if (!restoringSelection) settings.SetReasoningEffort(selectedProvider, selectedModel.Id, selectedEffort.Id);
                ScheduleSessionSave();
                if (!restoringSelection)
                    try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Reasoning effort save failed: " + ex.Message); }
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
                "Chargement des modèles du fournisseur…" : "Ce fournisseur n'est pas encore implémenté.";
            settings.ProviderName = ((LlmProvider)providerPicker.SelectedItem).Name;
            try { settings.Save(); }
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
                if (provider.IsCodex)
                {
                    if (codex == null)
                        codex = CreateCodexClient();
                    models = await codex.ListModelsAsync();
                }
                else models = await LlmChatClient.ListModelsAsync(provider, settings);
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
                SetStatus(models.Length == 0 ? "Aucun modèle disponible pour ce fournisseur." : provider.Name + " — prêt");
            }
            catch (Exception ex)
            {
                if (!IsDisposed && version == catalogueVersion)
                    SetStatus("Catalogue indisponible : " + ex.Message);
                LoadLog.Write("Chat model catalogue failed: " + provider.Name + " " + ex);
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

        public void ShowSettings()
        {
            if (busy) return;
            string previousProvider = settings.ProviderName;
            string previousOpenAiEndpoint = settings.OpenAiEndpoint;
            string previousOllamaEndpoint = settings.OllamaEndpoint;
            string previousKey = settings.EncryptedOpenAiKey;
            string previousProviders = json.Serialize(new { settings.ProviderEndpoints, settings.EncryptedProviderKeys, settings.AzureUseEntraToken, settings.ManualModelLists });
            using (var dialog = new LlmSettingsWindow(settings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
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
            foreach (var pair in entryViews)
                if (pair.Key.Change == selected)
                {
                    followConversation = false;
                    pair.Value.BringIntoView();
                    break;
                }
        }

        private void SetStatus(string text)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetStatus(text))); return; }
            status.Text = storageFailed ? text + " · Historique non enregistré" : text;
            toolTips.SetToolTip(status, status.Text);
        }

        private CodexAppServerClient CreateCodexClient()
        {
            var ownerSession = currentSession;
            var client = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                tools, SetStatus, settings, currentSession?.CodexThreadId);
            client.ChatUpdate += (kind, id, text, complete) => {
                if (currentSession == ownerSession && !IsDisposed) ReceiveChatUpdate(kind, id, text, complete);
            };
            client.ThreadReady += id => {
                if (currentSession == ownerSession && currentSession != null && !IsDisposed)
                { currentSession.CodexThreadId = id; SaveCurrentSession(); }
            };
            return client;
        }

        private void SetBusy(bool value)
        {
            busy = value;
            modePicker.Enabled = !value;
            send.Text = value ? "Arrêter ■" : "Envoyer ↑";
            toolTips.SetToolTip(send, value ? "Interrompre la réponse en cours. Les modifications déjà appliquées restent annulables dans le chat." : "Envoyer le message et son contexte à l’agent.");
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
            SetStatus("Arrêt en cours…");
            try
            {
                if (codex != null) await codex.InterruptAsync();
                else activeHttpClient?.Dispose();
            }
            catch (Exception ex) { SetStatus("Arrêt impossible : " + ex.Message); stopRequested = false; send.Enabled = true; }
        }

        private async Task SendAsync()
        {
            string question = prompt.Text.Trim();
            if (busy || question.Length == 0) return;
            var selectedModel = modelPicker.SelectedItem as LlmModelOption;
            if (selectedModel == null) { SetStatus("Choisissez un modèle disponible avant d'envoyer."); return; }
            string requestText;
            ChatAttachment[] attachments;
            try {
                EnsureCurrentScope(); attachments = PrepareAttachments(question);
                requestText = ChatCommand.Expand(question);
                foreach (var attachment in attachments) requestText += "\n\n<context label=\"" + attachment.Label + "\">\n" + attachment.Text + "\n</context>";
            }
            catch (Exception ex) { SetStatus("Contexte : " + ex.Message); return; }
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope != null) requestText = "Projet VBA de cette conversation : " + scope.Label +
                "\nIdentifiant Project à utiliser dans les outils : " + scope.Project + "\n\n" + requestText;
            string attachedMemory = attachMemory.Checked == true ? projectMemory : null;
            if (!string.IsNullOrWhiteSpace(attachedMemory))
                requestText += "\n\n<memoire-document>\n" + attachedMemory + "\n</memoire-document>";
            var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
            var attachedReferences = CurrentReferences(question);
            var pendingDraftAttachments = draftAttachments.ToArray();
            stopRequested = false;
            activeTurnId = Guid.NewGuid().ToString("N");
            tools.Mode = currentSession.Mode;
            requestText = "Mode de cette demande : " + currentSession.Mode + (currentSession.Mode == ChatMode.Agent ? ".\n" : ". Analyse uniquement ; aucune modification ni exécution de macro.\n") + requestText;
            requestText = "<vbe-encoding-context>\n" + LlmVbeContext.EncodingInstructions +
                "\n</vbe-encoding-context>\n\n" + requestText;
            if (!string.IsNullOrEmpty(currentSession.ResumeContext)) requestText = "Historique de la branche (contexte uniquement ; relire le code vivant) :\n" + currentSession.ResumeContext + "\n\n" + requestText;
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
                    if (codex == null)
                        codex = CreateCodexClient();
                    SetStatus("Codex — en cours");
                    CompleteAssistantResponse(await codex.TurnAsync(requestText, selectedModel.Id,
                        selectedEffort == null ? null : selectedEffort.Id));
                    currentSession.ResumeContext = null;
                    SetStatus("Codex — prêt");
                    return;
                }
                using (var client = new LlmChatClient((LlmProvider)providerPicker.SelectedItem, settings, selectedModel.Id))
                {
                    activeHttpClient = client;
                    client.ToolHandler = async (name, arguments) => {
                        if (stopRequested) throw new OperationCanceledException();
                        Append("Outil", name);
                        return await tools.InvokeAsync(name, arguments);
                    };
                    SetStatus(client.DisplayName + " — en cours");
                    for (int turn = 0; turn < 8; turn++)
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        providerStreamId = "http-" + Guid.NewGuid().ToString("N");
                        string streamId = providerStreamId;
                        bool receivedText = false;
                        client.TextDelta = fragment => {
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
                            CompleteAssistantResponse(string.IsNullOrWhiteSpace(answer) ? "Aucune réponse textuelle." : answer);
                            currentSession.ResumeContext = null;
                            SetStatus(client.DisplayName + " — prêt");
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
                            string result = await tools.InvokeAsync(name, arguments);
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
                    messages.Add(new { role = "assistant", content = "La réponse n'a pas abouti. Des actions peuvent déjà avoir été appliquées ; relire le code vivant avant de continuer." });
                }
                Append(stopRequested ? "Assistant" : "Erreur", stopRequested ? "Réponse interrompue. Les modifications déjà appliquées restent annulables dans le chat." : ex.Message);
                SetStatus(stopRequested ? "Arrêté" : "Erreur — vous pouvez reprendre la conversation");
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
                if (!IsDisposed) {
                    var intervention = codeChanges.FindAll(x => x.TurnId == activeTurnId);
                    if (intervention.Count > 0) AddEntry(new ChatEntry { Speaker = "Intervention", TurnId = activeTurnId,
                        Text = intervention.Count + " modification(s) appliquée(s). Retrouvez les fichiers et annulez cette intervention ci-dessous." });
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
