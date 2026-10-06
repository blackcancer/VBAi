using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Fenêtre principale de conversation avec les fournisseurs LLM et les outils VBE.</summary>
    internal sealed partial class ChatWindow : Form
    {

        /// <summary>Configuration persistée des fournisseurs et options de la conversation.</summary>
        private readonly LlmSettings settings;

        /// <summary>Outils permettant au fournisseur d’interroger ou modifier le projet VBE.</summary>
        private readonly LlmVbeTools tools;

        /// <summary>Client Codex App Server actif, lorsqu’un fournisseur Codex est sélectionné.</summary>
        private CodexAppServerClient codex;

        /// <summary>Historique de messages transmis au fournisseur courant.</summary>
        private readonly List<object> messages = new List<object>();

        /// <summary>Modifications de code enregistrées dans la session.</summary>
        private readonly List<CodeChange> codeChanges = new List<CodeChange>();

        /// <summary>Sérialiseur des messages et charges utiles fournisseur.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };

        /// <summary>Indique qu’un tour de conversation est actif.</summary>
        private bool busy;

        /// <summary>Version du chargement de catalogue la plus récente, utilisée pour ignorer les réponses périmées.</summary>
        private int catalogueVersion;

        /// <summary>Bloque la persistance des sélections pendant leur restauration depuis la session.</summary>
        private bool restoringSelection;

        /// <summary>Indique qu’une annulation du tour courant a été demandée.</summary>
        private bool stopRequested;

        /// <summary>Client HTTP actif, disposé pour interrompre un tour non Codex.</summary>
        private LlmChatClient activeHttpClient;

        /// <summary>Substitution facultative du tour Codex, principalement destinée aux validations automatisées.</summary>
        internal Func<string, string, string, Task<string>> CodexTurnOverride;

        /// <summary>Substitution facultative de l’interruption Codex.</summary>
        internal Func<Task> CodexInterruptOverride;

        /// <summary>Substitution facultative du chargement de catalogue de modèles.</summary>
        internal Func<LlmProvider, Task<LlmModelOption[]>> ModelCatalogueOverride;

        /// <summary>Fabrique facultative du gestionnaire HTTP utilisé par les clients de chat.</summary>
        internal Func<HttpMessageHandler> HttpHandlerOverride;

        /// <summary>Crée la fenêtre et initialise ses contrôles et ressources visuelles.</summary>
        public ChatWindow()
        {
            InitializeComponent();
            pendingMessagesPanel.SizeChanged += (sender, args) => { foreach (Control row in pendingMessagesPanel.Controls) row.Width = Math.Max(200, pendingMessagesPanel.ClientSize.Width - 24); };
            Icon = VbeWindowIcons.Icon("assistant");
            github.Image = VbeWindowIcons.Image("github");
            configure.Image = VbeWindowIcons.Image("settings");
            using (var identity = VbeWindowIcons.Icon("assistant")) about.Image = identity.ToBitmap();
            UiText.Apply(this, components);
        }

        /// <summary>Crée la fenêtre connectée à la session VBE et initialise le compositeur, le transcript et les fournisseurs.</summary>
        /// <param name="session">Session VBE liée à cette conversation.</param>
        public ChatWindow(VbeSession session) : this()
        {
            InitializeShell();
            InitializeComposer(session);
            InitializeTranscript();
            foreach (var item in LlmProvider.All) providerPicker.Items.Add(item);

            try { settings = ReadSettings(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
            RefreshApprovalSelection();
            tools = new LlmVbeTools(session, this, settings);
            tools.ValidateScope = EnsureCurrentScope;
            tools.ValidateCachedScope = () => { EnsureCachedScope(); };
            tools.FormCut += change => AddEntry(new ChatEntry { Speaker = "Designer", FormCut = change });
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
            projectAccess.Click += (sender, args) => ConfigureProjectAccess();
            resumeTurn.Click += async (sender, args) => await ResumeBudgetAsync();
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
                RefreshModelSummary();
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
                RefreshModelSummary();
                if (!restoringSelection) settings.SetReasoningEffort(selectedProvider, selectedModel.Id, selectedEffort.Id);
                ScheduleSessionSave();
                if (!restoringSelection)
                    try { WriteSettings(settings); } catch (Exception ex) { LoadLog.Write("Reasoning effort save failed: " + ex.Message); }
            };
            send.Click += async (sender, args) => await SendAsync();
            int defaultProvider = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            loadingSession = true;
            try { providerPicker.SelectedIndex = defaultProvider < 0 ? 0 : defaultProvider; }
            finally { loadingSession = false; }
            InitializeSessions(session);
            if (scopePicker.SelectedItem == null) _ = LoadModelsAsync();
        }

        /// <summary>Remplit le sélecteur d’effort si le fournisseur et le modèle prennent en charge cette option.</summary>
        /// <param name="provider">Fournisseur sélectionné.</param>
        /// <param name="model">Modèle sélectionné.</param>
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

        /// <summary>Libère le client Codex et met à jour l’état de connexion au fournisseur sélectionné.</summary>
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

        /// <summary>Charge les modèles du fournisseur sélectionné et ignore les réponses obsolètes de requêtes précédentes.</summary>
        /// <returns>Tâche terminée après le chargement et l’actualisation des sélecteurs.</returns>
        private async Task LoadModelsAsync()
        {
            int version = ++catalogueVersion;
            var provider = providerPicker.SelectedItem as LlmProvider;
            modelPicker.Items.Clear();
            modelPicker.Enabled = false;
            effortPicker.Items.Clear();
            effortPicker.Enabled = false;
            refreshModels.Enabled = false;
            RefreshModelSummary();
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
                    RefreshModelSummary();
                }
            }
        }

        /// <summary>Affiche les paramètres et recharge le catalogue lorsque le fournisseur ou sa connexion change.</summary>
        /// <param name="owner">Fenêtre propriétaire facultative du dialogue.</param>
        public void ShowSettings(IWin32Window owner = null)
        {
            if (loadingScope) return;
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
                RefreshApprovalSelection();
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

        /// <summary>Ajoute un message au transcript si la fenêtre est encore active.</summary>
        /// <param name="speaker">Locuteur ou catégorie.</param>
        /// <param name="content">Contenu du message.</param>
        private void Append(string speaker, string content)
        {
            if (IsDisposed) return;
            AddTranscriptMessage(speaker, content);
        }

        /// <summary>Affiche le menu des changements de session ou fait défiler jusqu’au changement sélectionné.</summary>
        /// <param name="selected">Changement à afficher, ou nul pour ouvrir la liste des changements.</param>
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

        /// <summary>Met à jour le statut sur le thread UI et signale si l’historique n’a pas pu être enregistré.</summary>
        /// <param name="text">Texte de statut à présenter.</param>
        private void SetStatus(string text)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetStatus(text))); return; }
            status.Text = storageFailed && !ChatSessionStore.IsTransientScope(currentSession?.Scope) ? text + UiText.Get(" · History not saved") : text;
            toolTips.SetToolTip(status, status.Text);
        }

        /// <summary>Crée un client Codex associé à la session courante et relaie ses événements uniquement vers celle-ci.</summary>
        /// <returns>Client App Server configuré pour la fenêtre.</returns>
        private CodexAppServerClient CreateCodexClient()
        {
            ProviderSessionStorage.PrepareCodexSession(currentSession);
            var ownerSession = currentSession;
            var client = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                tools, SetStatus, settings, currentSession?.CodexThreadId,
                currentSession?.CodexDeveloperInstructionsHash, TransportFactory());
            client.ChatUpdate += (kind, id, text, complete) =>
            {
                if (currentSession == ownerSession && !IsDisposed && kind != "tool" && kind != "summary") ReceiveChatUpdate(kind, id, text, complete);
            };
            client.ActivityUpdate += activity => { if (currentSession == ownerSession && !IsDisposed) ReceiveAgentActivity(activity); };
            client.ThreadReady += id =>
            {
                if (currentSession == ownerSession && currentSession != null && !IsDisposed)
                { currentSession.CodexThreadId = id; currentSession.CodexThreadHome = ProviderSessionStorage.CodexHome;
                    currentSession.CodexDeveloperInstructionsHash = client.AppliedInstructionsHash; SaveCurrentSession(); }
            };
            return client;
        }

        /// <summary>Met à jour l’état actif des commandes et l’indicateur de progression.</summary>
        /// <param name="value">Indique si un tour est en cours.</param>
        private void SetBusy(bool value)
        {
            busy = value;
            UpdateDeleteSessionButton();
            modePicker.Enabled = !value;
            approvalPicker.Enabled = modelSummary.Enabled = !value;
            newChat.Enabled = scopePicker.Enabled = sessionList.Enabled =
                providerPicker.Enabled = refreshModels.Enabled = configure.Enabled = projectAccess.Enabled = !value;
            modelPicker.Enabled = !value && modelPicker.Items.Count > 0;
            effortPicker.Enabled = !value && effortPicker.Items.Count > 0;
            activityBar.Visible = value;
            RefreshCodeChangeCards();
            UpdateBudgetControls();
            RefreshPendingMessages();
        }

        /// <summary>Demande l’interruption du tour actif et désactive temporairement la commande d’arrêt.</summary>
        /// <returns>Tâche terminée lorsque la demande d’interruption aboutit ou échoue.</returns>
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

        /// <summary>Valide le contexte courant, construit les messages, exécute les appels de modèle et enregistre le résultat.</summary>
        /// <returns>Tâche terminée lorsque le tour et son nettoyage sont achevés.</returns>
        private async Task SendAsync()
        {
            if (loadingScope) return;
            if (busy)
            {
                if (string.IsNullOrWhiteSpace(prompt.Text)) await StopTurnAsync();
                else QueueComposerMessage();
                return;
            }
            await SendRequestAsync(null);
        }

        /// <summary>Builds and sends a chat request, then removes its queued message after dispatch succeeds.</summary>
        /// <param name="queued">Previously prepared queued message, or null to use the current composer contents.</param>
        /// <returns>Task that performs provider dispatch and updates the session after successful send preparation.</returns>
        private async Task SendRequestAsync(QueuedChatMessage queued)
        {
            if (loadingScope) return;
            string question = (queued?.Text ?? prompt.Text).Trim();
            if (!busy && question.Length == 0 && currentSession?.BudgetPaused == true) { await ResumeBudgetAsync(); return; }
            if (busy || question.Length == 0) return;
            var selectedModel = modelPicker.SelectedItem as LlmModelOption;
            if (selectedModel == null) { SetStatus(UiText.Get("Choose an available model before sending.")); return; }
            string requestText;
            ChatAttachment[] attachments;
            try
            {
                EnsureCurrentScope(); attachments = queued == null ? PrepareAttachments(question) : PrepareRequestAttachments(question, queued.References ?? new VbeChatReference[0], queued.Attachments ?? new ChatAttachment[0]);
                requestText = ChatCommand.Expand(question);
                foreach (var attachment in attachments) requestText += "\n\n<context label=\"" + attachment.Label + "\">\n" + attachment.Text + "\n</context>";
            }
            catch (Exception ex) { SetStatus(UiText.Get("Context: ") + ex.Message); return; }
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope != null) requestText = UiText.Get("VBA project for this conversation: ") + scope.Label +
                UiText.Get("\nProject identifier to use in tools: ") + scope.Project + "\n\n" + requestText;
            string attachedMemory = queued != null ? queued.Memory : (attachMemory.Checked == true ? queuedDraftMemory ?? projectMemory : null);
            if (!string.IsNullOrWhiteSpace(attachedMemory))
                requestText += "\n\n<memoire-document>\n" + attachedMemory + "\n</memoire-document>";
            var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
            var attachedReferences = queued?.References ?? CurrentReferences(question);
            var pendingDraftAttachments = queued?.Attachments ?? draftAttachments.ToArray();
            stopRequested = false;
            activeTurnId = Guid.NewGuid().ToString("N");
            currentSession.BudgetPaused = false;
            currentSession.CompletedToolActions = new List<string>();
            tools.Mode = currentSession.Mode;
            requestText = UiText.Get("Mode for this request: ") + currentSession.Mode + (currentSession.Mode == ChatMode.Agent ? ".\n" : UiText.Get(". Analysis only; no edits or macro execution.\n")) + requestText;
            requestText = "<vbe-encoding-context>\n" + LlmVbeContext.EncodingInstructions +
                "\n</vbe-encoding-context>\n\n" + requestText;
            if (((LlmProvider)providerPicker.SelectedItem).IsCodex) ProviderSessionStorage.PrepareCodexSession(currentSession);
            if (!string.IsNullOrEmpty(currentSession.ResumeContext)) requestText = UiText.Get("Branch history (context only; read the live code again):\n") + currentSession.ResumeContext + "\n\n" + requestText;
            streamedFinalText = null;
            RenameFromQuestion(question);
            if (queued != null) currentSession.PendingMessages.Remove(queued);
            SetBusy(true);
            if (queued == null)
            {
                prompt.Clear();
                HideReferences();
                selectedReferences.Clear();
                draftAttachments.Clear();
                attachMemory.Checked = false; queuedDraftMemory = null;
            }
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
            var provider = (LlmProvider)providerPicker.SelectedItem;
            tools.CurrentProviderName = provider.Name;
            if (!provider.IsCodex) messages.Add(new { role = "user", content = requestText });
            SaveCurrentSession();
            providerStreamId = null;
            bool completed = false;
            try
            {
                if (provider.IsCodex)
                {
                    if (codex == null && CodexTurnOverride == null)
                        codex = CreateCodexClient();
                    SetStatus(UiText.Get("Codex — working"));
                    string answer = await (CodexTurnOverride != null
                        ? CodexTurnOverride(requestText, selectedModel.Id, selectedEffort == null ? null : selectedEffort.Id)
                        : codex.TurnAsync(requestText, selectedModel.Id, selectedEffort == null ? null : selectedEffort.Id));
                    if (stopRequested) throw new OperationCanceledException();
                    CompleteAssistantResponse(answer);
                    currentSession.ResumeContext = null;
                    SetStatus(UiText.Get("Codex — ready"));
                    completed = !stopRequested;
                    return;
                }
                completed = await RunHttpBudgetAsync(provider, selectedModel.Id);
            }
            catch (Exception ex)
            {
                if (providerStreamId != null && liveEntries.ContainsKey(providerStreamId)) ReceiveChatUpdate("tool", providerStreamId, null, true);
                if (provider.IsCodex && !stopRequested) { codex?.Dispose(); codex = null; }
                // Preserve recorded results. Complete missing responses without replaying any call.
                CompletePendingToolResponses();
                if (!provider.IsCodex)
                {
                    messages.Add(new { role = "assistant", content = UiText.Get("The response did not complete. Actions may already have been applied; read the live code again before continuing.") });
                }
                Append(stopRequested ? "Assistant" : "Erreur", stopRequested ? UiText.Get("Response interrupted. Changes already applied can still be undone in the chat.") : ex.Message);
                SetStatus(stopRequested ? UiText.Get("Stopped") : UiText.Get("Error — you can resume the conversation"));
                if (!stopRequested && prompt.Text.Length == 0)
                {
                    selectedReferences.AddRange(attachedReferences);
                    draftAttachments.AddRange(pendingDraftAttachments);
                    queuedDraftMemory = attachedMemory;
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
                    if (currentSession?.BudgetPaused != true && verifyAfterEdit.Checked == true && codeChanges.Exists(x => x.TurnId == activeTurnId)) await VerifyProjectAsync();
                    activeTurnId = null; SetBusy(false); SaveCurrentSession();
                    await DispatchPendingAsync(completed);
                }
            }
        }

        /// <summary>One-use guard preventing runtime resources from being disposed more than once.</summary>
        private bool runtimeDisposed;

        /// <summary>Saves the session and releases each resource even if another cleanup fails.</summary>
        private void DisposeRuntime()
        {
            if (runtimeDisposed) return;
            runtimeDisposed = true;
            CleanupRuntime(() => contextMonitorTimer?.Stop());
            CleanupRuntime(() => referenceEvents?.Dispose());
            CleanupRuntime(() => projectEvents?.Dispose());
            CleanupRuntime(() => componentEvents?.Dispose());
            CleanupRuntime(SaveCurrentSession);
            CleanupRuntime(() => streamRenderTimer?.Stop());
            CleanupRuntime(() => pendingFollow?.Abort()); pendingFollow = null;
            CleanupRuntime(() => saveTimer?.Stop());
            CleanupRuntime(() => historySearchTimer?.Stop());
            CleanupRuntime(() => projectRetryTimer?.Stop());
            CleanupRuntime(DisposeScopeIdentities);
            CleanupRuntime(() => persistenceWorker?.Dispose()); persistenceWorker = null;
            CleanupRuntime(() => sessionStore?.Dispose()); sessionStore = null;
            CleanupRuntime(() => activeHttpClient?.Dispose()); activeHttpClient = null;
            CleanupRuntime(() => codex?.Dispose()); codex = null;
            CleanupRuntime(() => changes?.ContextMenuStrip?.Dispose());
            CleanupRuntime(() => { if (about?.Image != null) { about.Image.Dispose(); about.Image = null; } });
            CleanupRuntime(DisposeEntryViews);
            CleanupRuntime(DisposeComposer);
        }

        /// <summary>Runs one cleanup action and logs its failure so later resources are still released.</summary>
        /// <param name="cleanup">Resource cleanup callback to attempt.</param>
        private static void CleanupRuntime(Action cleanup)
        {
            try { cleanup(); }
            catch (Exception ex) { LoadLog.Write("Chat cleanup failed: " + ex.Message); }
        }
    }
}


namespace VBAi
{

    /// <summary>Fenêtre de conversation avec adaptateurs substituables pour les effets externes.</summary>
    internal sealed partial class ChatWindow
    {
        // Native defaults; tests may substitute only the external effects of a chat.
        /// <summary>Charge les paramètres persistants des fournisseurs.</summary>
        internal static Func<LlmSettings> ReadSettings = LlmSettings.Load;

        /// <summary>Enregistre les paramètres persistants des fournisseurs.</summary>
        internal static Action<LlmSettings> WriteSettings = (Action<LlmSettings>)Delegate.CreateDelegate(typeof(Action<LlmSettings>), typeof(LlmSettings).GetMethod("Save"));

        /// <summary>Retourne le chemin de la base de sessions.</summary>
        internal static Func<string> HistoryPath = DefaultHistoryPath;

        /// <summary>Ouvre le stockage des sessions au chemin donné.</summary>
        internal static Func<string, ChatSessionStore> OpenHistory = OpenHistoryNative;

        /// <summary>Affiche un formulaire modal avec son propriétaire.</summary>
        internal static Func<Form, IWin32Window, DialogResult> ShowModal = (Func<Form, IWin32Window, DialogResult>)Delegate.CreateDelegate(typeof(Func<Form, IWin32Window, DialogResult>), typeof(Form).GetMethod("ShowDialog", new[] { typeof(IWin32Window) }));

        /// <summary>Affiche une boîte de dialogue d’enregistrement avec son propriétaire.</summary>
        internal static Func<CommonDialog, IWin32Window, DialogResult> ShowSaveDialog = (Func<CommonDialog, IWin32Window, DialogResult>)Delegate.CreateDelegate(typeof(Func<CommonDialog, IWin32Window, DialogResult>), typeof(CommonDialog).GetMethod("ShowDialog", new[] { typeof(IWin32Window) }));

        /// <summary>Affiche une notification à l’utilisateur.</summary>
        internal static Func<IWin32Window, string, string, MessageBoxButtons, MessageBoxIcon, DialogResult> ShowNotice = MessageBox.Show;

        /// <summary>Lit les touches de modification courantes du clavier.</summary>
        internal static Func<System.Windows.Input.ModifierKeys> ReadModifiers = (Func<System.Windows.Input.ModifierKeys>)Delegate.CreateDelegate(typeof(Func<System.Windows.Input.ModifierKeys>), typeof(System.Windows.Input.Keyboard).GetProperty("Modifiers").GetGetMethod());

        /// <summary>Copie du texte dans le presse-papiers Windows.</summary>
        internal static Action<string> WriteClipboard = System.Windows.Clipboard.SetText;

        /// <summary>Crée le transport natif vers Codex App Server.</summary>
        internal static Func<ICodexAppServerTransport> TransportFactory = CreateNativeTransport;

        /// <summary>Charge le catalogue des modèles d’un fournisseur.</summary>
        internal static Func<LlmProvider, LlmSettings, Task<LlmModelOption[]>> ReadModelCatalogue = LlmChatClient.ListModelsAsync;

        /// <summary>Exécute une commande protocole dans la session VBE.</summary>
        internal static Func<VbeSession, Request, Response> ReadHost = ReadHostNative;

        /// <summary>Exécute un outil VBE par son nom et ses arguments JSON.</summary>
        internal static Func<LlmVbeTools, string, string, Task<string>> InvokeTool = InvokeToolNative;

        /// <summary>Calcule l’emplacement par défaut de la base de sessions.</summary>
        /// <returns>Chemin local par défaut de chat.db.</returns>
        private static string DefaultHistoryPath() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBAi", "chat.db"); }

        /// <summary>Ouvre une instance native du stockage de sessions.</summary>
        /// <param name="path">Chemin du fichier de stockage des sessions.</param>
        /// <returns>Stockage natif associé au chemin fourni.</returns>
        private static ChatSessionStore OpenHistoryNative(string path) { return new ChatSessionStore(path); }

        /// <summary>Crée le transport processus vers Codex App Server.</summary>
        /// <returns>Transport Codex App Server démarré à la demande.</returns>
        private static ICodexAppServerTransport CreateNativeTransport() { return new CodexProcessTransport(); }

        /// <summary>Délègue l’exécution à la session VBE.</summary>
        /// <param name="session">Session VBE qui exécute la commande.</param>
        /// <param name="request">Commande reçue par le pont local.</param>
        /// <returns>Réponse du traitement de commande VBE.</returns>
        private static Response ReadHostNative(VbeSession session, Request request) { return session.Execute(request); }

        /// <summary>Délègue l’appel au service d’outils VBE.</summary>
        /// <param name="tools">Service d’outils à invoquer.</param>
        /// <param name="name">Nom stable de l’outil.</param>
        /// <param name="arguments">Arguments de l’outil sérialisés en JSON.</param>
        /// <returns>Tâche produisant le résultat sérialisé de l’outil.</returns>
        private static Task<string> InvokeToolNative(LlmVbeTools tools, string name, string arguments) { return tools.InvokeAsync(name, arguments); }

        /// <summary>Suppresses the approval selection handler while restoring policy from settings.</summary>
        private bool refreshingApproval;

        /// <summary>Restores the policy without treating restoration as a user edit.</summary>
        private void RefreshApprovalSelection()
        {
            refreshingApproval = true;
            try { approvalPicker.SelectedIndex = settings?.VbeEditApproval == "ReadOnly" ? 2 : settings?.VbeEditApproval == "AskEachTime" ? 1 : settings == null || settings.VbeEditApproval == "Automatic" ? 0 : -1; }
            finally { refreshingApproval = false; }
        }

        /// <summary>Changes the same VBE policy consumed by both native and HTTP tools.</summary>
        /// <param name="sender">Approval picker control raising the selection event.</param>
        /// <param name="e">Native event data.</param>
        private void ApprovalPicker_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (refreshingApproval || settings == null) return;
            if (busy || approvalPicker.SelectedIndex < 0) { RefreshApprovalSelection(); return; }
            string previous = settings.VbeEditApproval;
            settings.VbeEditApproval = approvalPicker.SelectedIndex == 2 ? "ReadOnly" : approvalPicker.SelectedIndex == 1 ? "AskEachTime" : "Automatic";
            try { WriteSettings(settings); }
            catch (Exception error) { settings.VbeEditApproval = previous; RefreshApprovalSelection(); SetStatus(error.Message); }
        }

        /// <summary>Mirrors the menu choice into the existing verification preference.</summary>
        /// <param name="sender">Verification preference control raising the change event.</param>
        /// <param name="e">Native event data.</param>
        private void VerifyChanges_CheckedChanged(object sender, EventArgs e)
        {
            if (verifyAfterEdit != null) verifyAfterEdit.Checked = verifyChanges.Checked;
        }

        /// <summary>Reveals the Designer-built provider, model and effort selectors.</summary>
        /// <param name="sender">Collapsed provider/model summary control raising the click event.</param>
        /// <param name="e">Native event data.</param>
        private void ModelSummary_Click(object sender, EventArgs e)
        {
            if (busy) return;
            bool expanded = rootLayout.RowStyles[6].Height == 0;
            rootLayout.RowStyles[6].Height = expanded ? providerLayout.Controls.Cast<System.Windows.Forms.Control>()
                .Max(control => Math.Max(control.Height, control.GetPreferredSize(System.Drawing.Size.Empty).Height) + control.Margin.Vertical)
                + providerLayout.Padding.Vertical + providerLayout.Margin.Vertical : 0;
            providerLayout.Visible = expanded;
            RefreshModelSummary();
            if (expanded) modelPicker.Focus();
        }

        /// <summary>Shows the selected model succinctly while retaining its full identity in the tooltip.</summary>
        private void RefreshModelSummary()
        {
            string model = (modelPicker.SelectedItem as LlmModelOption)?.Label ?? modelPicker.Text;
            string effort = effortPicker.Items.Count > 0 ? effortPicker.Text : "";
            modelSummary.Text = (string.IsNullOrWhiteSpace(model) ? UiText.Get("Model") : model) +
                (string.IsNullOrWhiteSpace(effort) ? "" : " · " + effort) + (rootLayout.RowStyles[6].Height > 0 ? " ▴" : " ▾");
            toolTips.SetToolTip(modelSummary, providerPicker.Text + " · " + modelPicker.Text +
                (string.IsNullOrWhiteSpace(effort) ? "" : " · " + effort));
        }
    }
}
