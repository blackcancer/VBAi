using System;
using System.Collections.Generic;
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
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private bool busy;
        private int catalogueVersion;

        public ChatWindow()
        {
            InitializeComponent();
        }

        public ChatWindow(VbeSession session)
        {
            InitializeComponent();
            providerPicker.Items.AddRange(LlmProvider.All);

            try { settings = LlmSettings.Load(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
            tools = new LlmVbeTools(session, this, settings);
            providerPicker.SelectedIndexChanged += async (sender, args) => { ResetConversation(); await LoadModelsAsync(); };
            int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            providerPicker.SelectedIndex = selected < 0 ? 0 : selected;
            configure.Click += (sender, args) => ShowSettings();
            refreshModels.Click += async (sender, args) => await LoadModelsAsync();
            modelPicker.SelectedIndexChanged += (sender, args) => {
                var selectedModel = modelPicker.SelectedItem as LlmModelOption;
                var selectedProvider = providerPicker.SelectedItem as LlmProvider;
                if (selectedModel == null || selectedProvider == null) return;
                settings.SetSelectedModel(selectedProvider, selectedModel.Id);
                UpdateEfforts(selectedProvider, selectedModel);
                try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Model selection save failed: " + ex.Message); }
            };
            effortPicker.SelectedIndexChanged += (sender, args) => {
                var selectedProvider = providerPicker.SelectedItem as LlmProvider;
                var selectedModel = modelPicker.SelectedItem as LlmModelOption;
                var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
                if (selectedProvider == null || selectedModel == null || selectedEffort == null) return;
                settings.SetReasoningEffort(selectedProvider, selectedModel.Id, selectedEffort.Id);
                try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Reasoning effort save failed: " + ex.Message); }
            };
            send.Click += async (sender, args) => await SendAsync();
        }

        private void UpdateEfforts(LlmProvider provider, LlmModelOption model)
        {
            effortPicker.Items.Clear();
            effortPicker.Enabled = false;
            if (!provider.IsCodex || model.Efforts.Length == 0) return;
            effortPicker.Items.AddRange(model.Efforts);
            string selected = settings.GetReasoningEffort(provider, model.Id);
            int index = Array.FindIndex(model.Efforts, item => item.Id == selected);
            if (index < 0) index = Array.FindIndex(model.Efforts, item => item.Id == model.DefaultEffort);
            effortPicker.SelectedIndex = index < 0 ? 0 : index;
            effortPicker.Enabled = !busy;
        }

        private void ResetConversation()
        {
            codex?.Dispose();
            codex = null;
            messages.Clear();
            messages.Add(new { role = "system", content = LlmVbeContext.DeveloperInstructions });
            transcript.Text = "CodexVBE est connecté au VBE. Écrivez une demande puis cliquez sur Envoyer.\r\n";
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
                        codex = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                            tools, SetStatus, settings);
                    models = await codex.ListModelsAsync();
                }
                else models = await LlmChatClient.ListModelsAsync(provider, settings);
                if (IsDisposed || version != catalogueVersion || provider != providerPicker.SelectedItem) return;
                modelPicker.Items.AddRange(models);
                string saved = settings.GetSelectedModel(provider);
                int selected = Array.FindIndex(models, item => item.Id == saved);
                if (selected < 0) selected = Array.FindIndex(models, item => item.IsDefault);
                if (selected < 0 && models.Length > 0) selected = 0;
                if (selected >= 0) modelPicker.SelectedIndex = selected;
                SetStatus(models.Length == 0 ? "Aucun modèle disponible pour ce fournisseur." : provider.Name + " — prêt");
            }
            catch (Exception ex)
            {
                if (!IsDisposed && version == catalogueVersion)
                    SetStatus("Catalogue indisponible : " + ex.Message);
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
            using (var dialog = new LlmSettingsWindow(settings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
                if (selected >= 0 && providerPicker.SelectedIndex != selected) providerPicker.SelectedIndex = selected;
                else
                {
                    bool connectionChanged = previousProvider != settings.ProviderName ||
                        previousOpenAiEndpoint != settings.OpenAiEndpoint ||
                        previousOllamaEndpoint != settings.OllamaEndpoint || previousKey != settings.EncryptedOpenAiKey;
                    if (connectionChanged) ResetConversation();
                    _ = LoadModelsAsync();
                }
            }
        }

        private void Append(string speaker, string content)
        {
            if (IsDisposed || transcript.IsDisposed) return;
            transcript.AppendText("\r\n" + speaker + " : " + content + "\r\n");
        }

        private void SetStatus(string text)
        {
            if (!IsDisposed && !status.IsDisposed) status.Text = text;
        }

        private async Task SendAsync()
        {
            string question = prompt.Text.Trim();
            if (busy || question.Length == 0) return;
            var selectedModel = modelPicker.SelectedItem as LlmModelOption;
            if (selectedModel == null) { SetStatus("Choisissez un modèle disponible avant d'envoyer."); return; }
            var selectedEffort = effortPicker.SelectedItem as LlmEffortOption;
            busy = true;
            send.Enabled = false;
            providerPicker.Enabled = false;
            modelPicker.Enabled = false;
            effortPicker.Enabled = false;
            refreshModels.Enabled = false;
            configure.Enabled = false;
            prompt.Clear();
            Append("Vous", question);
            tools.NoteUserRequest(question);
            int checkpoint = messages.Count;
            var provider = (LlmProvider)providerPicker.SelectedItem;
            tools.CurrentProviderName = provider.Name;
            settings.ProviderName = provider.Name;
            try { settings.Save(); }
            catch (Exception saveError) { LoadLog.Write("LLM settings save failed: " + saveError.Message); }
            if (!provider.IsCodex) messages.Add(new { role = "user", content = question });
            try
            {
                if (provider.IsCodex)
                {
                    if (codex == null)
                        codex = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                            tools, SetStatus, settings);
                    SetStatus("Codex — en cours");
                    Append("Assistant", await codex.TurnAsync(question, selectedModel.Id,
                        selectedEffort == null ? null : selectedEffort.Id));
                    SetStatus("Codex — prêt");
                    return;
                }
                using (var client = new LlmChatClient((LlmProvider)providerPicker.SelectedItem, settings, selectedModel.Id))
                {
                    SetStatus(client.DisplayName + " — en cours");
                    for (int turn = 0; turn < 8; turn++)
                    {
                        var message = await client.CompleteAsync(messages, LlmVbeTools.Definitions);
                        messages.Add(message);
                        object rawCalls;
                        var calls = message.TryGetValue("tool_calls", out rawCalls) ? rawCalls as object[] : null;
                        if (calls == null || calls.Length == 0)
                        {
                            string answer = message.ContainsKey("content") ? Convert.ToString(message["content"]) : "";
                            Append("Assistant", string.IsNullOrWhiteSpace(answer) ? "Aucune réponse textuelle." : answer);
                            SetStatus(client.DisplayName + " — prêt");
                            return;
                        }
                        foreach (object rawCall in calls)
                        {
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
                if (provider.IsCodex) { codex?.Dispose(); codex = null; }
                // A failed request must not leave an orphaned tool call in the next API request.
                messages.RemoveRange(checkpoint, messages.Count - checkpoint);
                Append("Erreur", ex.Message);
                SetStatus("Erreur — vérifiez la configuration et réessayez");
            }
            finally
            {
                busy = false;
                if (!IsDisposed) { send.Enabled = true; providerPicker.Enabled = true;
                    modelPicker.Enabled = modelPicker.Items.Count > 0;
                    effortPicker.Enabled = effortPicker.Items.Count > 0;
                    refreshModels.Enabled = true; configure.Enabled = true; }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { codex?.Dispose(); codex = null; }
            base.Dispose(disposing);
        }
    }
}
