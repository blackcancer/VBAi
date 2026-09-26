using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class ChatWindow : Form
    {
        private readonly TextBox transcript;
        private readonly TextBox prompt;
        private readonly Button send;
        private readonly Label status;
        private readonly ComboBox providerPicker;
        private readonly ComboBox modelPicker;
        private readonly ComboBox effortPicker;
        private readonly Button refreshModels;
        private readonly Button configure;
        private readonly LlmSettings settings;
        private readonly LlmVbeTools tools;
        private CodexAppServerClient codex;
        private readonly List<object> messages = new List<object>();
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private bool busy;
        private int catalogueVersion;

        public ChatWindow(VbeSession session)
        {
            Text = "CodexVBE — Assistant";
            Width = 760;
            Height = 700;
            MinimumSize = new Size(650, 470);
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = true;
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;

            transcript = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BackColor = Color.White, Font = new Font("Segoe UI", 9.5F),
                BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(4, 3, 4, 8) };
            prompt = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, Margin = new Padding(4, 3, 4, 4) };
            send = new Button { Text = "Envoyer", Dock = DockStyle.Fill, Margin = new Padding(5, 4, 0, 4) };
            status = new Label { Text = "Chargement…", Dock = DockStyle.Fill, AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(4, 0, 0, 0) };
            providerPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(4, 8, 10, 4) };
            providerPicker.Items.AddRange(LlmProvider.All);
            modelPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(4, 7, 8, 4) };
            effortPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(4, 7, 8, 4), Enabled = false };
            refreshModels = new Button { Text = "Actualiser", Dock = DockStyle.Fill, Margin = new Padding(3, 4, 3, 4) };
            configure = new Button { Text = "Configuration…", Dock = DockStyle.Fill, Margin = new Padding(0, 5, 0, 5) };

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12),
                ColumnCount = 1, RowCount = 5 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132));
            header.Controls.Add(new Label { Text = "Fournisseur", Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            header.Controls.Add(providerPicker, 1, 0);
            header.Controls.Add(configure, 2, 0);
            var actions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 1 };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 87));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            actions.Controls.Add(new Label { Text = "Modèle", Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            actions.Controls.Add(modelPicker, 1, 0);
            actions.Controls.Add(new Label { Text = "Raisonnement", Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft }, 2, 0);
            actions.Controls.Add(effortPicker, 3, 0);
            actions.Controls.Add(refreshModels, 4, 0);
            actions.Controls.Add(send, 5, 0);
            layout.Controls.Add(header, 0, 0);
            layout.Controls.Add(transcript, 0, 1);
            layout.Controls.Add(prompt, 0, 2);
            layout.Controls.Add(actions, 0, 3);
            layout.Controls.Add(status, 0, 4);
            Controls.Add(layout);

            tools = new LlmVbeTools(session, this);
            try { settings = LlmSettings.Load(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
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
            messages.Add(new { role = "system", content =
                "You are a VBE assistant. Use tools to inspect the live VBA project before stating facts about it. " +
                "Read current code or form state before edits; pass the returned revision to every edit. " +
                "Do not invent project, module, form or control names. Keep answers concise and in the user's language." });
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
            int checkpoint = messages.Count;
            var provider = (LlmProvider)providerPicker.SelectedItem;
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
                            string result = tools.Invoke(name, arguments);
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
