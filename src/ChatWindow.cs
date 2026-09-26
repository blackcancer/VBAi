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
            Width = 620;
            Height = 720;
            MinimumSize = new Size(420, 400);
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = true;
            Font = new Font("Segoe UI", 9F);

            transcript = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BackColor = Color.White, Font = new Font("Consolas", 9F) };
            prompt = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, Height = 100 };
            send = new Button { Text = "Envoyer", Dock = DockStyle.Right, Width = 100 };
            status = new Label { Text = "Modèle non configuré ou prêt à répondre", Dock = DockStyle.Top,
                Height = 25, AutoEllipsis = true };
            providerPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            providerPicker.Items.AddRange(LlmProvider.All);
            modelPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            refreshModels = new Button { Text = "Actualiser", Dock = DockStyle.Right, Width = 90 };
            configure = new Button { Text = "Configuration…", Dock = DockStyle.Right, Width = 130 };
            var toolbar = new TableLayoutPanel { Dock = DockStyle.Top, Height = 34, ColumnCount = 2, RowCount = 1 };
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var providerRow = new Panel { Dock = DockStyle.Fill };
            providerRow.Controls.Add(providerPicker);
            providerRow.Controls.Add(configure);
            var modelRow = new Panel { Dock = DockStyle.Bottom, Height = 34 };
            modelRow.Controls.Add(modelPicker);
            modelRow.Controls.Add(refreshModels);
            toolbar.Controls.Add(new Label { Text = "Fournisseur", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            toolbar.Controls.Add(providerRow, 1, 0);
            var composer = new Panel { Dock = DockStyle.Bottom, Height = 110 };
            composer.Controls.Add(prompt);
            composer.Controls.Add(send);
            Controls.Add(transcript);
            Controls.Add(modelRow);
            Controls.Add(composer);
            Controls.Add(status);
            Controls.Add(toolbar);

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
                try { settings.Save(); } catch (Exception ex) { LoadLog.Write("Model selection save failed: " + ex.Message); }
            };
            send.Click += async (sender, args) => await SendAsync();
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
            using (var dialog = new LlmSettingsWindow(settings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
                if (selected >= 0 && providerPicker.SelectedIndex != selected) providerPicker.SelectedIndex = selected;
                else { ResetConversation(); _ = LoadModelsAsync(); }
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
            busy = true;
            send.Enabled = false;
            providerPicker.Enabled = false;
            modelPicker.Enabled = false;
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
                    Append("Assistant", await codex.TurnAsync(question, selectedModel.Id));
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
                    modelPicker.Enabled = modelPicker.Items.Count > 0; refreshModels.Enabled = true; configure.Enabled = true; }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { codex?.Dispose(); codex = null; }
            base.Dispose(disposing);
        }
    }
}
