using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class LlmSettingsWindow : Form
    {
        private readonly ComboBox provider = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        private readonly Label codexStatus = new Label { Dock = DockStyle.Fill, Text = "Vérification de ChatGPT…", AutoEllipsis = true };
        private readonly Button codexLogin = new Button { Text = "Se connecter à ChatGPT", Width = 185 };
        private readonly Button codexRefresh = new Button { Text = "Actualiser l'état", Width = 130 };
        private readonly TextBox openAiEndpoint = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox ollamaEndpoint = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox openAiKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        private readonly CheckBox clearKey = new CheckBox { Text = "Supprimer la clé API enregistrée", Dock = DockStyle.Fill };
        private readonly LlmSettings settings;
        private readonly TableLayoutPanel grid;

        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            Text = "CodexVBE — Configuration LLM";
            Width = 640;
            Height = 270;
            MinimumSize = new Size(560, 235);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            provider.Items.AddRange(LlmProvider.All);
            int current = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            provider.SelectedIndex = current < 0 ? 0 : current;
            openAiEndpoint.Text = settings.OpenAiEndpoint ?? "";
            ollamaEndpoint.Text = settings.OllamaEndpoint ?? "";

            grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 9,
                Padding = new Padding(12), AutoScroll = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            AddRow(grid, 0, "Fournisseur", provider);
            AddRow(grid, 1, "Compte ChatGPT", codexStatus);
            var codexActions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0) };
            codexActions.Controls.Add(codexLogin);
            codexActions.Controls.Add(codexRefresh);
            AddRow(grid, 2, "Authentification", codexActions);
            AddRow(grid, 3, "URL OpenAI (facultatif)", openAiEndpoint);
            AddRow(grid, 4, "URL Ollama (facultatif)", ollamaEndpoint);
            AddRow(grid, 5, "Nouvelle clé OpenAI API", openAiKey);
            AddRow(grid, 6, "", clearKey);
            AddRow(grid, 7, "", new Label { Text = "Clé vide : conserver la clé actuelle. Les secrets sont chiffrés pour ce compte Windows.",
                AutoSize = true, Dock = DockStyle.Fill });
            AddRow(grid, 8, "", new Label { Text = "", Dock = DockStyle.Fill });
            provider.SelectedIndexChanged += (sender, args) => UpdateRows();
            codexLogin.Click += (sender, args) => {
                try { CodexAccount.StartLogin(); codexStatus.Text = "Connexion ouverte. Cliquez sur Actualiser après authentification."; }
                catch (Exception ex) { codexStatus.Text = ex.Message; }
            };
            codexRefresh.Click += async (sender, args) => await RefreshCodexStatusAsync();
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.RightToLeft };
            var save = new Button { Text = "Enregistrer", Width = 105 };
            var cancel = new Button { Text = "Annuler", Width = 105, DialogResult = DialogResult.Cancel };
            save.Click += (sender, args) => Save();
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);
            Controls.Add(grid);
            Controls.Add(buttons);
            CancelButton = cancel;
            UpdateRows();
        }

        private void UpdateRows()
        {
            var selected = provider.SelectedItem as LlmProvider;
            if (selected == null) return;
            bool codex = selected.IsCodex;
            bool openAi = selected.Name == "OpenAI API";
            bool ollama = selected.Name == "Ollama";
            int[] heights = { 38, codex ? 38 : 0, codex ? 42 : 0, openAi ? 38 : 0,
                ollama ? 38 : 0, openAi ? 38 : 0, openAi ? 38 : 0, openAi ? 55 : 0, 0 };
            for (int i = 0; i < heights.Length; i++) grid.RowStyles[i].Height = heights[i];
            foreach (Control control in grid.Controls)
            {
                int row = grid.GetRow(control);
                control.Visible = heights[row] > 0;
            }
            if (codex) _ = RefreshCodexStatusAsync();
            Height = codex ? 260 : openAi ? 360 : ollama ? 260 : 230;
        }

        private async System.Threading.Tasks.Task RefreshCodexStatusAsync()
        {
            codexStatus.Text = "Vérification de la connexion ChatGPT…";
            try
            {
                string result = await CodexAccount.ReadStatusAsync();
                if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCodex) codexStatus.Text = result;
            }
            catch (Exception ex)
            {
                if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCodex) codexStatus.Text = ex.Message;
            }
        }

        private static void AddRow(TableLayoutPanel grid, int row, string label, Control control)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, row == 8 ? 55 : 38));
            grid.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill }, 0, row);
            grid.Controls.Add(control, 1, row);
        }

        private void Save()
        {
            try
            {
                ValidateEndpoint(openAiEndpoint.Text);
                ValidateEndpoint(ollamaEndpoint.Text);
                settings.ProviderName = ((LlmProvider)provider.SelectedItem).Name;
                settings.OpenAiEndpoint = openAiEndpoint.Text.Trim();
                settings.OllamaEndpoint = ollamaEndpoint.Text.Trim();
                if (clearKey.Checked) settings.EncryptedOpenAiKey = null;
                if (openAiKey.Text.Length > 0) settings.SetOpenAiKey(openAiKey.Text);
                settings.Save();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Configuration CodexVBE", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ValidateEndpoint(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            Uri endpoint;
            if (!Uri.TryCreate(raw, UriKind.Absolute, out endpoint) ||
                (endpoint.Scheme != Uri.UriSchemeHttps && !(endpoint.Scheme == Uri.UriSchemeHttp && endpoint.IsLoopback)))
                throw new ArgumentException("L'URL doit utiliser HTTPS, ou HTTP sur localhost.");
        }
    }
}
