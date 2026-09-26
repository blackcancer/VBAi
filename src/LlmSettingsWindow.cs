using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class LlmSettingsWindow : Form
    {
        private readonly ComboBox provider = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        private readonly TextBox codexModel = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox openAiModel = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox ollamaModel = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox openAiEndpoint = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox ollamaEndpoint = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox openAiKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        private readonly CheckBox clearKey = new CheckBox { Text = "Supprimer la clé API enregistrée", Dock = DockStyle.Fill };
        private readonly LlmSettings settings;

        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            Text = "CodexVBE — Configuration LLM";
            Width = 650;
            Height = 460;
            MinimumSize = new Size(560, 420);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9F);
            provider.Items.AddRange(LlmProvider.All);
            int current = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            provider.SelectedIndex = current < 0 ? 0 : current;
            codexModel.Text = settings.CodexModel ?? "";
            openAiModel.Text = settings.OpenAiModel ?? "";
            ollamaModel.Text = settings.OllamaModel ?? "";
            openAiEndpoint.Text = settings.OpenAiEndpoint ?? "";
            ollamaEndpoint.Text = settings.OllamaEndpoint ?? "";

            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 9,
                Padding = new Padding(12), AutoScroll = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            AddRow(grid, 0, "Fournisseur", provider);
            AddRow(grid, 1, "Modèle Codex (facultatif)", codexModel);
            AddRow(grid, 2, "Modèle OpenAI API", openAiModel);
            AddRow(grid, 3, "Modèle Ollama", ollamaModel);
            AddRow(grid, 4, "URL OpenAI (facultatif)", openAiEndpoint);
            AddRow(grid, 5, "URL Ollama (facultatif)", ollamaEndpoint);
            AddRow(grid, 6, "Nouvelle clé OpenAI API", openAiKey);
            AddRow(grid, 7, "", clearKey);
            AddRow(grid, 8, "", new Label { Text = "Clé vide : conserver la clé actuelle. Les secrets sont chiffrés pour ce compte Windows.",
                AutoSize = true, Dock = DockStyle.Fill });
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.RightToLeft };
            var save = new Button { Text = "Enregistrer", Width = 105 };
            var cancel = new Button { Text = "Annuler", Width = 105, DialogResult = DialogResult.Cancel };
            save.Click += (sender, args) => Save();
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);
            Controls.Add(grid);
            Controls.Add(buttons);
            CancelButton = cancel;
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
                settings.CodexModel = codexModel.Text.Trim();
                settings.OpenAiModel = openAiModel.Text.Trim();
                settings.OllamaModel = ollamaModel.Text.Trim();
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
