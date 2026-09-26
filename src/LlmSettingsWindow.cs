using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow : Form
    {
        private readonly LlmSettings settings;

        public LlmSettingsWindow()
        {
            InitializeComponent();
        }

        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            InitializeComponent();
            provider.Items.AddRange(LlmProvider.All);
            approvalPicker.Items.AddRange(new object[] { "Automatique", "Demander à chaque action", "Lecture seule" });
            approvalPicker.SelectedIndex = settings.VbeEditApproval == "ReadOnly" ? 2 :
                settings.VbeEditApproval == "AskEachTime" ? 1 : 0;
            int current = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            provider.SelectedIndex = current < 0 ? 0 : current;
            openAiEndpoint.Text = settings.OpenAiEndpoint ?? "";
            ollamaEndpoint.Text = settings.OllamaEndpoint ?? "";
            provider.SelectedIndexChanged += (sender, args) => UpdateRows();
            codexLogin.Click += (sender, args) => {
                try { CodexAccount.StartLogin(); codexStatus.Text = "Connexion ouverte. Cliquez sur Actualiser après authentification."; }
                catch (Exception ex) { codexStatus.Text = ex.Message; }
            };
            codexRefresh.Click += async (sender, args) => await RefreshCodexStatusAsync();
            saveButton.Click += (sender, args) => Save();
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
                ollama ? 38 : 0, openAi ? 38 : 0, openAi ? 38 : 0, openAi ? 55 : 0, 38 };
            for (int i = 0; i < heights.Length; i++) grid.RowStyles[i].Height = heights[i];
            foreach (Control control in grid.Controls)
            {
                int row = grid.GetRow(control);
                control.Visible = heights[row] > 0;
            }
            if (codex) _ = RefreshCodexStatusAsync();
            Height = codex ? 300 : openAi ? 400 : ollama ? 300 : 270;
        }

        private async System.Threading.Tasks.Task RefreshCodexStatusAsync()
        {
            codexStatus.Text = "Vérification de la connexion ChatGPT…";
            try
            {
                var result = await CodexAccount.ReadStatusAsync();
                if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCodex)
                {
                    codexStatus.Text = result.Text;
                    codexLogin.Enabled = !result.ChatGptConnected;
                }
            }
            catch (Exception ex)
            {
                if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCodex)
                {
                    codexStatus.Text = ex.Message;
                    codexLogin.Enabled = true;
                }
            }
        }

        private void Save()
        {
            try
            {
                ValidateEndpoint(openAiEndpoint.Text);
                ValidateEndpoint(ollamaEndpoint.Text);
                settings.ProviderName = ((LlmProvider)provider.SelectedItem).Name;
                settings.VbeEditApproval = approvalPicker.SelectedIndex == 2 ? "ReadOnly" :
                    approvalPicker.SelectedIndex == 1 ? "AskEachTime" : "Automatic";
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
