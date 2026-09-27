using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow : Form
    {
        private readonly LlmSettings settings;
        private bool fittingContent;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            FitContentHeight();
        }

        private void FitContentHeight()
        {
            if (fittingContent || IsDisposed) return;
            fittingContent = true;
            try
            {
                // Measure at the current width so wrapped descriptions and authentication
                // buttons contribute their actual height, including the current DPI/font.
                int width = ClientSize.Width;
                int height = grid.GetPreferredSize(new Size(width, 0)).Height +
                    buttons.GetPreferredSize(new Size(width, 0)).Height;
                int frameHeight = Height - ClientSize.Height;
                int available = Screen.FromControl(this).WorkingArea.Height - frameHeight;
                MinimumSize = new Size(MinimumSize.Width, 0);
                ClientSize = new Size(width, Math.Min(height, available));
                MinimumSize = new Size(MinimumSize.Width, Height);
                PerformLayout();
            }
            finally { fittingContent = false; }
        }

        public LlmSettingsWindow()
        {
            InitializeComponent();
            ApplyLayoutTuning();
        }

        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            InitializeComponent();
            ApplyLayoutTuning();
            provider.Items.AddRange(LlmProvider.All);
            approvalPicker.Items.AddRange(new object[] { "Automatique", "Demander pour les autres actions", "Lecture seule" });
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

        private void ApplyLayoutTuning()
        {
            foreach (Control control in grid.Controls)
            {
                control.Margin = new Padding(0, 4, 12, 8);
                var label = control as Label;
                if (label != null) label.AutoSize = true;
                if (label != null && grid.GetColumn(control) == 0)
                    label.MinimumSize = new Size(156, 0);
                var check = control as CheckBox;
                if (check != null) check.AutoSize = true;
            }
            foreach (Button button in new[] { codexLogin, codexRefresh, saveButton, cancelButton })
            {
                button.AutoSize = true;
                button.MinimumSize = new Size(button.Width, 30);
                button.Padding = new Padding(6, 2, 6, 2);
            }
        }

        private void UpdateRows()
        {
            var selected = provider.SelectedItem as LlmProvider;
            if (selected == null) return;
            bool codex = selected.IsCodex;
            bool openAi = selected.Name == "OpenAI API";
            bool ollama = selected.Name == "Ollama";
            bool[] visible = { true, codex, codex, openAi, ollama, openAi, openAi, openAi, true };
            grid.SuspendLayout();
            for (int i = 0; i < visible.Length; i++)
            {
                grid.RowStyles[i].SizeType = visible[i] ? SizeType.AutoSize : SizeType.Absolute;
                grid.RowStyles[i].Height = 0;
            }
            foreach (Control control in grid.Controls)
            {
                int row = grid.GetRow(control);
                control.Visible = visible[row];
            }
            grid.ResumeLayout(true);
            if (Visible) FitContentHeight();
            if (codex) _ = RefreshCodexStatusAsync();
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
            if (!IsDisposed && Visible) FitContentHeight();
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
