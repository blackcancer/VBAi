using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow : Form
    {
        private readonly LlmSettings settings;
        private bool fittingContent;
        private readonly System.Threading.CancellationTokenSource githubCancellation = new System.Threading.CancellationTokenSource();
        private bool githubBusy;
        private bool githubLoaded;
        private bool resourcesDisposed;
        private GitHubAccountService githubService = new GitHubAccountService();
        private LlmProvider displayedProvider;
        private readonly System.Collections.Generic.Dictionary<string, string> endpointDrafts = new System.Collections.Generic.Dictionary<string, string>();
        private readonly System.Collections.Generic.Dictionary<string, string> keyDrafts = new System.Collections.Generic.Dictionary<string, string>();
        private readonly System.Collections.Generic.HashSet<string> clearedKeys = new System.Collections.Generic.HashSet<string>();
        private readonly System.Collections.Generic.Dictionary<string, string> modelDrafts = new System.Collections.Generic.Dictionary<string, string>();

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            FitContentHeight();
            if (settings != null && !githubLoaded) { githubLoaded = true; _ = RefreshGitHubAsync(false); }
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
                    buttons.GetPreferredSize(new Size(width, 0)).Height + themePanel.GetPreferredSize(new Size(width, 0)).Height;
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
            Icon = VbeWindowIcons.Icon("settings");
            UiText.Apply(this, null, githubToolTips);
            themePicker.SelectedIndex = (int)UiTheme.Choice;
            themePicker.SelectedIndexChanged += (s, e) => { try { if (themePicker.SelectedIndex >= 0) UiTheme.Select((ThemeChoice)themePicker.SelectedIndex); } catch (Exception ex) { MessageBox.Show(this, ex.Message); } };
        }

        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            InitializeComponent();
            Icon = VbeWindowIcons.Icon("settings");
            UiText.Apply(this, null, githubToolTips);
            themePicker.SelectedIndex = (int)UiTheme.Choice;
            themePicker.SelectedIndexChanged += (s, e) => { try { if (themePicker.SelectedIndex >= 0) UiTheme.Select((ThemeChoice)themePicker.SelectedIndex); } catch (Exception ex) { MessageBox.Show(this, ex.Message); } };
            githubAccount.Items.Add(UiText.Get("Automatic Git selection"));
            githubAccount.SelectedIndex = 0;
            customName.Text = settings.CustomProviderName ?? "";
            if (!string.IsNullOrEmpty(settings.GitHubAccount)) { githubAccount.Items.Add(settings.GitHubAccount); githubAccount.SelectedItem = settings.GitHubAccount; }
            azureEntra.Checked = settings.AzureUseEntraToken;
            provider.Items.AddRange(LlmProvider.All);
            approvalPicker.Items.AddRange(new object[] { UiText.Get("Automatic"), UiText.Get("Ask for other actions"), UiText.Get("Read-only") });
            approvalPicker.SelectedIndex = settings.VbeEditApproval == "ReadOnly" ? 2 :
                settings.VbeEditApproval == "AskEachTime" ? 1 : 0;
            int current = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            provider.SelectedIndex = current < 0 ? 0 : current;
            provider.SelectedIndexChanged += (sender, args) => UpdateRows();
            codexLogin.Click += (sender, args) => {
                try { if (((LlmProvider)provider.SelectedItem).IsCopilot) CopilotClient.StartLogin(); else CodexAccount.StartLogin(); codexStatus.Text = UiText.Get("Sign-in opened. Click Refresh after authenticating."); }
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
            CaptureDraft();
            displayedProvider = selected;
            string draft;
            string endpoint = endpointDrafts.TryGetValue(selected.Name, out draft) ? draft : settings.ResolveEndpoint(selected) ?? "";
            openAiEndpoint.Text = endpoint; ollamaEndpoint.Text = endpoint;
            openAiKey.Text = keyDrafts.TryGetValue(selected.Name, out draft) ? draft : "";
            clearKey.Checked = clearedKeys.Contains(selected.Name);
            manualModels.Text = modelDrafts.TryGetValue(selected.Name, out draft) ? draft : settings.GetManualModels(selected) ?? "";
            bool codex = selected.IsCodex;
            bool cli = codex || selected.IsCopilot;
            bool[] visible = { true, cli, cli, !cli && !selected.Local, !cli && selected.Local, !cli, !cli, !cli, true, selected.ManualModels, selected.IsCustom, selected.IsAzure };
            openAiEndpointLabel.Text = ollamaEndpointLabel.Text = UiText.Get("API URL");
            if (selected.IsBedrock) openAiEndpointLabel.Text = UiText.Get("Bedrock Runtime URL");
            manualModelsLabel.Text = selected.IsAzure ? UiText.Get("Deployments (one per line)") : selected.IsBedrock ? UiText.Get("Models / profiles (one per line)") : UiText.Get("Models (one per line)");
            openAiEndpoint.AccessibleName = ollamaEndpoint.AccessibleName = UiText.Get("API URL for ") + selected.Name;
            keyLabel.Text = UiText.Get("New key for ") + selected.Name;
            openAiKey.AccessibleName = keyLabel.Text;
            keyNote.Text = UiText.Get("Leave the key blank to keep the current one. Encryption is tied to your Windows account. ") +
                (selected.Local ? UiText.Get("Optional key for this local server. ") : "") +
                (selected.KeyVariable == null ? "" : UiText.Get("Fallback: ") + selected.KeyVariable + UiText.Get(". Removing the saved key enables this fallback again."));
            if (selected.IsCustom) keyNote.Text += UiText.Get(" Optional key. Enter the full /chat/completions URL and model identifiers.");
            if (selected.IsAzure) keyNote.Text += UiText.Get(" URL: https://<resource>.openai.azure.com/openai/v1/chat/completions. Models are your deployment names. In Entra mode: AZURE_OPENAI_ENTRA_TOKEN.");
            if (selected.IsBedrock) keyNote.Text += UiText.Get(" Bedrock API key (Bearer), not an IAM secret key. URL: https://bedrock-runtime.<region>.amazonaws.com.");
            accountLabel.Text = selected.IsCopilot ? UiText.Get("GitHub account") : UiText.Get("ChatGPT account");
            codexLogin.Text = selected.IsCopilot ? UiText.Get("Sign in to GitHub") : UiText.Get("Sign in to ChatGPT");
            codexLogin.Enabled = true;
            grid.SuspendLayout();
            for (int i = 0; i < visible.Length; i++)
            {
                grid.RowStyles[i].SizeType = visible[i] ? SizeType.AutoSize : SizeType.Absolute;
                grid.RowStyles[i].Height = 0;
            }
            foreach (Control control in grid.Controls)
            {
                int row = grid.GetRow(control);
                control.Visible = row >= visible.Length || visible[row];
            }
            grid.ResumeLayout(true);
            if (Visible) FitContentHeight();
            if (cli) _ = RefreshCodexStatusAsync();
        }

        private void CaptureDraft()
        {
            if (displayedProvider == null || displayedProvider.IsCodex || displayedProvider.IsCopilot) return;
            endpointDrafts[displayedProvider.Name] = (displayedProvider.Local ? ollamaEndpoint.Text : openAiEndpoint.Text).Trim();
            keyDrafts[displayedProvider.Name] = openAiKey.Text.Trim();
            if (displayedProvider.ManualModels) modelDrafts[displayedProvider.Name] = manualModels.Text.Trim();
            if (clearKey.Checked) clearedKeys.Add(displayedProvider.Name); else clearedKeys.Remove(displayedProvider.Name);
        }

        private async void GitHubLogin_Click(object sender, EventArgs e) { await RefreshGitHubAsync(true); }
        private async void GitHubRefresh_Click(object sender, EventArgs e) { await RefreshGitHubAsync(false); }

        private async System.Threading.Tasks.Task RefreshGitHubAsync(bool login)
        {
            if (githubBusy || settings == null) return;
            githubBusy = true;
            githubLogin.Enabled = githubRefresh.Enabled = githubAccount.Enabled = saveButton.Enabled = false;
            githubStatus.Text = login ? UiText.Get("Complete GitHub sign-in in your browser…") : UiText.Get("Looking for saved GitHub accounts…");
            try
            {
                var service = githubService;
                string selected = githubAccount.SelectedIndex > 0 ? Convert.ToString(githubAccount.SelectedItem) : null;
                if (login) await service.LoginAsync(githubCancellation.Token);
                string[] accounts = await service.ListAsync(githubCancellation.Token);
                if (IsDisposed) return;
                githubAccount.Items.Clear(); githubAccount.Items.Add(UiText.Get("Automatic Git selection"));
                githubAccount.Items.AddRange(accounts);
                if (selected != null && !githubAccount.Items.Contains(selected)) githubAccount.Items.Add(selected);
                githubAccount.SelectedItem = selected ?? (login && accounts.Length == 1 ? accounts[0] : UiText.Get("Automatic Git selection"));
                bool missing = selected != null && Array.IndexOf(accounts, selected) < 0;
                githubStatus.Text = missing ? UiText.Get("The selected account is no longer saved. Sign in again or choose another account.") :
                    accounts.Length == 0 ? UiText.Get("No saved GitHub account. Click Sign in.") :
                    accounts.Length + UiText.Get(" account(s) available through Git Credential Manager. Repository access will be checked during synchronization.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { if (!IsDisposed) githubStatus.Text = ex.Message; }
            finally
            {
                githubBusy = false;
                if (!IsDisposed)
                {
                    githubLogin.Enabled = githubRefresh.Enabled = githubAccount.Enabled = saveButton.Enabled = true;
                    if (Visible) FitContentHeight();
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !resourcesDisposed)
            {
                resourcesDisposed = true;
                githubCancellation.Cancel(); githubCancellation.Dispose(); githubToolTips?.Dispose();
            }
            base.Dispose(disposing);
        }

        private async System.Threading.Tasks.Task RefreshCodexStatusAsync()
        {
            if (((LlmProvider)provider.SelectedItem).IsCopilot) {
                codexStatus.Text = UiText.Get("Checking GitHub Copilot…");
                try { string status = await CopilotClient.ReadStatusAsync(); if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCopilot) codexStatus.Text = status; }
                catch (Exception ex) { if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCopilot) codexStatus.Text = ex.Message; }
                if (!IsDisposed && Visible) FitContentHeight();
                return;
            }
            codexStatus.Text = UiText.Get("Checking ChatGPT connection…");
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
                CaptureDraft();
                foreach (var endpoint in endpointDrafts.Values) ValidateEndpoint(endpoint);
                settings.ProviderName = ((LlmProvider)provider.SelectedItem).Name;
                settings.GitHubAccount = githubAccount.SelectedIndex > 0 ? Convert.ToString(githubAccount.SelectedItem) : null;
                settings.CustomProviderName = customName.Text.Trim();
                settings.AzureUseEntraToken = azureEntra.Checked;
                if (settings.ManualModelLists == null) settings.ManualModelLists = new System.Collections.Generic.Dictionary<string, string>();
                foreach (var pair in modelDrafts) settings.ManualModelLists[pair.Key] = pair.Value;
                settings.VbeEditApproval = approvalPicker.SelectedIndex == 2 ? "ReadOnly" :
                    approvalPicker.SelectedIndex == 1 ? "AskEachTime" : "Automatic";
                foreach (var item in LlmProvider.All) {
                    string value;
                    if (endpointDrafts.TryGetValue(item.Name, out value)) settings.SetEndpoint(item, value);
                    if (clearedKeys.Contains(item.Name)) settings.SetKey(item, null);
                    if (keyDrafts.TryGetValue(item.Name, out value) && !string.IsNullOrWhiteSpace(value)) settings.SetKey(item, value);
                }
                settings.Save();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, UiText.Get("VBAi settings"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ValidateEndpoint(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            Uri endpoint;
            if (!Uri.TryCreate(raw, UriKind.Absolute, out endpoint) ||
                (endpoint.Scheme != Uri.UriSchemeHttps && !(endpoint.Scheme == Uri.UriSchemeHttp && endpoint.IsLoopback)))
                throw new ArgumentException(UiText.Get("The URL must use HTTPS, or HTTP on localhost."));
        }
    }
}
