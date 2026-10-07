using System;
using System.Drawing;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Fenêtre de configuration des fournisseurs, comptes et préférences de conversation.</summary>
    internal sealed partial class LlmSettingsWindow : Form
    {

        /// <summary>Paramètres persistants modifiés par cette fenêtre.</summary>
        private readonly LlmSettings settings;
        // Keep authentication, persistence and notices at replaceable native boundaries.
        /// <summary>Enregistre les paramètres avec le stockage natif.</summary>
        internal static Action<LlmSettings> WriteSettings = (Action<LlmSettings>)Delegate.CreateDelegate(typeof(Action<LlmSettings>), typeof(LlmSettings).GetMethod("Save"));

        /// <summary>Ouvre l’authentification native Copilot.</summary>
        internal static Action StartCopilotLogin = CopilotClient.StartLogin;

        /// <summary>Ouvre l’authentification native Codex.</summary>
        internal static Action StartCodexLogin = CodexAccount.StartLogin;

        /// <summary>Lit l’état de connexion du CLI Copilot.</summary>
        internal static Func<System.Threading.Tasks.Task<string>> ReadCopilotStatus = CopilotClient.ReadStatusAsync;

        /// <summary>Lit l’état du compte Codex.</summary>
        internal static Func<System.Threading.Tasks.Task<CodexAccountStatus>> ReadCodexStatus = CodexAccount.ReadStatusAsync;

        /// <summary>Applique le thème sélectionné à l’interface.</summary>
        internal static Action<ThemeChoice> SelectTheme = UiTheme.Select;

        /// <summary>Applique l’habillage sombre aux fenêtres natives du VBE.</summary>
        internal static Action<bool> SelectNativeVbeTheme = VbeNativeTheme.SetEnabled;

        /// <summary>Affiche un message natif appartenant à la fenêtre de configuration.</summary>
        internal static Func<IWin32Window, string, string, MessageBoxButtons, MessageBoxIcon, DialogResult> ShowNotice = MessageBox.Show;

        /// <summary>Empêche les recalculs imbriqués de hauteur de contenu.</summary>
        private bool fittingContent;

        /// <summary>Annulation des opérations d’authentification et de lecture GitHub.</summary>
        private readonly System.Threading.CancellationTokenSource githubCancellation = new System.Threading.CancellationTokenSource();

        /// <summary>Indique si une opération GitHub est en cours.</summary>
        private bool githubBusy;

        /// <summary>Indique si la première lecture GitHub a été déclenchée.</summary>
        private bool githubLoaded;

        /// <summary>Empêche la libération répétée des ressources appartenant à la fenêtre.</summary>
        private bool resourcesDisposed;

        /// <summary>Service Git Credential Manager utilisé pour l’authentification GitHub.</summary>
        private GitHubAccountService githubService = new GitHubAccountService();

        /// <summary>Fournisseur affiché dont les saisies sont actuellement éditées.</summary>
        private LlmProvider displayedProvider;

        /// <summary>Adresses d’API en cours de modification, indexées par fournisseur.</summary>
        private readonly System.Collections.Generic.Dictionary<string, string> endpointDrafts = new System.Collections.Generic.Dictionary<string, string>();

        /// <summary>Clés API en cours de modification, indexées par fournisseur.</summary>
        private readonly System.Collections.Generic.Dictionary<string, string> keyDrafts = new System.Collections.Generic.Dictionary<string, string>();

        /// <summary>Fournisseurs dont la clé enregistrée doit être supprimée.</summary>
        private readonly System.Collections.Generic.HashSet<string> clearedKeys = new System.Collections.Generic.HashSet<string>();

        /// <summary>Identifiants de modèles en cours de modification, indexés par fournisseur.</summary>
        private readonly System.Collections.Generic.Dictionary<string, string> modelDrafts = new System.Collections.Generic.Dictionary<string, string>();

        /// <summary>Ollama generation drafts survive provider changes without altering saved settings.</summary>
        private string ollamaTemperatureDraft, ollamaTopPDraft;

        /// <summary>Ajuste la hauteur du contenu et déclenche le premier chargement GitHub.</summary>
        /// <param name="e">Données de l’événement WinForms.</param>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (settings != null && !githubLoaded) { githubLoaded = true; _ = RefreshGitHubAsync(false); }
        }

        /// <summary>Mesure le contenu à la largeur courante et ajuste la hauteur à l’espace disponible.</summary>
        private void FitContentHeight()
        {
            if (fittingContent || IsDisposed) return;
            fittingContent = true;
            try
            {
                // Measure at the current width so wrapped descriptions and authentication
                // buttons contribute their actual height, including the current DPI/font.
                int width = ClientSize.Width;
                int contentHeight = Math.Max(providerSettingsView.GetPreferredSize(new Size(width, 0)).Height,
                    Math.Max(gitHubAccountSettingsView.GetPreferredSize(new Size(width, 0)).Height,
                        appearanceSettingsView.GetPreferredSize(new Size(width, 0)).Height));
                int height = contentHeight + settingsTabs.ItemSize.Height + settingsTabs.Padding.Y * 2 +
                    buttons.GetPreferredSize(new Size(width, 0)).Height + contentLayout.Padding.Vertical;
                int frameHeight = Height - ClientSize.Height;
                int available = Screen.FromControl(this).WorkingArea.Height - frameHeight;
                MinimumSize = new Size(MinimumSize.Width, 0);
                ClientSize = new Size(width, Math.Min(height, available));
                MinimumSize = new Size(MinimumSize.Width, Height);
                PerformLayout();
            }
            finally { fittingContent = false; }
        }

        /// <summary>Crée et initialise la fenêtre sans paramètres persistants.</summary>
        public LlmSettingsWindow()
        {
            InitializeComponent();
            BindViews();
            Icon = VbeWindowIcons.Icon("settings");
            UiText.Apply(this, null, githubToolTips);
            InitializeTheme();
        }

        /// <summary>Crée la fenêtre et initialise ses contrôles avec les paramètres fournis.</summary>
        /// <param name="settings">Paramètres persistants du fournisseur à éditer.</param>
        public LlmSettingsWindow(LlmSettings settings)
        {
            this.settings = settings;
            InitializeComponent();
            BindViews();
            Icon = VbeWindowIcons.Icon("settings");
            UiText.Apply(this, null, githubToolTips);
            InitializeTheme();
            githubAccount.Items.Add(UiText.Get("Automatic Git selection"));
            githubAccount.SelectedIndex = 0;
            customName.Text = settings.CustomProviderName ?? "";
            if (!string.IsNullOrEmpty(settings.GitHubAccount)) { githubAccount.Items.Add(settings.GitHubAccount); githubAccount.SelectedItem = settings.GitHubAccount; }
            azureEntra.Checked = settings.AzureUseEntraToken;
            ollamaTemperatureDraft = settings.OllamaTemperature?.ToString("R", System.Globalization.CultureInfo.InvariantCulture) ?? "";
            ollamaTopPDraft = settings.OllamaTopP?.ToString("R", System.Globalization.CultureInfo.InvariantCulture) ?? "";
            nativeVbeDark.Checked = settings.NativeVbeDarkTheme;
            provider.Items.AddRange(LlmProvider.All);
            approvalPicker.Items.AddRange(new object[] { UiText.Get("Automatic"), UiText.Get("Ask for other actions"), UiText.Get("Read-only") });
            approvalPicker.SelectedIndex = settings.VbeEditApproval == "ReadOnly" ? 2 :
                settings.VbeEditApproval == "AskEachTime" ? 1 : 0;
            int current = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            provider.SelectedIndex = current < 0 ? 0 : current;
            provider.SelectedIndexChanged += (sender, args) => UpdateRows();
            codexLogin.Click += (sender, args) =>
            {
                try { if (((LlmProvider)provider.SelectedItem).IsCopilot) StartCopilotLogin(); else StartCodexLogin(); codexStatus.Text = UiText.Get("Sign-in opened. Click Refresh after authenticating."); }
                catch (Exception ex) { codexStatus.Text = ex.Message; }
            };
            codexRefresh.Click += async (sender, args) => await RefreshCodexStatusAsync();
            saveButton.Click += (sender, args) => Save();
            UpdateRows();
        }

        /// <summary>Initialise le thème et son unique gestionnaire hors du concepteur Visual Studio.</summary>
        private void InitializeTheme()
        {
            themePicker.SelectedIndex = (int)UiTheme.Choice;
            if (System.ComponentModel.LicenseManager.UsageMode != System.ComponentModel.LicenseUsageMode.Designtime)
                themePicker.SelectedIndexChanged += ThemePicker_SelectedIndexChanged;
        }

        /// <summary>Applique le thème choisi et affiche les erreurs de sélection.</summary>
        /// <param name="sender">Sélecteur de thème.</param>
        /// <param name="e">Événement de sélection.</param>
        private void ThemePicker_SelectedIndexChanged(object sender, EventArgs e)
        {
            try { if (themePicker.SelectedIndex >= 0) SelectTheme((ThemeChoice)themePicker.SelectedIndex); }
            catch (Exception ex) { ShowNotice(this, ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.None); }
        }

        /// <summary>Met à jour valeurs, libellés et visibilité selon le fournisseur sélectionné.</summary>
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
            ollamaTemperature.Text = ollamaTemperatureDraft;
            ollamaTopP.Text = ollamaTopPDraft;
            bool codex = selected.IsCodex;
            bool cli = codex || selected.IsCopilot;
            bool[] visible = { true, cli, cli, !cli && !selected.Local, !cli && selected.Local, !cli, !cli, !cli, true, selected.ManualModels, selected.IsCustom, selected.IsAzure, selected.IsOllama, selected.IsOllama };
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
            foreach (Control control in grid.Controls)
            {
                int row = grid.GetRow(control);
                control.Visible = row >= visible.Length || visible[row];
            }
            grid.ResumeLayout(true);
            FitContentHeight();
            if (cli) _ = RefreshCodexStatusAsync();
        }

        /// <summary>Conserve les saisies non enregistrées du fournisseur affiché.</summary>
        private void CaptureDraft()
        {
            if (displayedProvider == null || displayedProvider.IsCodex || displayedProvider.IsCopilot) return;
            endpointDrafts[displayedProvider.Name] = (displayedProvider.Local ? ollamaEndpoint.Text : openAiEndpoint.Text).Trim();
            keyDrafts[displayedProvider.Name] = openAiKey.Text.Trim();
            if (displayedProvider.ManualModels) modelDrafts[displayedProvider.Name] = manualModels.Text.Trim();
            if (displayedProvider.IsOllama)
            {
                ollamaTemperatureDraft = ollamaTemperature.Text.Trim();
                ollamaTopPDraft = ollamaTopP.Text.Trim();
            }
            if (clearKey.Checked) clearedKeys.Add(displayedProvider.Name); else clearedKeys.Remove(displayedProvider.Name);
        }

        /// <summary>Démarre l’authentification GitHub puis actualise les comptes.</summary>
        /// <param name="e">Données de l’événement WinForms.</param>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        private async void GitHubLogin_Click(object sender, EventArgs e) { await RefreshGitHubAsync(true); }

        /// <summary>Actualise la liste des comptes GitHub sauvegardés.</summary>
        /// <param name="e">Données de l’événement WinForms.</param>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        private async void GitHubRefresh_Click(object sender, EventArgs e) { await RefreshGitHubAsync(false); }

        /// <summary>Authentifie si demandé puis charge les comptes GitHub en préservant la sélection.</summary>
        /// <param name="login">Indique si la lecture doit être précédée d’une authentification interactive.</param>
        /// <returns>Tâche terminée après la lecture du compte et la mise à jour de l’interface.</returns>
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
                    (login ? UiText.Get("GitHub sign-in completed. ") : "") + accounts.Length +
                    UiText.Get(" account(s) available through Git Credential Manager. Repository access will be checked during synchronization.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { if (!IsDisposed) githubStatus.Text = ex.Message; }
            finally
            {
                githubBusy = false;
                if (!IsDisposed)
                {
                    githubLogin.Enabled = githubRefresh.Enabled = githubAccount.Enabled = saveButton.Enabled = true;
                }
            }
        }

        /// <summary>Annule les requêtes et libère les ressources détenues par la fenêtre.</summary>
        /// <param name="disposing">Indique si les ressources gérées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && !resourcesDisposed)
            {
                resourcesDisposed = true;
                githubCancellation.Cancel(); githubCancellation.Dispose(); githubToolTips?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>Lit l’état du compte Copilot ou Codex correspondant au fournisseur courant.</summary>
        /// <returns>Tâche terminée après la lecture de l’état du compte courant.</returns>
        private async System.Threading.Tasks.Task RefreshCodexStatusAsync()
        {
            if (((LlmProvider)provider.SelectedItem).IsCopilot)
            {
                codexStatus.Text = UiText.Get("Checking GitHub Copilot…");
                try { string status = await ReadCopilotStatus(); if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCopilot) codexStatus.Text = status; }
                catch (Exception ex) { if (!IsDisposed && ((LlmProvider)provider.SelectedItem).IsCopilot) codexStatus.Text = ex.Message; }
                return;
            }
            codexStatus.Text = UiText.Get("Checking ChatGPT connection…");
            try
            {
                var result = await ReadCodexStatus();
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

        /// <summary>Valide puis enregistre les réglages et ferme la fenêtre en cas de succès.</summary>
        private void Save()
        {
            double? previousOllamaTemperature = settings?.OllamaTemperature;
            double? previousOllamaTopP = settings?.OllamaTopP;
            try
            {
                CaptureDraft();
                foreach (var endpoint in endpointDrafts.Values) ValidateEndpoint(endpoint);
                double? temperature = ParseOllamaSampling(ollamaTemperatureDraft, true);
                double? topP = ParseOllamaSampling(ollamaTopPDraft, false);
                settings.OllamaTemperature = temperature;
                settings.OllamaTopP = topP;
                settings.ProviderName = ((LlmProvider)provider.SelectedItem).Name;
                settings.GitHubAccount = githubAccount.SelectedIndex > 0 ? Convert.ToString(githubAccount.SelectedItem) : null;
                settings.CustomProviderName = customName.Text.Trim();
                settings.AzureUseEntraToken = azureEntra.Checked;
                if (settings.ManualModelLists == null) settings.ManualModelLists = new System.Collections.Generic.Dictionary<string, string>();
                foreach (var pair in modelDrafts) settings.ManualModelLists[pair.Key] = pair.Value;
                settings.VbeEditApproval = approvalPicker.SelectedIndex == 2 ? "ReadOnly" :
                    approvalPicker.SelectedIndex == 1 ? "AskEachTime" : "Automatic";
                bool previousNativeTheme = settings.NativeVbeDarkTheme;
                settings.NativeVbeDarkTheme = nativeVbeDark.Checked;
                foreach (var item in LlmProvider.All)
                {
                    string value;
                    if (endpointDrafts.TryGetValue(item.Name, out value)) settings.SetEndpoint(item, value);
                    if (clearedKeys.Contains(item.Name)) settings.SetKey(item, null);
                    if (keyDrafts.TryGetValue(item.Name, out value) && !string.IsNullOrWhiteSpace(value)) settings.SetKey(item, value);
                }
                try
                {
                    SelectNativeVbeTheme(settings.NativeVbeDarkTheme);
                    WriteSettings(settings);
                }
                catch
                {
                    settings.NativeVbeDarkTheme = previousNativeTheme;
                    try { SelectNativeVbeTheme(previousNativeTheme); } catch { }
                    throw;
                }
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                if (settings != null)
                {
                    settings.OllamaTemperature = previousOllamaTemperature;
                    settings.OllamaTopP = previousOllamaTopP;
                }
                ShowNotice(this, ex.Message, UiText.Get("VBAi settings"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>Parses optional finite sampling values using invariant or current decimal notation.</summary>
        /// <param name="raw">Blank selects the server default.</param>
        /// <param name="temperature">Selects temperature bounds instead of top-p bounds.</param>
        /// <returns>A valid sampling override or null.</returns>
        private static double? ParseOllamaSampling(string raw, bool temperature)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            double value;
            if ((!double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value) &&
                 !double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out value)) ||
                double.IsNaN(value) || double.IsInfinity(value) || value > (temperature ? 2 : 1) ||
                (temperature ? value < 0 : value <= 0))
                throw new ArgumentException(UiText.Get(temperature ? "Ollama temperature must be a finite number from 0 to 2, or blank for the server default." :
                    "Ollama top-p must be a finite number above 0 and at most 1, or blank for the server default."));
            return value;
        }

        /// <summary>Accepte HTTPS et HTTP uniquement pour une adresse locale.</summary>
        /// <param name="raw">URL de point de terminaison à valider.</param>
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
