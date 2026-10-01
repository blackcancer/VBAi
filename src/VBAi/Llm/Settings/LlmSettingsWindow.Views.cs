namespace VBAi
{
    /// <summary>Références de commodité vers les contrôles des onglets de paramètres.</summary>
    internal sealed partial class LlmSettingsWindow
    {
        /// <summary>Grille de configuration des fournisseurs.</summary>
        private System.Windows.Forms.TableLayoutPanel grid;
        /// <summary>Actions de connexion et d’actualisation Codex.</summary>
        private System.Windows.Forms.FlowLayoutPanel codexActions;
        /// <summary>Sélecteur du fournisseur LLM.</summary>
        private System.Windows.Forms.ComboBox provider;
        /// <summary>Résultat de la vérification du compte Codex.</summary>
        private System.Windows.Forms.Label codexStatus;
        /// <summary>Action de connexion du fournisseur CLI sélectionné.</summary>
        private System.Windows.Forms.Button codexLogin;
        /// <summary>Action de relecture de l’état de connexion CLI.</summary>
        private System.Windows.Forms.Button codexRefresh;
        /// <summary>Champ du point de terminaison OpenAI.</summary>
        private System.Windows.Forms.TextBox openAiEndpoint;
        /// <summary>Champ du point de terminaison Ollama.</summary>
        private System.Windows.Forms.TextBox ollamaEndpoint;
        /// <summary>Champ de saisie de clé fournisseur.</summary>
        private System.Windows.Forms.TextBox openAiKey;
        /// <summary>Option de suppression de la clé enregistrée.</summary>
        private System.Windows.Forms.CheckBox clearKey;
        /// <summary>Libellé du sélecteur de fournisseur.</summary>
        private System.Windows.Forms.Label providerLabel;
        /// <summary>Libellé du compte de service.</summary>
        private System.Windows.Forms.Label accountLabel;
        /// <summary>Libellé de la section d’authentification.</summary>
        private System.Windows.Forms.Label authenticationLabel;
        /// <summary>Libellé de l’adresse OpenAI.</summary>
        private System.Windows.Forms.Label openAiEndpointLabel;
        /// <summary>Libellé de l’adresse Ollama.</summary>
        private System.Windows.Forms.Label ollamaEndpointLabel;
        /// <summary>Libellé du champ de clé.</summary>
        private System.Windows.Forms.Label keyLabel;
        /// <summary>Note explicative sur la clé.</summary>
        private System.Windows.Forms.Label keyNote;
        /// <summary>Libellé du réglage d’approbation VBE.</summary>
        private System.Windows.Forms.Label approvalLabel;
        /// <summary>Sélecteur de politique d’approbation.</summary>
        private System.Windows.Forms.ComboBox approvalPicker;
        /// <summary>Libellé de la liste de modèles saisie manuellement.</summary>
        private System.Windows.Forms.Label manualModelsLabel;
        /// <summary>Champ des identifiants de modèles saisis par l’utilisateur.</summary>
        private System.Windows.Forms.TextBox manualModels;
        /// <summary>Libellé du nom du fournisseur personnalisé.</summary>
        private System.Windows.Forms.Label customNameLabel;
        /// <summary>Champ du nom du fournisseur personnalisé.</summary>
        private System.Windows.Forms.TextBox customName;
        /// <summary>Option d’authentification Azure par jeton Entra.</summary>
        private System.Windows.Forms.CheckBox azureEntra;
        /// <summary>Optional Ollama temperature entry.</summary>
        private System.Windows.Forms.TextBox ollamaTemperature;
        /// <summary>Optional Ollama top-p entry.</summary>
        private System.Windows.Forms.TextBox ollamaTopP;
        /// <summary>Libellé de la section du compte GitHub.</summary>
        private System.Windows.Forms.Label githubLabel;
        /// <summary>État de la connexion GitHub.</summary>
        private System.Windows.Forms.Label githubStatus;
        /// <summary>Libellé du compte GitHub actif.</summary>
        private System.Windows.Forms.Label githubAccountLabel;
        /// <summary>Note explicative sur l’accès GitHub.</summary>
        private System.Windows.Forms.Label githubNote;
        /// <summary>Sélecteur du compte GitHub.</summary>
        private System.Windows.Forms.ComboBox githubAccount;
        /// <summary>Actions de connexion GitHub.</summary>
        private System.Windows.Forms.FlowLayoutPanel githubActions;
        /// <summary>Action de connexion GitHub.</summary>
        private System.Windows.Forms.Button githubLogin;
        /// <summary>Action de relecture du compte GitHub.</summary>
        private System.Windows.Forms.Button githubRefresh;
        /// <summary>Info-bulles utilisées par les contrôles GitHub.</summary>
        private System.Windows.Forms.ToolTip githubToolTips;
        /// <summary>Panneau des options d’apparence.</summary>
        private System.Windows.Forms.FlowLayoutPanel themePanel;
        /// <summary>Libellé du thème de l’interface.</summary>
        private System.Windows.Forms.Label themeLabel;
        /// <summary>Sélecteur du thème de l’interface.</summary>
        private System.Windows.Forms.ComboBox themePicker;
        /// <summary>Option d’habillage sombre des fenêtres natives du VBE.</summary>
        private System.Windows.Forms.CheckBox nativeVbeDark;
        /// <summary>Note de portée de l’habillage sombre natif.</summary>
        private System.Windows.Forms.Label nativeVbeDarkNote;
        /// <summary>Lie les champs de compatibilité aux contrôles de leurs vues et branche les actions GitHub.</summary>
        private void BindViews()
        {
            grid = providerSettingsView.grid;
            codexActions = providerSettingsView.codexActions;
            provider = providerSettingsView.provider;
            codexStatus = providerSettingsView.codexStatus;
            codexLogin = providerSettingsView.codexLogin;
            codexRefresh = providerSettingsView.codexRefresh;
            openAiEndpoint = providerSettingsView.openAiEndpoint;
            ollamaEndpoint = providerSettingsView.ollamaEndpoint;
            openAiKey = providerSettingsView.openAiKey;
            clearKey = providerSettingsView.clearKey;
            providerLabel = providerSettingsView.providerLabel;
            accountLabel = providerSettingsView.accountLabel;
            authenticationLabel = providerSettingsView.authenticationLabel;
            openAiEndpointLabel = providerSettingsView.openAiEndpointLabel;
            ollamaEndpointLabel = providerSettingsView.ollamaEndpointLabel;
            keyLabel = providerSettingsView.keyLabel;
            keyNote = providerSettingsView.keyNote;
            approvalLabel = providerSettingsView.approvalLabel;
            approvalPicker = providerSettingsView.approvalPicker;
            manualModelsLabel = providerSettingsView.manualModelsLabel;
            manualModels = providerSettingsView.manualModels;
            customNameLabel = providerSettingsView.customNameLabel;
            customName = providerSettingsView.customName;
            azureEntra = providerSettingsView.azureEntra;
            ollamaTemperature = providerSettingsView.ollamaTemperature;
            ollamaTopP = providerSettingsView.ollamaTopP;
            githubLabel = gitHubAccountSettingsView.githubLabel;
            githubStatus = gitHubAccountSettingsView.githubStatus;
            githubAccountLabel = gitHubAccountSettingsView.githubAccountLabel;
            githubNote = gitHubAccountSettingsView.githubNote;
            githubAccount = gitHubAccountSettingsView.githubAccount;
            githubActions = gitHubAccountSettingsView.githubActions;
            githubLogin = gitHubAccountSettingsView.githubLogin;
            githubRefresh = gitHubAccountSettingsView.githubRefresh;
            githubToolTips = gitHubAccountSettingsView.githubToolTips;
            themePanel = appearanceSettingsView.themePanel;
            themeLabel = appearanceSettingsView.themeLabel;
            themePicker = appearanceSettingsView.themePicker;
            nativeVbeDark = appearanceSettingsView.nativeVbeDark;
            nativeVbeDarkNote = appearanceSettingsView.nativeVbeDarkNote;
            this.githubLogin.Click += new System.EventHandler(this.GitHubLogin_Click);
            this.githubRefresh.Click += new System.EventHandler(this.GitHubRefresh_Click);
        }
    }
}
