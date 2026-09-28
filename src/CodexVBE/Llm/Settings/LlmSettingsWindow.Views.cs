namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow
    {
        private System.Windows.Forms.TableLayoutPanel grid;
        private System.Windows.Forms.FlowLayoutPanel codexActions;
        private System.Windows.Forms.ComboBox provider;
        private System.Windows.Forms.Label codexStatus;
        private System.Windows.Forms.Button codexLogin;
        private System.Windows.Forms.Button codexRefresh;
        private System.Windows.Forms.TextBox openAiEndpoint;
        private System.Windows.Forms.TextBox ollamaEndpoint;
        private System.Windows.Forms.TextBox openAiKey;
        private System.Windows.Forms.CheckBox clearKey;
        private System.Windows.Forms.Label providerLabel;
        private System.Windows.Forms.Label accountLabel;
        private System.Windows.Forms.Label authenticationLabel;
        private System.Windows.Forms.Label openAiEndpointLabel;
        private System.Windows.Forms.Label ollamaEndpointLabel;
        private System.Windows.Forms.Label keyLabel;
        private System.Windows.Forms.Label keyNote;
        private System.Windows.Forms.Label approvalLabel;
        private System.Windows.Forms.ComboBox approvalPicker;
        private System.Windows.Forms.Label manualModelsLabel;
        private System.Windows.Forms.TextBox manualModels;
        private System.Windows.Forms.Label customNameLabel;
        private System.Windows.Forms.TextBox customName;
        private System.Windows.Forms.CheckBox azureEntra;
        private System.Windows.Forms.Label githubLabel;
        private System.Windows.Forms.Label githubStatus;
        private System.Windows.Forms.Label githubAccountLabel;
        private System.Windows.Forms.Label githubNote;
        private System.Windows.Forms.ComboBox githubAccount;
        private System.Windows.Forms.FlowLayoutPanel githubActions;
        private System.Windows.Forms.Button githubLogin;
        private System.Windows.Forms.Button githubRefresh;
        private System.Windows.Forms.ToolTip githubToolTips;
        private System.Windows.Forms.FlowLayoutPanel themePanel;
        private System.Windows.Forms.Label themeLabel;
        private System.Windows.Forms.ComboBox themePicker;
        private System.Windows.Forms.CheckBox nativeVbeDark;
        private System.Windows.Forms.Label nativeVbeDarkNote;
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
