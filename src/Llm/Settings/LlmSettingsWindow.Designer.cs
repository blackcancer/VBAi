using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow
    {
        private TableLayoutPanel grid;
        private Label githubLabel;
        private Label githubStatus;
        private Label githubAccountLabel;
        private Label githubNote;
        private ComboBox githubAccount;
        private FlowLayoutPanel githubActions;
        private Button githubLogin;
        private Button githubRefresh;
        private ToolTip githubToolTips;
        private TableLayoutPanel contentLayout;
        private FlowLayoutPanel codexActions;
        private FlowLayoutPanel buttons;
        private ComboBox provider;
        private Label codexStatus;
        private Button codexLogin;
        private Button codexRefresh;
        private TextBox openAiEndpoint;
        private TextBox ollamaEndpoint;
        private TextBox openAiKey;
        private CheckBox clearKey;
        private Button saveButton;
        private Button cancelButton;
        private Label providerLabel;
        private Label accountLabel;
        private Label authenticationLabel;
        private Label openAiEndpointLabel;
        private Label ollamaEndpointLabel;
        private Label keyLabel;
        private Label keyNote;
        private Label approvalLabel;
        private ComboBox approvalPicker;
        private Label manualModelsLabel;
        private TextBox manualModels;
        private Label customNameLabel;
        private TextBox customName;
        private CheckBox azureEntra;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LlmSettingsWindow));
            this.grid = new System.Windows.Forms.TableLayoutPanel();
            this.providerLabel = new System.Windows.Forms.Label();
            this.provider = new System.Windows.Forms.ComboBox();
            this.accountLabel = new System.Windows.Forms.Label();
            this.codexStatus = new System.Windows.Forms.Label();
            this.authenticationLabel = new System.Windows.Forms.Label();
            this.codexActions = new System.Windows.Forms.FlowLayoutPanel();
            this.codexLogin = new System.Windows.Forms.Button();
            this.codexRefresh = new System.Windows.Forms.Button();
            this.openAiEndpointLabel = new System.Windows.Forms.Label();
            this.openAiEndpoint = new System.Windows.Forms.TextBox();
            this.ollamaEndpointLabel = new System.Windows.Forms.Label();
            this.ollamaEndpoint = new System.Windows.Forms.TextBox();
            this.keyLabel = new System.Windows.Forms.Label();
            this.openAiKey = new System.Windows.Forms.TextBox();
            this.clearKey = new System.Windows.Forms.CheckBox();
            this.keyNote = new System.Windows.Forms.Label();
            this.approvalLabel = new System.Windows.Forms.Label();
            this.approvalPicker = new System.Windows.Forms.ComboBox();
            this.contentLayout = new System.Windows.Forms.TableLayoutPanel();
            this.buttons = new System.Windows.Forms.FlowLayoutPanel();
            this.saveButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.manualModelsLabel = new System.Windows.Forms.Label();
            this.manualModels = new System.Windows.Forms.TextBox();
            this.customNameLabel = new System.Windows.Forms.Label();
            this.customName = new System.Windows.Forms.TextBox();
            this.azureEntra = new System.Windows.Forms.CheckBox();
            this.grid.SuspendLayout();
            this.githubLabel = new System.Windows.Forms.Label();
            this.githubStatus = new System.Windows.Forms.Label();
            this.githubAccountLabel = new System.Windows.Forms.Label();
            this.githubNote = new System.Windows.Forms.Label();
            this.githubAccount = new System.Windows.Forms.ComboBox();
            this.githubActions = new System.Windows.Forms.FlowLayoutPanel();
            this.githubLogin = new System.Windows.Forms.Button();
            this.githubRefresh = new System.Windows.Forms.Button();
            this.githubToolTips = new System.Windows.Forms.ToolTip();
            this.codexActions.SuspendLayout();
            this.contentLayout.SuspendLayout();
            this.buttons.SuspendLayout();
            this.SuspendLayout();
            //
            // grid
            //
            this.grid.AutoSize = true;
            this.grid.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.grid.ColumnCount = 2;
            this.grid.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.grid.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.grid.Controls.Add(this.providerLabel, 0, 0);
            this.grid.Controls.Add(this.provider, 1, 0);
            this.grid.Controls.Add(this.accountLabel, 0, 1);
            this.grid.Controls.Add(this.codexStatus, 1, 1);
            this.grid.Controls.Add(this.authenticationLabel, 0, 2);
            this.grid.Controls.Add(this.codexActions, 1, 2);
            this.grid.Controls.Add(this.openAiEndpointLabel, 0, 3);
            this.grid.Controls.Add(this.openAiEndpoint, 1, 3);
            this.grid.Controls.Add(this.ollamaEndpointLabel, 0, 4);
            this.grid.Controls.Add(this.ollamaEndpoint, 1, 4);
            this.grid.Controls.Add(this.keyLabel, 0, 5);
            this.grid.Controls.Add(this.openAiKey, 1, 5);
            this.grid.Controls.Add(this.clearKey, 1, 6);
            this.grid.Controls.Add(this.keyNote, 1, 7);
            this.grid.Controls.Add(this.approvalLabel, 0, 8);
            this.grid.Controls.Add(this.approvalPicker, 1, 8);
            this.grid.Controls.Add(this.manualModelsLabel, 0, 9);
            this.grid.Controls.Add(this.manualModels, 1, 9);
            this.grid.Controls.Add(this.customNameLabel, 0, 10);
            this.grid.Controls.Add(this.customName, 1, 10);
            this.grid.Controls.Add(this.azureEntra, 1, 11);
            this.grid.Controls.Add(this.githubLabel, 0, 12);
            this.grid.Controls.Add(this.githubStatus, 1, 12);
            this.grid.Controls.Add(this.githubAccountLabel, 0, 13);
            this.grid.Controls.Add(this.githubAccount, 1, 13);
            this.grid.Controls.Add(this.githubActions, 1, 14);
            this.grid.Controls.Add(this.githubNote, 1, 15);
            this.grid.Dock = System.Windows.Forms.DockStyle.Top;
            this.grid.Location = new System.Drawing.Point(0, 0);
            this.grid.Margin = new System.Windows.Forms.Padding(0);
            this.grid.Name = "grid";
            this.grid.Padding = new System.Windows.Forms.Padding(12, 12, 12, 0);
            this.grid.RowCount = 16;
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.grid.Size = new System.Drawing.Size(624, 256);
            this.grid.TabIndex = 0;
            // GitHub account (independent from the AI provider)
            this.githubLabel.Name = "githubLabel";
            this.githubLabel.Text = "GitHub · repositories";
            this.githubLabel.AutoSize = true;
            this.githubLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.githubStatus.Name = "githubStatus";
            this.githubStatus.Text = "Connect an account to synchronize your VBA sources.";
            this.githubStatus.AutoSize = true;
            this.githubStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubAccountLabel.Name = "githubAccountLabel";
            this.githubAccountLabel.Text = "GitHub account";
            this.githubAccountLabel.AutoSize = true;
            this.githubAccount.Name = "githubAccount";
            this.githubAccount.AccessibleName = "GitHub account for VBA repositories";
            this.githubAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.githubAccount.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubAccount.TabIndex = 10;
            this.githubActions.Name = "githubActions";
            this.githubActions.AutoSize = true;
            this.githubActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubActions.Controls.Add(this.githubLogin);
            this.githubActions.Controls.Add(this.githubRefresh);
            this.githubLogin.Name = "githubLogin";
            this.githubLogin.Text = "Sign in to GitHub";
            this.githubLogin.AutoSize = true;
            this.githubLogin.MinimumSize = new System.Drawing.Size(150, 30);
            this.githubLogin.TabIndex = 11;
            this.githubLogin.Click += new System.EventHandler(this.GitHubLogin_Click);
            this.githubRefresh.Name = "githubRefresh";
            this.githubRefresh.Text = "Refresh accounts";
            this.githubRefresh.AutoSize = true;
            this.githubRefresh.MinimumSize = new System.Drawing.Size(140, 30);
            this.githubRefresh.TabIndex = 12;
            this.githubRefresh.Click += new System.EventHandler(this.GitHubRefresh_Click);
            this.githubNote.Name = "githubNote";
            this.githubNote.Text = "Credentials stay in Git Credential Manager. Sign-in takes effect immediately; Save keeps the chosen account for VBAi. Copilot authentication is configured separately.";
            this.githubNote.AutoSize = true;
            this.githubNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubToolTips.SetToolTip(this.githubLogin, "Opens GitHub authentication in your browser. No token to copy into VBAi.");
            this.githubToolTips.SetToolTip(this.githubRefresh, "Reads locally saved accounts; does not check repository permissions yet.");
            this.githubToolTips.SetToolTip(this.githubAccount, "Account used for subsequent macro fetch, pull and push operations. Does not change global Git configuration.");
            //
            // providerLabel
            //
            this.providerLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.manualModelsLabel.Name = "manualModelsLabel";
            this.manualModelsLabel.Text = "Models (one per line)";
            this.manualModelsLabel.AutoSize = true;
            this.manualModelsLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.manualModels.Name = "manualModels";
            this.manualModels.AccessibleName = "Configured models or deployments, one identifier per line";
            this.manualModels.Multiline = true;
            this.manualModels.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.manualModels.Dock = System.Windows.Forms.DockStyle.Fill;
            this.manualModels.Size = new System.Drawing.Size(400, 70);
            this.manualModels.TabIndex = 7;
            this.customNameLabel.Name = "customNameLabel";
            this.customNameLabel.Text = "Provider name";
            this.customNameLabel.AutoSize = true;
            this.customNameLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.customName.Name = "customName";
            this.customName.AccessibleName = "Custom provider name";
            this.customName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.customName.TabIndex = 8;
            this.azureEntra.Name = "azureEntra";
            this.azureEntra.Text = "Use a Microsoft Entra token (renew after expiration)";
            this.azureEntra.AutoSize = true;
            this.azureEntra.Dock = System.Windows.Forms.DockStyle.Fill;
            this.azureEntra.TabIndex = 9;
            this.providerLabel.Location = new System.Drawing.Point(15, 12);
            this.providerLabel.Name = "providerLabel";
            this.providerLabel.Size = new System.Drawing.Size(100, 29);
            this.providerLabel.TabIndex = 0;
            this.providerLabel.Text = "Provider";
            this.providerLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // provider
            //
            this.provider.AccessibleName = "Provider";
            this.provider.Dock = System.Windows.Forms.DockStyle.Fill;
            this.provider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.provider.Location = new System.Drawing.Point(121, 15);
            this.provider.Name = "provider";
            this.provider.Size = new System.Drawing.Size(488, 23);
            this.provider.TabIndex = 0;
            //
            // accountLabel
            //
            this.accountLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.accountLabel.Location = new System.Drawing.Point(15, 41);
            this.accountLabel.Name = "accountLabel";
            this.accountLabel.Size = new System.Drawing.Size(100, 23);
            this.accountLabel.TabIndex = 1;
            this.accountLabel.Text = "ChatGPT account";
            this.accountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // codexStatus
            //
            this.codexStatus.AutoSize = true;
            this.codexStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.codexStatus.Location = new System.Drawing.Point(121, 41);
            this.codexStatus.Name = "codexStatus";
            this.codexStatus.Size = new System.Drawing.Size(488, 23);
            this.codexStatus.TabIndex = 2;
            this.codexStatus.Text = "Checking ChatGPT…";
            this.codexStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // authenticationLabel
            //
            this.authenticationLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.authenticationLabel.Location = new System.Drawing.Point(15, 64);
            this.authenticationLabel.Name = "authenticationLabel";
            this.authenticationLabel.Size = new System.Drawing.Size(100, 33);
            this.authenticationLabel.TabIndex = 3;
            this.authenticationLabel.Text = "Authentication";
            this.authenticationLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // codexActions
            //
            this.codexActions.AutoSize = true;
            this.codexActions.Controls.Add(this.codexLogin);
            this.codexActions.Controls.Add(this.codexRefresh);
            this.codexActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.codexActions.Location = new System.Drawing.Point(118, 64);
            this.codexActions.Margin = new System.Windows.Forms.Padding(0);
            this.codexActions.Name = "codexActions";
            this.codexActions.Padding = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.codexActions.Size = new System.Drawing.Size(494, 33);
            this.codexActions.TabIndex = 1;
            this.codexActions.WrapContents = false;
            //
            // codexLogin
            //
            this.codexLogin.Location = new System.Drawing.Point(3, 3);
            this.codexLogin.Name = "codexLogin";
            this.codexLogin.Size = new System.Drawing.Size(185, 23);
            this.codexLogin.TabIndex = 0;
            this.codexLogin.Text = "Sign in to ChatGPT";
            //
            // codexRefresh
            //
            this.codexRefresh.Location = new System.Drawing.Point(194, 3);
            this.codexRefresh.Name = "codexRefresh";
            this.codexRefresh.Size = new System.Drawing.Size(130, 23);
            this.codexRefresh.TabIndex = 1;
            this.codexRefresh.Text = "Refresh status";
            //
            // openAiEndpointLabel
            //
            this.openAiEndpointLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiEndpointLabel.Location = new System.Drawing.Point(15, 97);
            this.openAiEndpointLabel.Name = "openAiEndpointLabel";
            this.openAiEndpointLabel.Size = new System.Drawing.Size(100, 29);
            this.openAiEndpointLabel.TabIndex = 4;
            this.openAiEndpointLabel.Text = "OpenAI URL (optional)";
            this.openAiEndpointLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // openAiEndpoint
            //
            this.openAiEndpoint.AccessibleName = "Optional OpenAI URL";
            this.openAiEndpoint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiEndpoint.Location = new System.Drawing.Point(121, 100);
            this.openAiEndpoint.Name = "openAiEndpoint";
            this.openAiEndpoint.Size = new System.Drawing.Size(488, 23);
            this.openAiEndpoint.TabIndex = 2;
            //
            // ollamaEndpointLabel
            //
            this.ollamaEndpointLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaEndpointLabel.Location = new System.Drawing.Point(15, 126);
            this.ollamaEndpointLabel.Name = "ollamaEndpointLabel";
            this.ollamaEndpointLabel.Size = new System.Drawing.Size(100, 29);
            this.ollamaEndpointLabel.TabIndex = 5;
            this.ollamaEndpointLabel.Text = "Ollama URL (optional)";
            this.ollamaEndpointLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // ollamaEndpoint
            //
            this.ollamaEndpoint.AccessibleName = "Optional Ollama URL";
            this.ollamaEndpoint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaEndpoint.Location = new System.Drawing.Point(121, 129);
            this.ollamaEndpoint.Name = "ollamaEndpoint";
            this.ollamaEndpoint.Size = new System.Drawing.Size(488, 23);
            this.ollamaEndpoint.TabIndex = 3;
            //
            // keyLabel
            //
            this.keyLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.keyLabel.Location = new System.Drawing.Point(15, 155);
            this.keyLabel.Name = "keyLabel";
            this.keyLabel.Size = new System.Drawing.Size(100, 29);
            this.keyLabel.TabIndex = 6;
            this.keyLabel.Text = "New OpenAI API key";
            this.keyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // openAiKey
            //
            this.openAiKey.AccessibleName = "New OpenAI API key";
            this.openAiKey.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiKey.Location = new System.Drawing.Point(121, 158);
            this.openAiKey.Name = "openAiKey";
            this.openAiKey.Size = new System.Drawing.Size(488, 23);
            this.openAiKey.TabIndex = 4;
            this.openAiKey.UseSystemPasswordChar = true;
            //
            // clearKey
            //
            this.clearKey.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clearKey.Location = new System.Drawing.Point(121, 187);
            this.clearKey.Name = "clearKey";
            this.clearKey.Size = new System.Drawing.Size(488, 24);
            this.clearKey.TabIndex = 5;
            this.clearKey.Text = "Remove saved API key";
            //
            // keyNote
            //
            this.keyNote.AutoSize = true;
            this.keyNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.keyNote.Location = new System.Drawing.Point(121, 214);
            this.keyNote.Name = "keyNote";
            this.keyNote.Size = new System.Drawing.Size(488, 15);
            this.keyNote.TabIndex = 7;
            this.keyNote.Text = "Leave the key blank to keep the current one. Secrets are encrypted for this Windows account.";
            //
            // approvalLabel
            //
            this.approvalLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.approvalLabel.Location = new System.Drawing.Point(15, 229);
            this.approvalLabel.Name = "approvalLabel";
            this.approvalLabel.Size = new System.Drawing.Size(100, 27);
            this.approvalLabel.TabIndex = 8;
            this.approvalLabel.Text = "VBE edits";
            this.approvalLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // approvalPicker
            //
            this.approvalPicker.AccessibleName = "VBE edit permissions";
            this.approvalPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.approvalPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.approvalPicker.Location = new System.Drawing.Point(121, 232);
            this.approvalPicker.Name = "approvalPicker";
            this.approvalPicker.Size = new System.Drawing.Size(488, 23);
            this.approvalPicker.TabIndex = 6;
            //
            // contentLayout
            //
            this.contentLayout.AutoSize = true;
            this.contentLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.contentLayout.ColumnCount = 1;
            this.contentLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.contentLayout.Controls.Add(this.grid, 0, 0);
            this.contentLayout.Controls.Add(this.buttons, 0, 1);
            this.contentLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.contentLayout.Location = new System.Drawing.Point(0, 0);
            this.contentLayout.Name = "contentLayout";
            this.contentLayout.RowCount = 2;
            this.contentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.contentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.contentLayout.Size = new System.Drawing.Size(624, 301);
            this.contentLayout.TabIndex = 0;
            //
            // buttons
            //
            this.buttons.AutoSize = true;
            this.buttons.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttons.Controls.Add(this.saveButton);
            this.buttons.Controls.Add(this.cancelButton);
            this.buttons.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.buttons.Location = new System.Drawing.Point(0, 256);
            this.buttons.Margin = new System.Windows.Forms.Padding(0);
            this.buttons.Name = "buttons";
            this.buttons.Padding = new System.Windows.Forms.Padding(12, 4, 12, 12);
            this.buttons.Size = new System.Drawing.Size(624, 45);
            this.buttons.TabIndex = 1;
            //
            // saveButton
            //
            this.saveButton.Location = new System.Drawing.Point(492, 7);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(105, 23);
            this.saveButton.TabIndex = 0;
            this.saveButton.Text = "Save";
            //
            // cancelButton
            //
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(381, 7);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(105, 23);
            this.cancelButton.TabIndex = 1;
            this.cancelButton.Text = "Cancel";
            //
            // LlmSettingsWindow
            //
            this.AcceptButton = this.saveButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(624, 368);
            this.Controls.Add(this.contentLayout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(640, 39);
            this.Name = "LlmSettingsWindow";
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "VBAi — Settings";
            this.grid.ResumeLayout(false);
            this.grid.PerformLayout();
            this.codexActions.ResumeLayout(false);
            this.contentLayout.ResumeLayout(false);
            this.contentLayout.PerformLayout();
            this.buttons.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
