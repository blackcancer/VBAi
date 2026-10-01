namespace VBAi
{
    /// <summary>Contrôles générés pour les fournisseurs, modèles et stratégies d’approbation.</summary>
    public sealed partial class ProviderSettingsView
    {
        /// <summary>Grille des paramètres du fournisseur.</summary>
        internal System.Windows.Forms.TableLayoutPanel grid;
        /// <summary>Disposition des actions CLI Codex.</summary>
        internal System.Windows.Forms.FlowLayoutPanel codexActions;
        /// <summary>Sélecteur du fournisseur LLM.</summary>
        internal VBAi.ThemedComboBox provider;
        /// <summary>État de la connexion Codex ou Copilot.</summary>
        internal System.Windows.Forms.Label codexStatus;
        /// <summary>Action de connexion du CLI sélectionné.</summary>
        internal VBAi.ThemedButton codexLogin;
        /// <summary>Action de relecture du statut CLI.</summary>
        internal VBAi.ThemedButton codexRefresh;
        /// <summary>Champ de point de terminaison OpenAI.</summary>
        internal VBAi.UiTextBox openAiEndpoint;
        /// <summary>Champ de point de terminaison Ollama.</summary>
        internal VBAi.UiTextBox ollamaEndpoint;
        /// <summary>Champ de clé du fournisseur.</summary>
        internal VBAi.UiTextBox openAiKey;
        /// <summary>Option de suppression de la clé enregistrée.</summary>
        internal System.Windows.Forms.CheckBox clearKey;
        /// <summary>Libellé du sélecteur de fournisseur.</summary>
        internal System.Windows.Forms.Label providerLabel;
        /// <summary>Libellé du compte fournisseur.</summary>
        internal System.Windows.Forms.Label accountLabel;
        /// <summary>Libellé de la section d’authentification.</summary>
        internal System.Windows.Forms.Label authenticationLabel;
        /// <summary>Libellé de l’adresse OpenAI.</summary>
        internal System.Windows.Forms.Label openAiEndpointLabel;
        /// <summary>Libellé de l’adresse Ollama.</summary>
        internal System.Windows.Forms.Label ollamaEndpointLabel;
        /// <summary>Libellé du champ de clé.</summary>
        internal System.Windows.Forms.Label keyLabel;
        /// <summary>Note de configuration de la clé fournisseur.</summary>
        internal System.Windows.Forms.Label keyNote;
        /// <summary>Libellé de la stratégie d’approbation VBE.</summary>
        internal System.Windows.Forms.Label approvalLabel;
        /// <summary>Sélecteur de la stratégie d’approbation.</summary>
        internal VBAi.ThemedComboBox approvalPicker;
        /// <summary>Libellé de la liste de modèles manuelle.</summary>
        internal System.Windows.Forms.Label manualModelsLabel;
        /// <summary>Champ des modèles saisis manuellement.</summary>
        internal VBAi.UiTextBox manualModels;
        /// <summary>Libellé du fournisseur personnalisé.</summary>
        internal System.Windows.Forms.Label customNameLabel;
        /// <summary>Champ du nom personnalisé.</summary>
        internal VBAi.UiTextBox customName;
        /// <summary>Option d’authentification Azure par Entra.</summary>
        internal System.Windows.Forms.CheckBox azureEntra;
        /// <summary>Label for the optional Ollama sampling temperature.</summary>
        internal System.Windows.Forms.Label ollamaTemperatureLabel;
        /// <summary>Optional temperature override; blank retains server behavior.</summary>
        internal VBAi.UiTextBox ollamaTemperature;
        /// <summary>Label for the optional Ollama nucleus sampling limit.</summary>
        internal System.Windows.Forms.Label ollamaTopPLabel;
        /// <summary>Optional top-p override; blank retains server behavior.</summary>
        internal VBAi.UiTextBox ollamaTopP;
        /// <summary>Conteneur des composants WinForms non visuels.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Info-bulles appartenant à la vue.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Crée les contrôles des paramètres fournisseur et applique leurs propriétés Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.grid = new System.Windows.Forms.TableLayoutPanel();
            this.providerLabel = new System.Windows.Forms.Label();
            this.provider = new VBAi.ThemedComboBox();
            this.accountLabel = new System.Windows.Forms.Label();
            this.codexStatus = new System.Windows.Forms.Label();
            this.authenticationLabel = new System.Windows.Forms.Label();
            this.codexActions = new System.Windows.Forms.FlowLayoutPanel();
            this.codexLogin = new VBAi.ThemedButton();
            this.codexRefresh = new VBAi.ThemedButton();
            this.openAiEndpointLabel = new System.Windows.Forms.Label();
            this.openAiEndpoint = new VBAi.UiTextBox();
            this.ollamaEndpointLabel = new System.Windows.Forms.Label();
            this.ollamaEndpoint = new VBAi.UiTextBox();
            this.keyLabel = new System.Windows.Forms.Label();
            this.openAiKey = new VBAi.UiTextBox();
            this.clearKey = new System.Windows.Forms.CheckBox();
            this.keyNote = new System.Windows.Forms.Label();
            this.approvalLabel = new System.Windows.Forms.Label();
            this.approvalPicker = new VBAi.ThemedComboBox();
            this.manualModelsLabel = new System.Windows.Forms.Label();
            this.manualModels = new VBAi.UiTextBox();
            this.customNameLabel = new System.Windows.Forms.Label();
            this.customName = new VBAi.UiTextBox();
            this.azureEntra = new System.Windows.Forms.CheckBox();
            this.ollamaTemperatureLabel = new System.Windows.Forms.Label();
            this.ollamaTemperature = new VBAi.UiTextBox();
            this.ollamaTopPLabel = new System.Windows.Forms.Label();
            this.ollamaTopP = new VBAi.UiTextBox();
            this.grid.SuspendLayout();
            this.codexActions.SuspendLayout();
            this.SuspendLayout();
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
            this.grid.Controls.Add(this.ollamaTemperatureLabel, 0, 12);
            this.grid.Controls.Add(this.ollamaTemperature, 1, 12);
            this.grid.Controls.Add(this.ollamaTopPLabel, 0, 13);
            this.grid.Controls.Add(this.ollamaTopP, 1, 13);
            this.grid.Dock = System.Windows.Forms.DockStyle.Top;
            this.grid.Location = new System.Drawing.Point(0, 0);
            this.grid.Margin = new System.Windows.Forms.Padding(0);
            this.grid.Name = "grid";
            this.grid.Padding = new System.Windows.Forms.Padding(12, 12, 12, 0);
            this.grid.RowCount = 14;
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
            this.ollamaTemperatureLabel.Name = "ollamaTemperatureLabel";
            this.ollamaTemperatureLabel.Text = "Temperature (0 to 2)";
            this.ollamaTemperatureLabel.AutoSize = true;
            this.ollamaTemperatureLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ollamaTemperatureLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaTemperatureLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.ollamaTemperatureLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaTemperature.Name = "ollamaTemperature";
            this.ollamaTemperature.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ollamaTemperature.AccessibleName = "Ollama temperature; blank uses server default";
            this.ollamaTemperature.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaTemperature.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaTemperature.TabIndex = 10;
            this.toolTips.SetToolTip(this.ollamaTemperature, "Optional. Blank uses the server default; 0 reduces sampling randomness. Use a decimal point or your local decimal separator.");
            this.ollamaTopPLabel.Name = "ollamaTopPLabel";
            this.ollamaTopPLabel.Text = "Top-p (above 0 to 1)";
            this.ollamaTopPLabel.AutoSize = true;
            this.ollamaTopPLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ollamaTopPLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaTopPLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.ollamaTopPLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaTopP.Name = "ollamaTopP";
            this.ollamaTopP.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ollamaTopP.AccessibleName = "Ollama top-p; blank uses server default";
            this.ollamaTopP.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaTopP.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaTopP.TabIndex = 11;
            this.toolTips.SetToolTip(this.ollamaTopP, "Optional nucleus sampling limit. Blank uses the server default. Use a decimal point or your local decimal separator.");
            this.providerLabel.Location = new System.Drawing.Point(15, 12);
            this.providerLabel.Name = "providerLabel";
            this.providerLabel.Size = new System.Drawing.Size(100, 29);
            this.providerLabel.TabIndex = 0;
            this.providerLabel.Text = "Provider";
            this.providerLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.provider.AccessibleName = "Provider";
            this.provider.Dock = System.Windows.Forms.DockStyle.Fill;
            this.provider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.provider.Location = new System.Drawing.Point(121, 15);
            this.provider.Name = "provider";
            this.provider.Size = new System.Drawing.Size(488, 23);
            this.provider.TabIndex = 0;
            this.accountLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.accountLabel.Location = new System.Drawing.Point(15, 41);
            this.accountLabel.Name = "accountLabel";
            this.accountLabel.Size = new System.Drawing.Size(100, 23);
            this.accountLabel.TabIndex = 1;
            this.accountLabel.Text = "ChatGPT account";
            this.accountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.codexStatus.AutoSize = true;
            this.codexStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.codexStatus.Location = new System.Drawing.Point(121, 41);
            this.codexStatus.Name = "codexStatus";
            this.codexStatus.Size = new System.Drawing.Size(488, 23);
            this.codexStatus.TabIndex = 2;
            this.codexStatus.Text = "Checking ChatGPT…";
            this.codexStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.authenticationLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.authenticationLabel.Location = new System.Drawing.Point(15, 64);
            this.authenticationLabel.Name = "authenticationLabel";
            this.authenticationLabel.Size = new System.Drawing.Size(100, 33);
            this.authenticationLabel.TabIndex = 3;
            this.authenticationLabel.Text = "Authentication";
            this.authenticationLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            this.codexLogin.Location = new System.Drawing.Point(3, 3);
            this.codexLogin.Name = "codexLogin";
            this.codexLogin.Size = new System.Drawing.Size(185, 23);
            this.codexLogin.TabIndex = 0;
            this.codexLogin.Text = "Sign in to ChatGPT";
            this.codexRefresh.Location = new System.Drawing.Point(194, 3);
            this.codexRefresh.Name = "codexRefresh";
            this.codexRefresh.Size = new System.Drawing.Size(130, 23);
            this.codexRefresh.TabIndex = 1;
            this.codexRefresh.Text = "Refresh status";
            this.openAiEndpointLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiEndpointLabel.Location = new System.Drawing.Point(15, 97);
            this.openAiEndpointLabel.Name = "openAiEndpointLabel";
            this.openAiEndpointLabel.Size = new System.Drawing.Size(100, 29);
            this.openAiEndpointLabel.TabIndex = 4;
            this.openAiEndpointLabel.Text = "OpenAI URL (optional)";
            this.openAiEndpointLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.openAiEndpoint.AccessibleName = "Optional OpenAI URL";
            this.openAiEndpoint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiEndpoint.Location = new System.Drawing.Point(121, 100);
            this.openAiEndpoint.Name = "openAiEndpoint";
            this.openAiEndpoint.Size = new System.Drawing.Size(488, 23);
            this.openAiEndpoint.TabIndex = 2;
            this.ollamaEndpointLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaEndpointLabel.Location = new System.Drawing.Point(15, 126);
            this.ollamaEndpointLabel.Name = "ollamaEndpointLabel";
            this.ollamaEndpointLabel.Size = new System.Drawing.Size(100, 29);
            this.ollamaEndpointLabel.TabIndex = 5;
            this.ollamaEndpointLabel.Text = "Ollama URL (optional)";
            this.ollamaEndpointLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ollamaEndpoint.AccessibleName = "Optional Ollama URL";
            this.ollamaEndpoint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ollamaEndpoint.Location = new System.Drawing.Point(121, 129);
            this.ollamaEndpoint.Name = "ollamaEndpoint";
            this.ollamaEndpoint.Size = new System.Drawing.Size(488, 23);
            this.ollamaEndpoint.TabIndex = 3;
            this.keyLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.keyLabel.Location = new System.Drawing.Point(15, 155);
            this.keyLabel.Name = "keyLabel";
            this.keyLabel.Size = new System.Drawing.Size(100, 29);
            this.keyLabel.TabIndex = 6;
            this.keyLabel.Text = "New OpenAI API key";
            this.keyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.openAiKey.AccessibleName = "New OpenAI API key";
            this.openAiKey.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openAiKey.Location = new System.Drawing.Point(121, 158);
            this.openAiKey.Name = "openAiKey";
            this.openAiKey.Size = new System.Drawing.Size(488, 23);
            this.openAiKey.TabIndex = 4;
            this.openAiKey.UseSystemPasswordChar = true;
            this.clearKey.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clearKey.Location = new System.Drawing.Point(121, 187);
            this.clearKey.Name = "clearKey";
            this.clearKey.Size = new System.Drawing.Size(488, 24);
            this.clearKey.TabIndex = 5;
            this.clearKey.Text = "Remove saved API key";
            this.keyNote.AutoSize = true;
            this.keyNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.keyNote.Location = new System.Drawing.Point(121, 214);
            this.keyNote.Name = "keyNote";
            this.keyNote.Size = new System.Drawing.Size(488, 15);
            this.keyNote.TabIndex = 7;
            this.keyNote.Text = "Leave the key blank to keep the current one. Secrets are encrypted for this Windows account.";
            this.approvalLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.approvalLabel.Location = new System.Drawing.Point(15, 229);
            this.approvalLabel.Name = "approvalLabel";
            this.approvalLabel.Size = new System.Drawing.Size(100, 27);
            this.approvalLabel.TabIndex = 8;
            this.approvalLabel.Text = "VBE edits";
            this.approvalLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.approvalPicker.AccessibleName = "VBE edit permissions";
            this.approvalPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.approvalPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.approvalPicker.Location = new System.Drawing.Point(121, 232);
            this.approvalPicker.Name = "approvalPicker";
            this.approvalPicker.Size = new System.Drawing.Size(488, 23);
            this.approvalPicker.TabIndex = 6;
            this.providerLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.provider.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.accountLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.codexStatus.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.authenticationLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.codexActions.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.openAiEndpointLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.openAiEndpoint.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaEndpointLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.ollamaEndpoint.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.keyLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.openAiKey.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.clearKey.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.keyNote.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.approvalLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.approvalPicker.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.manualModelsLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.manualModels.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.customNameLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.customName.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.azureEntra.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.providerLabel.AutoSize = true;
            this.accountLabel.AutoSize = true;
            this.authenticationLabel.AutoSize = true;
            this.openAiEndpointLabel.AutoSize = true;
            this.ollamaEndpointLabel.AutoSize = true;
            this.keyLabel.AutoSize = true;
            this.approvalLabel.AutoSize = true;
            this.clearKey.AutoSize = true;
            this.providerLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.accountLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.authenticationLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.openAiEndpointLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.ollamaEndpointLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.keyLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.approvalLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.manualModelsLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.customNameLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.codexLogin.AutoSize = true;
            this.codexRefresh.AutoSize = true;
            this.codexLogin.MinimumSize = new System.Drawing.Size(185, 30);
            this.codexRefresh.MinimumSize = new System.Drawing.Size(130, 30);
            this.codexLogin.Padding = new System.Windows.Forms.Padding(6, 2, 6, 2);
            this.codexRefresh.Padding = new System.Windows.Forms.Padding(6, 2, 6, 2);
            this.AutoScroll = true;
            this.Controls.Add(this.grid);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ProviderSettingsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.manualModelsLabel.Location = new System.Drawing.Point(12, 345);
            this.manualModelsLabel.Size = new System.Drawing.Size(156, 70);
            this.manualModelsLabel.TabIndex = 16;
            this.manualModels.Location = new System.Drawing.Point(180, 345);
            this.customNameLabel.Location = new System.Drawing.Point(12, 427);
            this.customNameLabel.Size = new System.Drawing.Size(156, 23);
            this.customNameLabel.TabIndex = 18;
            this.customName.Location = new System.Drawing.Point(180, 427);
            this.customName.Size = new System.Drawing.Size(656, 23);
            this.azureEntra.Location = new System.Drawing.Point(180, 462);
            this.azureEntra.Size = new System.Drawing.Size(656, 22);
            this.grid.ResumeLayout(false);
            this.grid.PerformLayout();
            this.codexActions.ResumeLayout(false);
            this.codexActions.PerformLayout();
            this.codexLogin.Symbol = VBAi.UiSymbol.Next;
            this.codexRefresh.Symbol = VBAi.UiSymbol.Refresh;
            this.codexRefresh.IconOnly = true;
            this.codexRefresh.AutoSize = false;
            this.codexRefresh.MinimumSize = System.Drawing.Size.Empty;
            this.codexRefresh.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
