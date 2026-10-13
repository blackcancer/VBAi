namespace VBAi
{

    /// <summary>Fenêtre de conversation et commandes de son concepteur WinForms.</summary>
    internal sealed partial class ChatWindow
    {

        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components = null;
        private Blackcancer.ChatUI.Views.ChatWorkspaceView workspace;

        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Menu des commandes de configuration et d’intégration.</summary>
        private System.Windows.Forms.ContextMenuStrip optionsMenu;

        /// <summary>Commande qui ouvre les paramètres de l’application.</summary>
        private System.Windows.Forms.ToolStripMenuItem configure;

        /// <summary>Panel that displays and applies project read/write consent for this conversation.</summary>
        private System.Windows.Forms.ToolStripMenuItem projectAccess;

        /// <summary>Button that resumes a paused response when its session can continue.</summary>
        private System.Windows.Forms.ToolStripMenuItem resumeTurn;

        /// <summary>Commande de menu qui ouvre les informations de l’application.</summary>
        private System.Windows.Forms.ToolStripMenuItem about;

        /// <summary>Séparateur placé avant la commande d’informations.</summary>
        private System.Windows.Forms.ToolStripSeparator aboutSeparator;

        /// <summary>Commande qui actualise les modèles du fournisseur sélectionné.</summary>
        private System.Windows.Forms.ToolStripMenuItem refreshModels;

        /// <summary>Commande de gestion de l’ancrage de la fenêtre.</summary>
        private System.Windows.Forms.ToolStripMenuItem docking;

        /// <summary>Commande qui ouvre les fonctions GitHub.</summary>
        private System.Windows.Forms.ToolStripMenuItem github;

        /// <summary>Disposition racine du contenu de la fenêtre.</summary>
        private System.Windows.Forms.TableLayoutPanel rootLayout;

        /// <summary>Disposition du titre de la conversation et des actions de fenêtre.</summary>
        private System.Windows.Forms.TableLayoutPanel headingLayout;

        /// <summary>Libellé du nom de l’application.</summary>
        private System.Windows.Forms.Label appTitle;

        /// <summary>Libellé du titre de la session active.</summary>
        private System.Windows.Forms.Label sessionTitle;

        /// <summary>Commande de création d’une conversation.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton newChat;

        /// <summary>Commande d’accès aux options.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton options;

        /// <summary>Disposition du sélecteur de portée de projet et de l’historique.</summary>
        private System.Windows.Forms.TableLayoutPanel scopeLayout;

        /// <summary>Commande d’affichage de l’historique.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton history;

        /// <summary>Sélecteur du projet auquel la conversation est liée.</summary>
        private System.Windows.Forms.ComboBox scopePicker;

        /// <summary>Disposition des commandes de mode et de workflow.</summary>
        private System.Windows.Forms.TableLayoutPanel workflowLayout;

        /// <summary>Approval policy for the next VBA actions.</summary>
        private System.Windows.Forms.ComboBox approvalPicker;

        /// <summary>Compact model and effort summary.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton modelSummary;

        /// <summary>Automatic verification command.</summary>
        private System.Windows.Forms.ToolStripMenuItem verifyChanges;

        /// <summary>Sélecteur du mode de conversation.</summary>
        private System.Windows.Forms.ComboBox modePicker;

        /// <summary>Commande de sélection de code ou de référence.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton selection;

        /// <summary>Commande de compilation du projet courant.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton compile;

        /// <summary>Option de vérification après une modification de code.</summary>
        private System.Windows.Forms.CheckBox verifyAfterEdit;

        /// <summary>Conteneur de la conversation et de ses vues.</summary>
        private System.Windows.Forms.Panel conversationPanel;

        /// <summary>Panneau de recherche et de gestion des sessions.</summary>
        private System.Windows.Forms.Panel historyPanel;

        /// <summary>Disposition des contrôles de l’historique.</summary>
        private System.Windows.Forms.TableLayoutPanel historyLayout;

        /// <summary>Libellé de la liste des conversations.</summary>
        private System.Windows.Forms.Label historyLabel;

        /// <summary>Champ de filtrage des sessions.</summary>
        private System.Windows.Forms.TextBox historySearch;

        /// <summary>Liste des sessions de la portée courante.</summary>
        private System.Windows.Forms.ListBox sessionList;

        /// <summary>Option d’inclusion des sessions archivées.</summary>
        private System.Windows.Forms.CheckBox showArchived;

        /// <summary>Champ de modification du titre de la conversation.</summary>
        private System.Windows.Forms.TextBox chatTitleEditor;

        /// <summary>Disposition des actions appliquées à la session.</summary>
        private System.Windows.Forms.FlowLayoutPanel historyActions;

        /// <summary>Commande de renommage de la session.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton rename;

        /// <summary>Commande d’archivage ou de restauration de la session.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton archive;

        /// <summary>Commande d’épinglage de la session.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton pin;

        /// <summary>Deletes the selected local conversation.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton deleteSession;

        /// <summary>Commande d’export de la conversation.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton export;

        /// <summary>Commande d’affichage ou de masquage de la mémoire de projet.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton memoryToggle;

        /// <summary>Groupe d’édition de la mémoire de projet.</summary>
        private System.Windows.Forms.GroupBox memoryPanel;

        /// <summary>Disposition du champ et des commandes de mémoire.</summary>
        private System.Windows.Forms.TableLayoutPanel memoryLayout;

        /// <summary>Champ d’édition de la mémoire de projet.</summary>
        private System.Windows.Forms.TextBox memoryEditor;

        /// <summary>Commande d’enregistrement de la mémoire de projet.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton saveMemory;

        /// <summary>Option d’inclusion de la mémoire dans le prochain message.</summary>
        private System.Windows.Forms.CheckBox attachMemory;

        /// <summary>Panneau contenant le transcript de conversation.</summary>
        private System.Windows.Forms.Panel transcriptPanel;

        /// <summary>Hôte WinForms du transcript WPF.</summary>
        private VBAi.ChatContentHost transcriptHost;

        /// <summary>Conteneur de composition du message et de ses commandes.</summary>
        private System.Windows.Forms.TableLayoutPanel composerLayout;

        /// <summary>Disposition des références et pièces jointes sélectionnées.</summary>
        private System.Windows.Forms.FlowLayoutPanel contextChips;

        /// <summary>Contains the transcript rows for messages waiting to be sent.</summary>
        private System.Windows.Forms.FlowLayoutPanel pendingMessagesPanel;

        /// <summary>Hôte WinForms du compositeur WPF.</summary>
        private Blackcancer.ChatUI.Views.ChatInputView promptHost;

        /// <summary>Disposition des commandes du compositeur.</summary>
        private System.Windows.Forms.TableLayoutPanel composerActions;

        /// <summary>Commande d’ajout ou de sélection d’un module.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton modules;

        /// <summary>Commande d’ajout ou de sélection d’une procédure.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton methods;

        /// <summary>Commande d’envoi du message.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton send;

        /// <summary>Commande d’affichage du contexte de conversation.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton contextToggle;

        /// <summary>Panneau contenant l’aperçu du contexte courant.</summary>
        private System.Windows.Forms.Panel contextPanel;

        /// <summary>Disposition des éléments d’aperçu du contexte.</summary>
        private System.Windows.Forms.FlowLayoutPanel contextPreview;

        /// <summary>Disposition des sélecteurs de fournisseur, modèle et effort.</summary>
        private System.Windows.Forms.TableLayoutPanel providerLayout;

        /// <summary>Sélecteur du fournisseur LLM.</summary>
        private System.Windows.Forms.ComboBox providerPicker;

        /// <summary>Sélecteur du modèle du fournisseur.</summary>
        private System.Windows.Forms.ComboBox modelPicker;

        /// <summary>Sélecteur du niveau d’effort du modèle.</summary>
        private System.Windows.Forms.ComboBox effortPicker;

        /// <summary>Disposition du pied de fenêtre.</summary>
        private System.Windows.Forms.TableLayoutPanel footerLayout;

        /// <summary>Commande de retour au dernier message du transcript.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton jumpToLatest;

        /// <summary>Commande d’affichage des modifications de code de la session.</summary>
        private Blackcancer.ChatUI.Views.ChatActionButton changes;

        /// <summary>Disposition du statut et de l’indicateur d’activité.</summary>
        private System.Windows.Forms.TableLayoutPanel statusLayout;

        /// <summary>Libellé des messages d’état de l’interface.</summary>
        private System.Windows.Forms.Label status;

        /// <summary>Indicateur de progression de l’activité en cours.</summary>
        private System.Windows.Forms.ProgressBar activityBar;

        /// <summary>Libère les ressources d’exécution et les composants WinForms du formulaire.</summary>
        /// <param name="disposing"><see langword="true"/> pour libérer les ressources managées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                DisposeRuntime();
                shutdownDiagnostic?.Enter("chat.Components.Dispose");
                if (components != null) components.Dispose();
                shutdownDiagnostic?.Returned("chat.Components.Dispose");
            }
            shutdownDiagnostic?.Enter("chat.Base.Dispose");
            base.Dispose(disposing);
            shutdownDiagnostic?.Returned("chat.Base.Dispose");
        }

        #region Windows Form Designer generated code
        // The fixed workspace is owned by the shared Designer. Domain controls below
        // are VBAi extensions; no provider or native service is started here.
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            workspace = new Blackcancer.ChatUI.Views.ChatWorkspaceView();
            workspace.Dock = System.Windows.Forms.DockStyle.Fill;
            methods = new Blackcancer.ChatUI.Views.ChatActionButton();
            compile = new Blackcancer.ChatUI.Views.ChatActionButton();
            changes = new Blackcancer.ChatUI.Views.ChatActionButton();
            transcriptHost = new VBAi.ChatContentHost();
            projectAccess = new System.Windows.Forms.ToolStripMenuItem();
            resumeTurn = new System.Windows.Forms.ToolStripMenuItem();
            refreshModels = new System.Windows.Forms.ToolStripMenuItem();
            verifyChanges = new System.Windows.Forms.ToolStripMenuItem();
            docking = new System.Windows.Forms.ToolStripMenuItem();
            github = new System.Windows.Forms.ToolStripMenuItem();
            aboutSeparator = new System.Windows.Forms.ToolStripSeparator();
            about = new System.Windows.Forms.ToolStripMenuItem();
            verifyAfterEdit = new System.Windows.Forms.CheckBox();
            SuspendLayout();
            methods.Name = "methods";
            methods.Text = "@ Function";
            methods.Symbol = "\uE943";
            methods.IconOnly = true;
            methods.AutoSize = false;
            methods.Size = new System.Drawing.Size(32, 30);
            compile.Name = "compile";
            compile.Text = "Compile";
            compile.Symbol = "\uE73E";
            compile.IconOnly = true;
            compile.AutoSize = false;
            compile.Size = new System.Drawing.Size(32, 30);
            changes.Name = "changes";
            changes.Text = "Changes · 0";
            changes.Symbol = "\uE943";
            changes.IconOnly = true;
            changes.AutoSize = false;
            changes.Size = new System.Drawing.Size(32, 30);
            changes.Enabled = false;
            workspace.HeaderActionsHost.Controls.Add(methods);
            workspace.HeaderActionsHost.Controls.Add(compile);
            workspace.HeaderActionsHost.Controls.Add(changes);
            transcriptHost.Name = "transcriptHost";
            transcriptHost.Dock = System.Windows.Forms.DockStyle.Fill;
            workspace.TranscriptHost.Controls.Add(transcriptHost);
            projectAccess.Name = "projectAccess";
            projectAccess.Text = "Project access…";
            projectAccess.ToolTipText = "Choose which other projects this conversation may read.";
            resumeTurn.Name = "resumeTurn";
            resumeTurn.Text = "Resume paused turn";
            resumeTurn.ToolTipText = "Continue from saved tool results without repeating completed actions.";
            refreshModels.Name = "refreshModels";
            refreshModels.Text = "Refresh models";
            verifyChanges.Name = "verifyChanges";
            verifyChanges.Text = "Verify after changes";
            verifyChanges.CheckOnClick = true;
            verifyChanges.Checked = true;
            docking.Name = "docking";
            docking.Text = "Docked / floating window";
            github.Name = "github";
            github.Text = "GitHub · synchronize VBA…";
            github.ToolTipText = "Export and synchronize the current document's sources with a GitHub repository.";
            aboutSeparator.Name = "aboutSeparator";
            about.Name = "about";
            about.Text = "About VBAi";
            about.ToolTipText = "Version, technical details and project resources";
            verifyAfterEdit.Name = "verifyAfterEdit";
            verifyAfterEdit.Checked = true;
            verifyAfterEdit.Visible = false;
            components.Add(verifyAfterEdit);
            workspace.optionsMenu.Items.Clear();
            workspace.optionsMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                workspace.configureMenuItem, projectAccess, resumeTurn, refreshModels, verifyChanges,
                docking, github, aboutSeparator, about });
            workspace.attachContextButton.Text = "# Context";
            workspace.attachContextButton.Symbol = "\uE8B7";
            workspace.attachSelectionButton.Text = "Attach code selection";
            workspace.attachSelectionButton.Symbol = "\uE723";
            workspace.AdditionalContextHost.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            workspace.AdditionalContextHost.WrapContents = false;
            workspace.AdditionalContextHost.AutoScroll = true;
            workspace.AdditionalContextHost.MinimumSize = new System.Drawing.Size(0, 128);
            Controls.Add(workspace);
            AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            ClientSize = new System.Drawing.Size(600, 820);
            Font = new System.Drawing.Font("Segoe UI", 9F);
            KeyPreview = true;
            MinimumSize = new System.Drawing.Size(440, 560);
            Name = "ChatWindow";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "VBAi — Your AI agent for VBA";
            ResumeLayout(false);
            PerformLayout();
        }

        private void BindSharedParts()
        {
            workspace.ApplyCulture(UiText.Culture);
            workspace.ResetFont();
            var parts = workspace.Parts;
            toolTips = parts.toolTips;
            optionsMenu = parts.optionsMenu;
            configure = parts.configureMenuItem;
            rootLayout = parts.rootLayout;
            headingLayout = parts.headingLayout;
            appTitle = parts.appTitle;
            sessionTitle = parts.sessionTitle;
            newChat = parts.newChatButton;
            options = parts.settingsButton;
            scopeLayout = parts.scopeLayout;
            history = parts.historyButton;
            scopePicker = parts.scopePicker;
            workflowLayout = parts.workflowLayout;
            approvalPicker = parts.approvalPicker;
            modelSummary = parts.modelSummary;
            modePicker = parts.modePicker;
            selection = parts.attachSelectionButton;
            modules = parts.attachContextButton;
            conversationPanel = parts.conversationPanel;
            historyPanel = parts.historyPanel;
            historyLayout = parts.historyLayout;
            historyLabel = parts.historyLabel;
            historySearch = parts.historySearchTextBox;
            sessionList = parts.sessionList;
            showArchived = parts.showArchivedCheckBox;
            chatTitleEditor = parts.chatTitleEditor;
            historyActions = parts.historyActions;
            rename = parts.renameChatButton;
            archive = parts.archiveChatButton;
            pin = parts.pinChatButton;
            export = parts.exportChatButton;
            deleteSession = parts.deleteChatButton;
            // A saved conversation must be selected before runtime session state can enable deletion.
            deleteSession.Enabled = false;
            memoryToggle = parts.memoryToggleButton;
            memoryPanel = parts.memoryPanel;
            memoryLayout = parts.memoryLayout;
            memoryEditor = parts.memoryEditor;
            saveMemory = parts.saveMemoryButton;
            attachMemory = parts.attachMemoryCheckBox;
            transcriptPanel = parts.TranscriptHost;
            composerLayout = parts.composerLayout;
            contextChips = parts.contextChips;
            pendingMessagesPanel = parts.PendingMessagesHost;
            promptHost = parts.inputTextBox;
            composerActions = parts.actionsLayout;
            send = parts.sendButton;
            contextToggle = parts.contextButton;
            contextPanel = parts.AdditionalContextHost;
            contextPreview = parts.AdditionalContextHost;
            providerLayout = parts.providerLayout;
            providerPicker = parts.providerPicker;
            modelPicker = parts.modelPicker;
            effortPicker = parts.effortPicker;
            footerLayout = parts.footerLayout;
            jumpToLatest = parts.latestButton;
            statusLayout = parts.footerLayout;
            status = parts.statusLabel;
            activityBar = parts.activityBar;
            appTitle.Text = "VBAi";
            parts.windowPinButton.Visible = true;
            parts.windowPinButton.Text = UiText.Get("Docked / floating window");
            parts.windowPinButton.Click += Docking_Click;
            // Resource captions are shared; these two commands expose the VBE domain.
            modules.Text = UiText.Get("# Context");
            selection.Text = UiText.Get("Attach code selection");
            promptHost.UseNativeEditor = false;
            parts.projectLayout.Visible = false;
            int projectRow = rootLayout.GetRow(parts.projectLayout);
            rootLayout.RowStyles[projectRow].SizeType = System.Windows.Forms.SizeType.Absolute;
            rootLayout.RowStyles[projectRow].Height = 0;
            parts.createProjectButton.Visible = false;
            parts.attachFileButton.Visible = false;
            parts.contextPreview.Visible = false;
            parts.welcomePanel.Visible = false;
            parts.transcriptPanel.Visible = false;
            parts.closeHistoryButton.Click += (sender, args) => SetHistoryVisible(false);
            parts.stopButton.Visible = false;
            parts.stopButton.Click += async (sender, args) => await StopTurnAsync();

            pendingMessagesPanel.ControlAdded += (sender, args) => UpdateSharedQueueVisibility();
            pendingMessagesPanel.ControlRemoved += (sender, args) => UpdateSharedQueueVisibility();
            UpdateSharedQueueVisibility();

            approvalPicker.Items.Clear();
            approvalPicker.Items.AddRange(new object[] { "Automatic", "Ask for other actions", "Read-only" });
            approvalPicker.SelectedIndex = 0;
            int providerRow = rootLayout.GetRow(providerLayout);
            rootLayout.RowStyles[providerRow].SizeType = System.Windows.Forms.SizeType.Absolute;
            rootLayout.RowStyles[providerRow].Height = 0;
            providerLayout.Visible = false;
            newChat.Click += NewChat_Click;
            modules.Click += Modules_Click;
            methods.Click += Methods_Click;
            modePicker.SelectedIndexChanged += ModePicker_SelectedIndexChanged;
            about.Click += About_Click;
            github.Click += GitHub_Click;
            docking.Click += Docking_Click;
            options.Click += Options_Click;
            history.Click += History_Click;
            approvalPicker.SelectedIndexChanged += ApprovalPicker_SelectedIndexChanged;
            verifyChanges.CheckedChanged += VerifyChanges_CheckedChanged;
            selection.Click += Selection_Click;
            compile.Click += Compile_Click;
            showArchived.CheckedChanged += ShowArchived_CheckedChanged;
            rename.Click += Rename_Click;
            archive.Click += Archive_Click;
            pin.Click += Pin_Click;
            export.Click += Export_Click;
            deleteSession.Click += DeleteSession_Click;
            memoryToggle.Click += MemoryToggle_Click;
            saveMemory.Click += SaveMemory_Click;
            attachMemory.CheckedChanged += AttachMemory_CheckedChanged;
            modelSummary.Click += ModelSummary_Click;
            contextToggle.Click += ContextToggle_Click;
            jumpToLatest.Click += JumpToLatest_Click;
        }

        private void UpdateSharedQueueVisibility()
        {
            if (workspace == null || workspace.IsDisposed || pendingMessagesPanel.IsDisposed) return;
            workspace.SetExtensionVisible(pendingMessagesPanel, pendingMessagesPanel.Controls.Count > 0);
        }

        private static string SharedGlyph(UiSymbol symbol) => symbol == UiSymbol.None ? string.Empty : ((char)symbol).ToString();

        #endregion
    }
}
