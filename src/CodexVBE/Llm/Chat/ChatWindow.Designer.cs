namespace CodexVBE
{
    /// <summary>Fenêtre de conversation et commandes de son concepteur WinForms.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Conteneur des composants WinForms dont la durée de vie est gérée par le formulaire.</summary>
        private System.ComponentModel.IContainer components = null;
        /// <summary>Gestionnaire des infobulles attachées aux commandes du formulaire.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Menu des commandes de configuration et d’intégration.</summary>
        private System.Windows.Forms.ContextMenuStrip optionsMenu;
        /// <summary>Commande qui ouvre les paramètres de l’application.</summary>
        private System.Windows.Forms.ToolStripMenuItem configure;
        private System.Windows.Forms.ToolStripMenuItem projectAccess;
        private System.Windows.Forms.ToolStripMenuItem resumeTurn;
        private System.Windows.Forms.ToolStripMenuItem about;
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
        private CodexVBE.ChatActionButton newChat;
        /// <summary>Commande d’accès aux options.</summary>
        private CodexVBE.ChatActionButton options;
        /// <summary>Disposition du sélecteur de portée de projet et de l’historique.</summary>
        private System.Windows.Forms.TableLayoutPanel scopeLayout;
        /// <summary>Commande d’affichage de l’historique.</summary>
        private CodexVBE.ChatActionButton history;
        /// <summary>Sélecteur du projet auquel la conversation est liée.</summary>
        private CodexVBE.ChatChoiceBox scopePicker;
        /// <summary>Disposition des commandes de mode et de workflow.</summary>
        private System.Windows.Forms.TableLayoutPanel workflowLayout;
        /// <summary>Sélecteur du mode de conversation.</summary>
        private CodexVBE.ChatChoiceBox modePicker;
        /// <summary>Commande de sélection de code ou de référence.</summary>
        private CodexVBE.ChatActionButton selection;
        /// <summary>Commande de compilation du projet courant.</summary>
        private CodexVBE.ChatActionButton compile;
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
        private CodexVBE.ChatActionButton rename;
        /// <summary>Commande d’archivage ou de restauration de la session.</summary>
        private CodexVBE.ChatActionButton archive;
        /// <summary>Commande d’épinglage de la session.</summary>
        private CodexVBE.ChatActionButton pin;
        /// <summary>Commande d’export de la conversation.</summary>
        private CodexVBE.ChatActionButton export;
        /// <summary>Commande d’affichage ou de masquage de la mémoire de projet.</summary>
        private CodexVBE.ChatActionButton memoryToggle;
        /// <summary>Groupe d’édition de la mémoire de projet.</summary>
        private System.Windows.Forms.GroupBox memoryPanel;
        /// <summary>Disposition du champ et des commandes de mémoire.</summary>
        private System.Windows.Forms.TableLayoutPanel memoryLayout;
        /// <summary>Champ d’édition de la mémoire de projet.</summary>
        private System.Windows.Forms.TextBox memoryEditor;
        /// <summary>Commande d’enregistrement de la mémoire de projet.</summary>
        private CodexVBE.ChatActionButton saveMemory;
        /// <summary>Option d’inclusion de la mémoire dans le prochain message.</summary>
        private System.Windows.Forms.CheckBox attachMemory;
        /// <summary>Panneau contenant le transcript de conversation.</summary>
        private System.Windows.Forms.Panel transcriptPanel;
        /// <summary>Libellé affiché avant la création de la vue du transcript.</summary>
        private System.Windows.Forms.Label transcriptPlaceholder;
        /// <summary>Hôte WinForms du transcript WPF.</summary>
        private CodexVBE.ChatContentHost transcriptHost;
        /// <summary>Conteneur de composition du message et de ses commandes.</summary>
        private CodexVBE.ChatComposerPanel composerLayout;
        /// <summary>Disposition des références et pièces jointes sélectionnées.</summary>
        private System.Windows.Forms.FlowLayoutPanel contextChips;
        private System.Windows.Forms.FlowLayoutPanel pendingMessagesPanel;
        /// <summary>Panneau du champ de saisie du message.</summary>
        private System.Windows.Forms.Panel promptPanel;
        /// <summary>Libellé indicatif du champ de saisie.</summary>
        private System.Windows.Forms.Label promptPlaceholder;
        /// <summary>Hôte WinForms du compositeur WPF.</summary>
        private CodexVBE.ChatInputView promptHost;
        /// <summary>Disposition des commandes du compositeur.</summary>
        private System.Windows.Forms.TableLayoutPanel composerActions;
        /// <summary>Commande d’ajout ou de sélection d’un module.</summary>
        private CodexVBE.ChatActionButton modules;
        /// <summary>Commande d’ajout ou de sélection d’une procédure.</summary>
        private CodexVBE.ChatActionButton methods;
        /// <summary>Commande d’envoi du message.</summary>
        private CodexVBE.ChatActionButton send;
        /// <summary>Commande d’affichage du contexte de conversation.</summary>
        private CodexVBE.ChatActionButton contextToggle;
        /// <summary>Panneau contenant l’aperçu du contexte courant.</summary>
        private System.Windows.Forms.Panel contextPanel;
        /// <summary>Disposition des éléments d’aperçu du contexte.</summary>
        private System.Windows.Forms.FlowLayoutPanel contextPreview;
        /// <summary>Disposition des sélecteurs de fournisseur, modèle et effort.</summary>
        private System.Windows.Forms.TableLayoutPanel providerLayout;
        /// <summary>Sélecteur du fournisseur LLM.</summary>
        private CodexVBE.ChatChoiceBox providerPicker;
        /// <summary>Sélecteur du modèle du fournisseur.</summary>
        private CodexVBE.ChatChoiceBox modelPicker;
        /// <summary>Sélecteur du niveau d’effort du modèle.</summary>
        private CodexVBE.ChatChoiceBox effortPicker;
        /// <summary>Disposition du pied de fenêtre.</summary>
        private System.Windows.Forms.TableLayoutPanel footerLayout;
        /// <summary>Commande de retour au dernier message du transcript.</summary>
        private CodexVBE.ChatActionButton jumpToLatest;
        /// <summary>Commande d’affichage des modifications de code de la session.</summary>
        private CodexVBE.ChatActionButton changes;
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
                if (components != null) components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>Crée les contrôles du formulaire et configure leur disposition et leurs événements.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.newChat = new CodexVBE.ChatActionButton();
            this.historySearch = new System.Windows.Forms.TextBox();
            this.modules = new CodexVBE.ChatActionButton();
            this.methods = new CodexVBE.ChatActionButton();
            this.modePicker = new CodexVBE.ChatChoiceBox();
            this.promptHost = new CodexVBE.ChatInputView();
            this.optionsMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.github = new System.Windows.Forms.ToolStripMenuItem();
            this.configure = new System.Windows.Forms.ToolStripMenuItem();
            this.projectAccess = new System.Windows.Forms.ToolStripMenuItem();
            this.resumeTurn = new System.Windows.Forms.ToolStripMenuItem();
            this.about = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.refreshModels = new System.Windows.Forms.ToolStripMenuItem();
            this.docking = new System.Windows.Forms.ToolStripMenuItem();
            this.rootLayout = new System.Windows.Forms.TableLayoutPanel();
            this.headingLayout = new System.Windows.Forms.TableLayoutPanel();
            this.appTitle = new System.Windows.Forms.Label();
            this.sessionTitle = new System.Windows.Forms.Label();
            this.options = new CodexVBE.ChatActionButton();
            this.scopeLayout = new System.Windows.Forms.TableLayoutPanel();
            this.history = new CodexVBE.ChatActionButton();
            this.scopePicker = new CodexVBE.ChatChoiceBox();
            this.workflowLayout = new System.Windows.Forms.TableLayoutPanel();
            this.selection = new CodexVBE.ChatActionButton();
            this.compile = new CodexVBE.ChatActionButton();
            this.verifyAfterEdit = new System.Windows.Forms.CheckBox();
            this.conversationPanel = new System.Windows.Forms.Panel();
            this.transcriptPanel = new System.Windows.Forms.Panel();
            this.transcriptPlaceholder = new System.Windows.Forms.Label();
            this.transcriptHost = new CodexVBE.ChatContentHost();
            this.historyPanel = new System.Windows.Forms.Panel();
            this.historyLayout = new System.Windows.Forms.TableLayoutPanel();
            this.historyLabel = new System.Windows.Forms.Label();
            this.sessionList = new System.Windows.Forms.ListBox();
            this.showArchived = new System.Windows.Forms.CheckBox();
            this.chatTitleEditor = new System.Windows.Forms.TextBox();
            this.historyActions = new System.Windows.Forms.FlowLayoutPanel();
            this.rename = new CodexVBE.ChatActionButton();
            this.archive = new CodexVBE.ChatActionButton();
            this.pin = new CodexVBE.ChatActionButton();
            this.export = new CodexVBE.ChatActionButton();
            this.memoryToggle = new CodexVBE.ChatActionButton();
            this.memoryPanel = new System.Windows.Forms.GroupBox();
            this.memoryLayout = new System.Windows.Forms.TableLayoutPanel();
            this.memoryEditor = new System.Windows.Forms.TextBox();
            this.saveMemory = new CodexVBE.ChatActionButton();
            this.attachMemory = new System.Windows.Forms.CheckBox();
            this.composerLayout = new CodexVBE.ChatComposerPanel();
            this.contextChips = new System.Windows.Forms.FlowLayoutPanel();
            this.pendingMessagesPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.promptPanel = new System.Windows.Forms.Panel();
            this.promptPlaceholder = new System.Windows.Forms.Label();
            this.composerActions = new System.Windows.Forms.TableLayoutPanel();
            this.send = new CodexVBE.ChatActionButton();
            this.contextToggle = new CodexVBE.ChatActionButton();
            this.contextPanel = new System.Windows.Forms.Panel();
            this.contextPreview = new System.Windows.Forms.FlowLayoutPanel();
            this.providerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.providerPicker = new CodexVBE.ChatChoiceBox();
            this.modelPicker = new CodexVBE.ChatChoiceBox();
            this.effortPicker = new CodexVBE.ChatChoiceBox();
            this.footerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.jumpToLatest = new CodexVBE.ChatActionButton();
            this.changes = new CodexVBE.ChatActionButton();
            this.statusLayout = new System.Windows.Forms.TableLayoutPanel();
            this.status = new System.Windows.Forms.Label();
            this.activityBar = new System.Windows.Forms.ProgressBar();
            this.optionsMenu.SuspendLayout();
            this.rootLayout.SuspendLayout();
            this.headingLayout.SuspendLayout();
            this.scopeLayout.SuspendLayout();
            this.workflowLayout.SuspendLayout();
            this.conversationPanel.SuspendLayout();
            this.historyPanel.SuspendLayout();
            this.historyLayout.SuspendLayout();
            this.historyActions.SuspendLayout();
            this.memoryPanel.SuspendLayout();
            this.memoryLayout.SuspendLayout();
            this.transcriptPanel.SuspendLayout();
            this.composerLayout.SuspendLayout();
            this.promptPanel.SuspendLayout();
            this.promptHost.SuspendLayout();
            this.composerActions.SuspendLayout();
            this.contextPanel.SuspendLayout();
            this.providerLayout.SuspendLayout();
            this.footerLayout.SuspendLayout();
            this.statusLayout.SuspendLayout();
            this.SuspendLayout();
            //
            // newChat
            //
            this.newChat.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.newChat.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.newChat.Cursor = System.Windows.Forms.Cursors.Hand;
            this.newChat.FlatAppearance.BorderSize = 0;
            this.newChat.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.newChat.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.newChat.Location = new System.Drawing.Point(417, 15);
            this.newChat.Name = "newChat";
            this.headingLayout.SetRowSpan(this.newChat, 2);
            this.newChat.Size = new System.Drawing.Size(100, 30);
            this.newChat.TabIndex = 9;
            this.newChat.Text = "+ New";
            this.toolTips.SetToolTip(this.newChat, "New conversation · Ctrl+N");
            this.newChat.UseVisualStyleBackColor = false;
            this.newChat.Click += new System.EventHandler(this.NewChat_Click);
            //
            // historySearch
            //
            this.historySearch.AccessibleName = "Search conversations";
            this.historySearch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.historySearch.Location = new System.Drawing.Point(4, 28);
            this.historySearch.Margin = new System.Windows.Forms.Padding(4);
            this.historySearch.Name = "historySearch";
            this.historySearch.Size = new System.Drawing.Size(239, 23);
            this.historySearch.TabIndex = 23;
            this.toolTips.SetToolTip(this.historySearch, "Search titles, messages or code");
            //
            // modules
            //
            this.modules.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.modules.Cursor = System.Windows.Forms.Cursors.Hand;
            this.modules.FlatAppearance.BorderSize = 0;
            this.modules.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.modules.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.modules.Location = new System.Drawing.Point(3, 3);
            this.modules.Name = "modules";
            this.modules.Size = new System.Drawing.Size(98, 30);
            this.modules.TabIndex = 47;
            this.modules.Text = "# Context";
            this.toolTips.SetToolTip(this.modules, "Reference a project or module");
            this.modules.UseVisualStyleBackColor = false;
            this.modules.Click += new System.EventHandler(this.Modules_Click);
            //
            // methods
            //
            this.methods.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.methods.Cursor = System.Windows.Forms.Cursors.Hand;
            this.methods.FlatAppearance.BorderSize = 0;
            this.methods.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.methods.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.methods.Location = new System.Drawing.Point(109, 3);
            this.methods.Name = "methods";
            this.methods.Size = new System.Drawing.Size(100, 30);
            this.methods.TabIndex = 48;
            this.methods.Text = "@ Function";
            this.toolTips.SetToolTip(this.methods, "Reference a Sub, Function or Property");
            this.methods.UseVisualStyleBackColor = false;
            this.methods.Click += new System.EventHandler(this.Methods_Click);
            //
            // modePicker
            //
            this.modePicker.AccessibleName = "Working mode";
            this.modePicker.BackColor = System.Drawing.Color.White;
            this.modePicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modePicker.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.modePicker.DropDownHeight = 280;
            this.modePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modePicker.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.modePicker.IntegralHeight = false;
            this.modePicker.ItemHeight = 22;
            this.modePicker.Location = new System.Drawing.Point(4, 6);
            this.modePicker.Margin = new System.Windows.Forms.Padding(4, 6, 4, 3);
            this.modePicker.Name = "modePicker";
            this.modePicker.Size = new System.Drawing.Size(92, 28);
            this.modePicker.TabIndex = 15;
            this.toolTips.SetToolTip(this.modePicker, "Chat and Plan: no project changes");
            this.modePicker.SelectedIndexChanged += new System.EventHandler(this.ModePicker_SelectedIndexChanged);
            //
            // promptHost
            //
            this.promptHost.BackColor = System.Drawing.Color.White;
            this.promptHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.promptHost.Location = new System.Drawing.Point(0, 0);
            this.promptHost.Margin = new System.Windows.Forms.Padding(0);
            this.promptHost.Name = "promptHost";
            this.promptHost.Size = new System.Drawing.Size(552, 88);
            this.promptHost.TabIndex = 45;
            this.toolTips.SetToolTip(this.promptHost, "Message with spell checking");
            this.promptHost.Font = new System.Drawing.Font("Segoe UI", 10.5F);
            //
            // optionsMenu
            //
            this.optionsMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.configure,
            this.projectAccess,
            this.resumeTurn,
            this.refreshModels,
            this.docking,
            this.github,
            this.aboutSeparator,
            this.about});
            this.optionsMenu.Name = "optionsMenu";
            this.aboutSeparator.Name = "aboutSeparator";
            this.about.Name = "about";
            this.about.Text = "About VBAi";
            this.about.ToolTipText = "Version, technical details and project resources";
            this.about.Click += new System.EventHandler(this.About_Click);
            this.github.Name = "github";
            this.github.Text = "GitHub · synchronize VBA…";
            this.github.ToolTipText = "Export and synchronize the current document's sources with a GitHub repository.";
            this.github.Click += new System.EventHandler(this.GitHub_Click);
            this.optionsMenu.Size = new System.Drawing.Size(222, 70);
            //
            // configure
            //
            this.configure.Name = "configure";
            this.configure.Size = new System.Drawing.Size(221, 22);
            this.configure.Text = "Settings…";
            this.projectAccess.Name = "projectAccess";
            this.projectAccess.Text = "Project access…";
            this.projectAccess.ToolTipText = "Choose which other projects this conversation may read.";
            this.resumeTurn.Name = "resumeTurn";
            this.resumeTurn.Text = "Resume paused turn";
            this.resumeTurn.ToolTipText = "Continue from saved tool results without repeating completed actions.";
            //
            // refreshModels
            //
            this.refreshModels.Name = "refreshModels";
            this.refreshModels.Size = new System.Drawing.Size(221, 22);
            this.refreshModels.Text = "Refresh models";
            //
            // docking
            //
            this.docking.Name = "docking";
            this.docking.Size = new System.Drawing.Size(221, 22);
            this.docking.Text = "Docked / floating window";
            this.docking.Click += new System.EventHandler(this.Docking_Click);
            //
            // rootLayout
            //
            this.rootLayout.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.rootLayout.ColumnCount = 1;
            this.rootLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.Controls.Add(this.headingLayout, 0, 0);
            this.rootLayout.Controls.Add(this.scopeLayout, 0, 1);
            this.rootLayout.Controls.Add(this.workflowLayout, 0, 2);
            this.rootLayout.Controls.Add(this.verifyAfterEdit, 0, 3);
            this.rootLayout.Controls.Add(this.conversationPanel, 0, 4);
            this.rootLayout.Controls.Add(this.composerLayout, 0, 5);
            this.rootLayout.Controls.Add(this.providerLayout, 0, 6);
            this.rootLayout.Controls.Add(this.footerLayout, 0, 7);
            this.rootLayout.Controls.Add(this.statusLayout, 0, 8);
            this.rootLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rootLayout.Location = new System.Drawing.Point(0, 0);
            this.rootLayout.Margin = new System.Windows.Forms.Padding(0);
            this.rootLayout.Name = "rootLayout";
            this.rootLayout.Padding = new System.Windows.Forms.Padding(18, 8, 18, 10);
            this.rootLayout.RowCount = 9;
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.rootLayout.Size = new System.Drawing.Size(600, 820);
            this.rootLayout.TabIndex = 5;
            //
            // headingLayout
            //
            this.headingLayout.ColumnCount = 3;
            this.headingLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.headingLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 108F));
            this.headingLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.headingLayout.Controls.Add(this.appTitle, 0, 0);
            this.headingLayout.Controls.Add(this.sessionTitle, 0, 1);
            this.headingLayout.Controls.Add(this.newChat, 1, 0);
            this.headingLayout.Controls.Add(this.options, 2, 0);
            this.headingLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.headingLayout.Location = new System.Drawing.Point(18, 8);
            this.headingLayout.Margin = new System.Windows.Forms.Padding(0);
            this.headingLayout.Name = "headingLayout";
            this.headingLayout.RowCount = 2;
            this.headingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 55F));
            this.headingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45F));
            this.headingLayout.Size = new System.Drawing.Size(564, 60);
            this.headingLayout.TabIndex = 6;
            //
            // appTitle
            //
            this.appTitle.AutoEllipsis = true;
            this.appTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.appTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 14.25F);
            this.appTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.appTitle.Location = new System.Drawing.Point(4, 0);
            this.appTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.appTitle.Name = "appTitle";
            this.appTitle.Size = new System.Drawing.Size(404, 33);
            this.appTitle.TabIndex = 7;
            this.appTitle.Text = "VBAi";
            this.appTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // sessionTitle
            //
            this.sessionTitle.AutoEllipsis = true;
            this.sessionTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sessionTitle.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.sessionTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.sessionTitle.Location = new System.Drawing.Point(4, 33);
            this.sessionTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.sessionTitle.Name = "sessionTitle";
            this.sessionTitle.Size = new System.Drawing.Size(404, 27);
            this.sessionTitle.TabIndex = 8;
            this.sessionTitle.Text = "Your AI agent for VBA";
            this.sessionTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // options
            //
            this.options.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.options.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.options.Cursor = System.Windows.Forms.Cursors.Hand;
            this.options.FlatAppearance.BorderSize = 0;
            this.options.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.options.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.options.Location = new System.Drawing.Point(523, 15);
            this.options.Name = "options";
            this.toolTips.SetToolTip(this.options, "Configure the provider, refresh models or change docking mode.");
            this.headingLayout.SetRowSpan(this.options, 2);
            this.options.Size = new System.Drawing.Size(38, 30);
            this.options.TabIndex = 10;
            this.options.Text = "···";
            this.options.UseVisualStyleBackColor = false;
            this.options.Click += new System.EventHandler(this.Options_Click);
            //
            // scopeLayout
            //
            this.scopeLayout.ColumnCount = 2;
            this.scopeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 86F));
            this.scopeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.scopeLayout.Controls.Add(this.history, 0, 0);
            this.scopeLayout.Controls.Add(this.scopePicker, 1, 0);
            this.scopeLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scopeLayout.Location = new System.Drawing.Point(18, 68);
            this.scopeLayout.Margin = new System.Windows.Forms.Padding(0);
            this.scopeLayout.Name = "scopeLayout";
            this.scopeLayout.RowCount = 1;
            this.scopeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.scopeLayout.Size = new System.Drawing.Size(564, 38);
            this.scopeLayout.TabIndex = 11;
            //
            // history
            //
            this.history.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.history.Cursor = System.Windows.Forms.Cursors.Hand;
            this.history.FlatAppearance.BorderSize = 0;
            this.history.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.history.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.history.Location = new System.Drawing.Point(3, 3);
            this.history.Name = "history";
            this.toolTips.SetToolTip(this.history, "Show or hide conversations for the selected document.");
            this.history.Size = new System.Drawing.Size(78, 30);
            this.history.TabIndex = 12;
            this.history.Text = "☰ Chats";
            this.history.UseVisualStyleBackColor = false;
            this.history.Click += new System.EventHandler(this.History_Click);
            //
            // scopePicker
            //
            this.scopePicker.AccessibleName = "VBA project · document history";
            this.scopePicker.BackColor = System.Drawing.Color.White;
            this.scopePicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scopePicker.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.scopePicker.DropDownHeight = 280;
            this.scopePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.scopePicker.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.scopePicker.IntegralHeight = false;
            this.scopePicker.ItemHeight = 22;
            this.scopePicker.Location = new System.Drawing.Point(90, 6);
            this.scopePicker.Margin = new System.Windows.Forms.Padding(4, 6, 4, 3);
            this.scopePicker.Name = "scopePicker";
            this.toolTips.SetToolTip(this.scopePicker, "Choose the VBA document associated with this context and its conversations.");
            this.scopePicker.Size = new System.Drawing.Size(470, 28);
            this.scopePicker.TabIndex = 13;
            //
            // workflowLayout
            //
            this.workflowLayout.ColumnCount = 4;
            this.workflowLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.workflowLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.workflowLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 112F));
            this.workflowLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.workflowLayout.Controls.Add(this.modePicker, 0, 0);
            this.workflowLayout.Controls.Add(this.selection, 1, 0);
            this.workflowLayout.Controls.Add(this.compile, 2, 0);
            this.workflowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.workflowLayout.Location = new System.Drawing.Point(18, 106);
            this.workflowLayout.Margin = new System.Windows.Forms.Padding(0);
            this.workflowLayout.Name = "workflowLayout";
            this.workflowLayout.RowCount = 1;
            this.workflowLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.workflowLayout.Size = new System.Drawing.Size(564, 38);
            this.workflowLayout.TabIndex = 14;
            //
            // selection
            //
            this.selection.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.selection.Cursor = System.Windows.Forms.Cursors.Hand;
            this.selection.FlatAppearance.BorderSize = 0;
            this.selection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.selection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.selection.Location = new System.Drawing.Point(103, 3);
            this.selection.Name = "selection";
            this.toolTips.SetToolTip(this.selection, "Attach the code selected in the VBE; attach the current line when nothing is selected.");
            this.selection.Size = new System.Drawing.Size(152, 30);
            this.selection.TabIndex = 16;
            this.selection.Text = "Attach selection";
            this.selection.UseVisualStyleBackColor = false;
            this.selection.Click += new System.EventHandler(this.Selection_Click);
            //
            // compile
            //
            this.compile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.compile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.compile.FlatAppearance.BorderSize = 0;
            this.compile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.compile.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.compile.Location = new System.Drawing.Point(263, 3);
            this.compile.Name = "compile";
            this.toolTips.SetToolTip(this.compile, "Compile the VBA project and display diagnostics in the chat.");
            this.compile.Size = new System.Drawing.Size(104, 30);
            this.compile.TabIndex = 17;
            this.compile.Text = "Check VBA";
            this.compile.UseVisualStyleBackColor = false;
            this.compile.Click += new System.EventHandler(this.Compile_Click);
            //
            // verifyAfterEdit
            //
            this.verifyAfterEdit.AutoSize = true;
            this.verifyAfterEdit.Checked = true;
            this.verifyAfterEdit.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verifyAfterEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.verifyAfterEdit.Location = new System.Drawing.Point(22, 147);
            this.verifyAfterEdit.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.verifyAfterEdit.Name = "verifyAfterEdit";
            this.toolTips.SetToolTip(this.verifyAfterEdit, "Automatically compile the project after agent edits.");
            this.verifyAfterEdit.Size = new System.Drawing.Size(556, 22);
            this.verifyAfterEdit.TabIndex = 18;
            this.verifyAfterEdit.Text = "Compile after changes";
            //
            // conversationPanel
            //
            this.conversationPanel.Controls.Add(this.transcriptPanel);
            this.conversationPanel.Controls.Add(this.historyPanel);
            this.conversationPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conversationPanel.Location = new System.Drawing.Point(18, 176);
            this.conversationPanel.Margin = new System.Windows.Forms.Padding(0, 4, 0, 4);
            this.conversationPanel.Name = "conversationPanel";
            this.conversationPanel.Size = new System.Drawing.Size(564, 236);
            this.conversationPanel.TabIndex = 19;
            //
            // transcriptPanel
            //
            this.transcriptPanel.BackColor = System.Drawing.Color.White;
            this.transcriptPanel.Controls.Add(this.transcriptPlaceholder);
            this.transcriptPanel.Controls.Add(this.transcriptHost);
            this.transcriptPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transcriptPanel.Location = new System.Drawing.Point(0, 0);
            this.transcriptPanel.Margin = new System.Windows.Forms.Padding(0);
            this.transcriptPanel.Name = "transcriptPanel";
            this.transcriptPanel.Size = new System.Drawing.Size(564, 236);
            this.transcriptPanel.TabIndex = 38;
            //
            // transcriptPlaceholder
            //
            this.transcriptPlaceholder.AutoEllipsis = true;
            this.transcriptPlaceholder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transcriptPlaceholder.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.transcriptPlaceholder.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.transcriptPlaceholder.Location = new System.Drawing.Point(0, 0);
            this.transcriptPlaceholder.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.transcriptPlaceholder.Name = "transcriptPlaceholder";
            this.transcriptPlaceholder.Size = new System.Drawing.Size(564, 236);
            this.transcriptPlaceholder.TabIndex = 39;
            this.transcriptPlaceholder.Text = "Conversation: messages, references and dynamic diffs";
            this.transcriptPlaceholder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // transcriptHost
            //
            this.transcriptHost.BackColor = System.Drawing.Color.White;
            this.transcriptHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transcriptHost.Location = new System.Drawing.Point(0, 0);
            this.transcriptHost.Margin = new System.Windows.Forms.Padding(0);
            this.transcriptHost.Name = "transcriptHost";
            this.transcriptHost.Size = new System.Drawing.Size(564, 236);
            this.transcriptHost.TabIndex = 40;
            this.transcriptHost.Child = null;
            //
            // historyPanel
            //
            this.historyPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)));
            this.historyPanel.AutoScroll = true;
            this.historyPanel.BackColor = System.Drawing.Color.White;
            this.historyPanel.Controls.Add(this.historyLayout);
            this.historyPanel.Location = new System.Drawing.Point(0, 0);
            this.historyPanel.Name = "historyPanel";
            this.historyPanel.Padding = new System.Windows.Forms.Padding(6);
            this.historyPanel.Size = new System.Drawing.Size(276, 236);
            this.historyPanel.TabIndex = 20;
            this.historyPanel.Visible = false;
            //
            // historyLayout
            //
            this.historyLayout.AutoSize = true;
            this.historyLayout.ColumnCount = 1;
            this.historyLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.historyLayout.Controls.Add(this.historyLabel, 0, 0);
            this.historyLayout.Controls.Add(this.historySearch, 0, 1);
            this.historyLayout.Controls.Add(this.sessionList, 0, 2);
            this.historyLayout.Controls.Add(this.showArchived, 0, 3);
            this.historyLayout.Controls.Add(this.chatTitleEditor, 0, 4);
            this.historyLayout.Controls.Add(this.historyActions, 0, 5);
            this.historyLayout.Controls.Add(this.memoryToggle, 0, 6);
            this.historyLayout.Controls.Add(this.memoryPanel, 0, 7);
            this.historyLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.historyLayout.Location = new System.Drawing.Point(6, 6);
            this.historyLayout.Margin = new System.Windows.Forms.Padding(0);
            this.historyLayout.Name = "historyLayout";
            this.historyLayout.RowCount = 8;
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 128F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.historyLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 226F));
            this.historyLayout.Size = new System.Drawing.Size(247, 642);
            this.historyLayout.TabIndex = 21;
            //
            // historyLabel
            //
            this.historyLabel.AutoEllipsis = true;
            this.historyLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.historyLabel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.historyLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.historyLabel.Location = new System.Drawing.Point(4, 0);
            this.historyLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.historyLabel.Name = "historyLabel";
            this.historyLabel.Size = new System.Drawing.Size(239, 24);
            this.historyLabel.TabIndex = 22;
            this.historyLabel.Text = "DOCUMENT CONVERSATIONS";
            this.historyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // sessionList
            //
            this.sessionList.AccessibleName = "Document conversations";
            this.sessionList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.sessionList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sessionList.HorizontalScrollbar = true;
            this.sessionList.IntegralHeight = false;
            this.sessionList.ItemHeight = 15;
            this.sessionList.Location = new System.Drawing.Point(4, 60);
            this.sessionList.Margin = new System.Windows.Forms.Padding(4);
            this.sessionList.Name = "sessionList";
            this.toolTips.SetToolTip(this.sessionList, "Select a conversation to resume its history and draft.");
            this.sessionList.Size = new System.Drawing.Size(239, 120);
            this.sessionList.TabIndex = 24;
            //
            // showArchived
            //
            this.showArchived.AutoSize = true;
            this.showArchived.Dock = System.Windows.Forms.DockStyle.Fill;
            this.showArchived.Location = new System.Drawing.Point(4, 187);
            this.showArchived.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.showArchived.Name = "showArchived";
            this.toolTips.SetToolTip(this.showArchived, "Include archived conversations in the list.");
            this.showArchived.Size = new System.Drawing.Size(239, 20);
            this.showArchived.TabIndex = 25;
            this.showArchived.Text = "Show archived conversations";
            this.showArchived.CheckedChanged += new System.EventHandler(this.ShowArchived_CheckedChanged);
            //
            // chatTitleEditor
            //
            this.chatTitleEditor.AccessibleName = "Conversation title";
            this.chatTitleEditor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chatTitleEditor.Location = new System.Drawing.Point(4, 214);
            this.chatTitleEditor.Margin = new System.Windows.Forms.Padding(4);
            this.chatTitleEditor.Name = "chatTitleEditor";
            this.toolTips.SetToolTip(this.chatTitleEditor, "Enter a title, then click Rename.");
            this.chatTitleEditor.Size = new System.Drawing.Size(239, 23);
            this.chatTitleEditor.TabIndex = 26;
            //
            // historyActions
            //
            this.historyActions.AutoSize = true;
            this.historyActions.Controls.Add(this.rename);
            this.historyActions.Controls.Add(this.archive);
            this.historyActions.Controls.Add(this.pin);
            this.historyActions.Controls.Add(this.export);
            this.historyActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.historyActions.Location = new System.Drawing.Point(0, 242);
            this.historyActions.Margin = new System.Windows.Forms.Padding(0);
            this.historyActions.Name = "historyActions";
            this.historyActions.Size = new System.Drawing.Size(247, 144);
            this.historyActions.TabIndex = 27;
            //
            // rename
            //
            this.rename.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.rename.Cursor = System.Windows.Forms.Cursors.Hand;
            this.rename.FlatAppearance.BorderSize = 0;
            this.rename.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.rename.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.rename.Location = new System.Drawing.Point(3, 3);
            this.rename.Name = "rename";
            this.toolTips.SetToolTip(this.rename, "Save the new title of the selected conversation.");
            this.rename.Size = new System.Drawing.Size(90, 30);
            this.rename.TabIndex = 28;
            this.rename.Text = "Rename";
            this.rename.UseVisualStyleBackColor = false;
            this.rename.Click += new System.EventHandler(this.Rename_Click);
            //
            // archive
            //
            this.archive.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.archive.Cursor = System.Windows.Forms.Cursors.Hand;
            this.archive.FlatAppearance.BorderSize = 0;
            this.archive.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.archive.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.archive.Location = new System.Drawing.Point(3, 39);
            this.archive.Name = "archive";
            this.toolTips.SetToolTip(this.archive, "Archive the selected conversation or restore it if archived.");
            this.archive.Size = new System.Drawing.Size(158, 30);
            this.archive.TabIndex = 29;
            this.archive.Text = "Archive / restore";
            this.archive.UseVisualStyleBackColor = false;
            this.archive.Click += new System.EventHandler(this.Archive_Click);
            //
            // pin
            //
            this.pin.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.pin.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pin.FlatAppearance.BorderSize = 0;
            this.pin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.pin.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.pin.Location = new System.Drawing.Point(3, 75);
            this.pin.Name = "pin";
            this.toolTips.SetToolTip(this.pin, "Pin the conversation to the top of the list or unpin it.");
            this.pin.Size = new System.Drawing.Size(126, 30);
            this.pin.TabIndex = 30;
            this.pin.Text = "Pin / unpin";
            this.pin.UseVisualStyleBackColor = false;
            this.pin.Click += new System.EventHandler(this.Pin_Click);
            //
            // export
            //
            this.export.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.export.Cursor = System.Windows.Forms.Cursors.Hand;
            this.export.FlatAppearance.BorderSize = 0;
            this.export.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.export.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.export.Location = new System.Drawing.Point(3, 111);
            this.export.Name = "export";
            this.toolTips.SetToolTip(this.export, "Save the conversation as a Markdown file.");
            this.export.Size = new System.Drawing.Size(122, 30);
            this.export.TabIndex = 31;
            this.export.Text = "Export Markdown";
            this.export.UseVisualStyleBackColor = false;
            this.export.Click += new System.EventHandler(this.Export_Click);
            //
            // memoryToggle
            //
            this.memoryToggle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.memoryToggle.Cursor = System.Windows.Forms.Cursors.Hand;
            this.memoryToggle.FlatAppearance.BorderSize = 0;
            this.memoryToggle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.memoryToggle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.memoryToggle.Location = new System.Drawing.Point(3, 389);
            this.memoryToggle.Name = "memoryToggle";
            this.toolTips.SetToolTip(this.memoryToggle, "Show or hide local notes for this document.");
            this.memoryToggle.Size = new System.Drawing.Size(241, 24);
            this.memoryToggle.TabIndex = 32;
            this.memoryToggle.Text = "Document memory";
            this.memoryToggle.UseVisualStyleBackColor = false;
            this.memoryToggle.Click += new System.EventHandler(this.MemoryToggle_Click);
            //
            // memoryPanel
            //
            this.memoryPanel.Controls.Add(this.memoryLayout);
            this.memoryPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.memoryPanel.Location = new System.Drawing.Point(4, 420);
            this.memoryPanel.Margin = new System.Windows.Forms.Padding(4);
            this.memoryPanel.Name = "memoryPanel";
            this.memoryPanel.Size = new System.Drawing.Size(239, 218);
            this.memoryPanel.TabIndex = 33;
            this.memoryPanel.TabStop = false;
            this.memoryPanel.Text = "Local document notes";
            //
            // memoryLayout
            //
            this.memoryLayout.ColumnCount = 1;
            this.memoryLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.memoryLayout.Controls.Add(this.memoryEditor, 0, 0);
            this.memoryLayout.Controls.Add(this.saveMemory, 0, 1);
            this.memoryLayout.Controls.Add(this.attachMemory, 0, 2);
            this.memoryLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.memoryLayout.Location = new System.Drawing.Point(3, 19);
            this.memoryLayout.Margin = new System.Windows.Forms.Padding(0);
            this.memoryLayout.Name = "memoryLayout";
            this.memoryLayout.RowCount = 3;
            this.memoryLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.memoryLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.memoryLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 65F));
            this.memoryLayout.Size = new System.Drawing.Size(233, 196);
            this.memoryLayout.TabIndex = 34;
            //
            // memoryEditor
            //
            this.memoryEditor.AccessibleName = "Document memory";
            this.memoryEditor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.memoryEditor.Location = new System.Drawing.Point(4, 4);
            this.memoryEditor.Margin = new System.Windows.Forms.Padding(4);
            this.memoryEditor.MaxLength = 16000;
            this.memoryEditor.Multiline = true;
            this.memoryEditor.Name = "memoryEditor";
            this.toolTips.SetToolTip(this.memoryEditor, "Enter notes to keep for this document, then save them.");
            this.memoryEditor.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.memoryEditor.Size = new System.Drawing.Size(225, 87);
            this.memoryEditor.TabIndex = 35;
            //
            // saveMemory
            //
            this.saveMemory.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.saveMemory.Cursor = System.Windows.Forms.Cursors.Hand;
            this.saveMemory.FlatAppearance.BorderSize = 0;
            this.saveMemory.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveMemory.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.saveMemory.Location = new System.Drawing.Point(3, 98);
            this.saveMemory.Name = "saveMemory";
            this.toolTips.SetToolTip(this.saveMemory, "Save notes on this computer without sending them to the provider.");
            this.saveMemory.Size = new System.Drawing.Size(180, 30);
            this.saveMemory.TabIndex = 36;
            this.saveMemory.Text = "Save locally";
            this.saveMemory.UseVisualStyleBackColor = false;
            this.saveMemory.Click += new System.EventHandler(this.SaveMemory_Click);
            //
            // attachMemory
            //
            this.attachMemory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.attachMemory.Location = new System.Drawing.Point(4, 134);
            this.attachMemory.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.attachMemory.Name = "attachMemory";
            this.toolTips.SetToolTip(this.attachMemory, "Include saved notes in the next message sent to the provider.");
            this.attachMemory.Size = new System.Drawing.Size(225, 59);
            this.attachMemory.TabIndex = 37;
            this.attachMemory.Text = "Attach these notes to the next message sent to the selected provider";
            this.attachMemory.CheckedChanged += new System.EventHandler(this.AttachMemory_CheckedChanged);
            //
            // composerLayout
            //
            this.composerLayout.AutoSize = true;
            this.composerLayout.BackColor = System.Drawing.Color.White;
            this.composerLayout.ColumnCount = 1;
            this.composerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.composerLayout.Controls.Add(this.pendingMessagesPanel, 0, 1);
            this.composerLayout.Controls.Add(this.contextChips, 0, 0);
            this.composerLayout.Controls.Add(this.promptPanel, 0, 2);
            this.composerLayout.Controls.Add(this.composerActions, 0, 3);
            this.composerLayout.Controls.Add(this.contextToggle, 0, 4);
            this.composerLayout.Controls.Add(this.contextPanel, 0, 5);
            this.composerLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composerLayout.Location = new System.Drawing.Point(18, 416);
            this.composerLayout.Margin = new System.Windows.Forms.Padding(0);
            this.composerLayout.Name = "composerLayout";
            this.composerLayout.Padding = new System.Windows.Forms.Padding(6);
            this.composerLayout.RowCount = 6;
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 88F));
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.composerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.composerLayout.Size = new System.Drawing.Size(564, 294);
            this.composerLayout.TabIndex = 41;
            //
            // pendingMessagesPanel
            //
            this.pendingMessagesPanel.AutoSize = true;
            this.pendingMessagesPanel.AutoScroll = true;
            this.pendingMessagesPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pendingMessagesPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.pendingMessagesPanel.WrapContents = false;
            this.pendingMessagesPanel.MaximumSize = new System.Drawing.Size(0, 144);
            this.pendingMessagesPanel.Margin = new System.Windows.Forms.Padding(0);
            this.pendingMessagesPanel.Name = "pendingMessagesPanel";
            this.pendingMessagesPanel.Visible = false;
            //
            // contextChips
            //
            this.contextChips.AutoScroll = true;
            this.contextChips.AutoSize = true;
            this.contextChips.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contextChips.Location = new System.Drawing.Point(6, 6);
            this.contextChips.Margin = new System.Windows.Forms.Padding(0);
            this.contextChips.MaximumSize = new System.Drawing.Size(0, 76);
            this.contextChips.Name = "contextChips";
            this.contextChips.Size = new System.Drawing.Size(552, 1);
            this.contextChips.TabIndex = 42;
            this.contextChips.Visible = false;
            //
            // promptPanel
            //
            this.promptPanel.Controls.Add(this.promptPlaceholder);
            this.promptPanel.Controls.Add(this.promptHost);
            this.promptPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.promptPanel.Location = new System.Drawing.Point(6, 6);
            this.promptPanel.Margin = new System.Windows.Forms.Padding(0);
            this.promptPanel.Name = "promptPanel";
            this.promptPanel.Size = new System.Drawing.Size(552, 88);
            this.promptPanel.TabIndex = 43;
            //
            // promptPlaceholder
            //
            this.promptPlaceholder.AutoEllipsis = true;
            this.promptPlaceholder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.promptPlaceholder.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.promptPlaceholder.ForeColor = System.Drawing.Color.SlateGray;
            this.promptPlaceholder.Location = new System.Drawing.Point(0, 0);
            this.promptPlaceholder.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.promptPlaceholder.Name = "promptPlaceholder";
            this.promptPlaceholder.Size = new System.Drawing.Size(552, 88);
            this.promptPlaceholder.TabIndex = 44;
            this.promptPlaceholder.Text = "Your message…";
            this.promptPlaceholder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // composerActions
            //
            this.composerActions.ColumnCount = 4;
            this.composerActions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 106F));
            this.composerActions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 108F));
            this.composerActions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.composerActions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 98F));
            this.composerActions.Controls.Add(this.modules, 0, 0);
            this.composerActions.Controls.Add(this.methods, 1, 0);
            this.composerActions.Controls.Add(this.send, 3, 0);
            this.composerActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composerActions.Location = new System.Drawing.Point(6, 94);
            this.composerActions.Margin = new System.Windows.Forms.Padding(0);
            this.composerActions.Name = "composerActions";
            this.composerActions.RowCount = 1;
            this.composerActions.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.composerActions.Size = new System.Drawing.Size(552, 36);
            this.composerActions.TabIndex = 46;
            //
            // send
            //
            this.send.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.send.Cursor = System.Windows.Forms.Cursors.Hand;
            this.send.FlatAppearance.BorderSize = 0;
            this.send.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.send.ForeColor = System.Drawing.Color.White;
            this.send.Location = new System.Drawing.Point(457, 3);
            this.send.Name = "send";
            this.toolTips.SetToolTip(this.send, "Send the message and its context to the agent.");
            this.send.Size = new System.Drawing.Size(90, 30);
            this.send.TabIndex = 49;
            this.send.Text = "Send ↑";
            this.send.UseVisualStyleBackColor = false;
            //
            // contextToggle
            //
            this.contextToggle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.contextToggle.Cursor = System.Windows.Forms.Cursors.Hand;
            this.contextToggle.FlatAppearance.BorderSize = 0;
            this.contextToggle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.contextToggle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.contextToggle.Location = new System.Drawing.Point(9, 133);
            this.contextToggle.Name = "contextToggle";
            this.toolTips.SetToolTip(this.contextToggle, "Inspect and refresh references, selected code and notes to send.");
            this.contextToggle.Size = new System.Drawing.Size(342, 24);
            this.contextToggle.TabIndex = 50;
            this.contextToggle.Text = "Outgoing context · inspect and refresh";
            this.contextToggle.UseVisualStyleBackColor = false;
            this.contextToggle.Click += new System.EventHandler(this.ContextToggle_Click);
            //
            // contextPanel
            //
            this.contextPanel.AutoScroll = true;
            this.contextPanel.Controls.Add(this.contextPreview);
            this.contextPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contextPanel.Location = new System.Drawing.Point(6, 160);
            this.contextPanel.Margin = new System.Windows.Forms.Padding(0);
            this.contextPanel.MinimumSize = new System.Drawing.Size(0, 128);
            this.contextPanel.Name = "contextPanel";
            this.contextPanel.Size = new System.Drawing.Size(552, 128);
            this.contextPanel.TabIndex = 51;
            this.contextPanel.Visible = false;
            //
            // contextPreview
            //
            this.contextPreview.AutoScroll = true;
            this.contextPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contextPreview.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.contextPreview.Location = new System.Drawing.Point(0, 0);
            this.contextPreview.Name = "contextPreview";
            this.contextPreview.Size = new System.Drawing.Size(552, 128);
            this.contextPreview.TabIndex = 52;
            this.contextPreview.WrapContents = false;
            //
            // providerLayout
            //
            this.providerLayout.ColumnCount = 3;
            this.providerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28F));
            this.providerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44F));
            this.providerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28F));
            this.providerLayout.Controls.Add(this.providerPicker, 0, 0);
            this.providerLayout.Controls.Add(this.modelPicker, 1, 0);
            this.providerLayout.Controls.Add(this.effortPicker, 2, 0);
            this.providerLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.providerLayout.Location = new System.Drawing.Point(18, 710);
            this.providerLayout.Margin = new System.Windows.Forms.Padding(0);
            this.providerLayout.Name = "providerLayout";
            this.providerLayout.RowCount = 1;
            this.providerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.providerLayout.Size = new System.Drawing.Size(564, 40);
            this.providerLayout.TabIndex = 53;
            //
            // providerPicker
            //
            this.providerPicker.AccessibleName = "Provider";
            this.providerPicker.BackColor = System.Drawing.Color.White;
            this.providerPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.providerPicker.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.providerPicker.DropDownHeight = 280;
            this.providerPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.providerPicker.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.providerPicker.IntegralHeight = false;
            this.providerPicker.ItemHeight = 22;
            this.providerPicker.Location = new System.Drawing.Point(4, 6);
            this.providerPicker.Margin = new System.Windows.Forms.Padding(4, 6, 4, 3);
            this.providerPicker.Name = "providerPicker";
            this.toolTips.SetToolTip(this.providerPicker, "Choose the AI provider for this conversation.");
            this.providerPicker.Size = new System.Drawing.Size(149, 28);
            this.providerPicker.TabIndex = 54;
            //
            // modelPicker
            //
            this.modelPicker.AccessibleName = "Model";
            this.modelPicker.BackColor = System.Drawing.Color.White;
            this.modelPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modelPicker.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.modelPicker.DropDownHeight = 280;
            this.modelPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modelPicker.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.modelPicker.IntegralHeight = false;
            this.modelPicker.ItemHeight = 22;
            this.modelPicker.Location = new System.Drawing.Point(161, 6);
            this.modelPicker.Margin = new System.Windows.Forms.Padding(4, 6, 4, 3);
            this.modelPicker.Name = "modelPicker";
            this.toolTips.SetToolTip(this.modelPicker, "Choose a model offered by the selected provider.");
            this.modelPicker.Size = new System.Drawing.Size(240, 28);
            this.modelPicker.TabIndex = 55;
            //
            // effortPicker
            //
            this.effortPicker.AccessibleName = "Reasoning effort";
            this.effortPicker.BackColor = System.Drawing.Color.White;
            this.effortPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.effortPicker.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.effortPicker.DropDownHeight = 280;
            this.effortPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.effortPicker.Enabled = false;
            this.effortPicker.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.effortPicker.IntegralHeight = false;
            this.effortPicker.ItemHeight = 22;
            this.effortPicker.Location = new System.Drawing.Point(409, 6);
            this.effortPicker.Margin = new System.Windows.Forms.Padding(4, 6, 4, 3);
            this.effortPicker.Name = "effortPicker";
            this.toolTips.SetToolTip(this.effortPicker, "Set the model's reasoning effort, when supported.");
            this.effortPicker.Size = new System.Drawing.Size(151, 28);
            this.effortPicker.TabIndex = 56;
            //
            // footerLayout
            //
            this.footerLayout.ColumnCount = 2;
            this.footerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.footerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 154F));
            this.footerLayout.Controls.Add(this.jumpToLatest, 0, 0);
            this.footerLayout.Controls.Add(this.changes, 1, 0);
            this.footerLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.footerLayout.Location = new System.Drawing.Point(18, 750);
            this.footerLayout.Margin = new System.Windows.Forms.Padding(0);
            this.footerLayout.Name = "footerLayout";
            this.footerLayout.RowCount = 1;
            this.footerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.footerLayout.Size = new System.Drawing.Size(564, 36);
            this.footerLayout.TabIndex = 57;
            //
            // jumpToLatest
            //
            this.jumpToLatest.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.jumpToLatest.Cursor = System.Windows.Forms.Cursors.Hand;
            this.jumpToLatest.FlatAppearance.BorderSize = 0;
            this.jumpToLatest.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.jumpToLatest.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.jumpToLatest.Location = new System.Drawing.Point(3, 3);
            this.jumpToLatest.Name = "jumpToLatest";
            this.toolTips.SetToolTip(this.jumpToLatest, "Go to the latest message and follow new responses.");
            this.jumpToLatest.Size = new System.Drawing.Size(148, 30);
            this.jumpToLatest.TabIndex = 58;
            this.jumpToLatest.Text = "↓ Latest message";
            this.jumpToLatest.UseVisualStyleBackColor = false;
            this.jumpToLatest.Click += new System.EventHandler(this.JumpToLatest_Click);
            //
            // changes
            //
            this.changes.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.changes.Cursor = System.Windows.Forms.Cursors.Hand;
            this.changes.Enabled = false;
            this.changes.FlatAppearance.BorderSize = 0;
            this.changes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.changes.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.changes.Location = new System.Drawing.Point(413, 3);
            this.changes.Name = "changes";
            this.toolTips.SetToolTip(this.changes, "View code changes and their undo actions.");
            this.changes.Size = new System.Drawing.Size(146, 30);
            this.changes.TabIndex = 59;
            this.changes.Text = "Changes · 0";
            this.changes.UseVisualStyleBackColor = false;
            //
            // statusLayout
            //
            this.statusLayout.ColumnCount = 2;
            this.statusLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.statusLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.statusLayout.Controls.Add(this.status, 0, 0);
            this.statusLayout.Controls.Add(this.activityBar, 1, 0);
            this.statusLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statusLayout.Location = new System.Drawing.Point(18, 786);
            this.statusLayout.Margin = new System.Windows.Forms.Padding(0);
            this.statusLayout.Name = "statusLayout";
            this.statusLayout.RowCount = 1;
            this.statusLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.statusLayout.Size = new System.Drawing.Size(564, 24);
            this.statusLayout.TabIndex = 60;
            //
            // status
            //
            this.status.AutoEllipsis = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.status.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.status.Location = new System.Drawing.Point(4, 0);
            this.status.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(476, 24);
            this.status.TabIndex = 61;
            this.status.Text = "Ready";
            this.status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // activityBar
            //
            this.activityBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.activityBar.Location = new System.Drawing.Point(488, 7);
            this.activityBar.Margin = new System.Windows.Forms.Padding(4, 7, 4, 7);
            this.activityBar.MarqueeAnimationSpeed = 30;
            this.activityBar.Name = "activityBar";
            this.activityBar.Size = new System.Drawing.Size(72, 10);
            this.activityBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.activityBar.TabIndex = 62;
            //
            // ChatWindow
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(600, 820);
            this.Controls.Add(this.rootLayout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(440, 560);
            this.Name = "ChatWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VBAi — Your AI agent for VBA";

            this.optionsMenu.Location = new System.Drawing.Point(0, 0);
            this.optionsMenu.TabIndex = 0;
            this.optionsMenu.ResumeLayout(false);
            this.optionsMenu.PerformLayout();
            this.rootLayout.ResumeLayout(false);
            this.rootLayout.PerformLayout();
            this.headingLayout.ResumeLayout(false);
            this.headingLayout.PerformLayout();
            this.scopeLayout.ResumeLayout(false);
            this.scopeLayout.PerformLayout();
            this.workflowLayout.ResumeLayout(false);
            this.workflowLayout.PerformLayout();
            this.conversationPanel.ResumeLayout(false);
            this.conversationPanel.PerformLayout();
            this.historyPanel.ResumeLayout(false);
            this.historyPanel.PerformLayout();
            this.historyLayout.ResumeLayout(false);
            this.historyLayout.PerformLayout();
            this.historyActions.ResumeLayout(false);
            this.historyActions.PerformLayout();
            this.memoryPanel.ResumeLayout(false);
            this.memoryPanel.PerformLayout();
            this.memoryLayout.ResumeLayout(false);
            this.memoryLayout.PerformLayout();
            this.transcriptPanel.ResumeLayout(false);
            this.transcriptPanel.PerformLayout();
            this.composerLayout.ResumeLayout(false);
            this.composerLayout.PerformLayout();
            this.promptPanel.ResumeLayout(false);
            this.promptPanel.PerformLayout();
            this.promptHost.ResumeLayout(false);
            this.promptHost.PerformLayout();
            this.composerActions.ResumeLayout(false);
            this.composerActions.PerformLayout();
            this.contextPanel.ResumeLayout(false);
            this.contextPanel.PerformLayout();
            this.providerLayout.ResumeLayout(false);
            this.providerLayout.PerformLayout();
            this.footerLayout.ResumeLayout(false);
            this.footerLayout.PerformLayout();
            this.statusLayout.ResumeLayout(false);
            this.statusLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
