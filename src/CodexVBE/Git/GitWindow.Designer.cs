namespace CodexVBE
{
    /// <summary>Fenêtre de gestion du dépôt et de l’historique du document VBA.</summary>
    internal sealed partial class GitWindow
    {
        /// <summary>Conteneur des composants du formulaire.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Disposition racine de la fenêtre Git.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Libellé du document VBA lié au dépôt.</summary>
        private System.Windows.Forms.Label documentLabel;
        /// <summary>Libellé de l’adresse distante Git.</summary>
        private System.Windows.Forms.Label remoteLabel;
        /// <summary>Libellé de la branche courante.</summary>
        private System.Windows.Forms.Label branchLabel;
        /// <summary>Libellé du message de commit.</summary>
        private System.Windows.Forms.Label messageLabel;
        /// <summary>Instructions de configuration de la connexion Git.</summary>
        private System.Windows.Forms.Label help;
        /// <summary>Champ d’adresse du dépôt distant.</summary>
        private System.Windows.Forms.TextBox remote;
        /// <summary>Champ de nom de branche.</summary>
        private System.Windows.Forms.TextBox branch;
        /// <summary>Champ de message du commit.</summary>
        private System.Windows.Forms.TextBox commitMessage;
        /// <summary>Commande de connexion au dépôt.</summary>
        private System.Windows.Forms.Button connect;
        /// <summary>Commande de création d’un commit local.</summary>
        private System.Windows.Forms.Button commit;
        /// <summary>Commande de récupération des mises à jour distantes.</summary>
        private System.Windows.Forms.Button fetch;
        /// <summary>Libellé d’état de synchronisation.</summary>
        private System.Windows.Forms.Label syncStatus;
        /// <summary>Onglets des fonctions Git du document.</summary>
        private System.Windows.Forms.TabControl tabs;
        /// <summary>Onglet des changements locaux.</summary>
        private System.Windows.Forms.TabPage changesTab;
        /// <summary>Onglet de l’historique des commits.</summary>
        private System.Windows.Forms.TabPage historyTab;
        /// <summary>Disposition de la liste des changements et de leur diff.</summary>
        private System.Windows.Forms.SplitContainer changeSplit;
        /// <summary>Vue de comparaison de code.</summary>
        private CodexVBE.CodeDiffView diff;
        /// <summary>Colonne de code avant modification.</summary>
        private System.Windows.Forms.DataGridViewTextBoxColumn beforeColumn;
        /// <summary>Colonne de code après modification.</summary>
        private System.Windows.Forms.DataGridViewTextBoxColumn afterColumn;
        /// <summary>Liste des commits de l’historique.</summary>
        private System.Windows.Forms.ListBox history;
        /// <summary>Commande de comparaison de versions.</summary>
        private System.Windows.Forms.Button compare;
        /// <summary>Commande de publication vers la branche distante.</summary>
        private System.Windows.Forms.Button push;
        /// <summary>Commande de récupération et intégration distante.</summary>
        private System.Windows.Forms.Button pull;
        /// <summary>Commande de restauration d’un point d’historique.</summary>
        private System.Windows.Forms.Button restore;
        /// <summary>Disposition des commandes relatives aux changements.</summary>
        private System.Windows.Forms.FlowLayoutPanel actions;
        /// <summary>Liste des fichiers ou changements locaux.</summary>
        private System.Windows.Forms.CheckedListBox changes;
        /// <summary>Libellé de résultat des opérations Git.</summary>
        private System.Windows.Forms.Label status;
        /// <summary>Gestionnaire des infobulles des contrôles.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Onglet de gestion des branches.</summary>
        private System.Windows.Forms.TabPage branchesTab;
        /// <summary>Onglet des points de contrôle locaux.</summary>
        private System.Windows.Forms.TabPage checkpointsTab;
        /// <summary>Onglet de résolution des conflits.</summary>
        private System.Windows.Forms.TabPage conflictsTab;
        /// <summary>Disposition des commandes de branche.</summary>
        private System.Windows.Forms.FlowLayoutPanel branchActions;
        /// <summary>Disposition des commandes de points de contrôle.</summary>
        private System.Windows.Forms.FlowLayoutPanel checkpointActions;
        /// <summary>Disposition des commandes de résolution de conflit.</summary>
        private System.Windows.Forms.FlowLayoutPanel conflictActions;
        /// <summary>Liste des branches disponibles.</summary>
        private System.Windows.Forms.ListBox branchList;
        /// <summary>Liste des points de contrôle.</summary>
        private System.Windows.Forms.ListBox checkpointList;
        /// <summary>Liste des conflits détectés.</summary>
        private System.Windows.Forms.ListBox conflictList;
        /// <summary>Champ du nom de branche à créer ou suivre.</summary>
        private System.Windows.Forms.ComboBox branchName;
        /// <summary>Champ de référence distante de la branche.</summary>
        private System.Windows.Forms.Button branchRemote;
        /// <summary>Champ du nom du point de contrôle.</summary>
        private System.Windows.Forms.TextBox checkpointName;
        /// <summary>Champ du texte de résolution du conflit.</summary>
        private System.Windows.Forms.TextBox resolutionText;
        /// <summary>Commande de création d’une branche.</summary>
        private System.Windows.Forms.Button branchCreate;
        /// <summary>Commande de suivi d’une branche distante.</summary>
        private System.Windows.Forms.Button branchTrack;
        /// <summary>Commande de bascule vers une branche.</summary>
        private System.Windows.Forms.Button branchSwitch;
        /// <summary>Commande de démarrage d’une fusion.</summary>
        private System.Windows.Forms.Button mergeBegin;
        /// <summary>Commande de création d’un point de contrôle.</summary>
        private System.Windows.Forms.Button checkpointCreate;
        /// <summary>Commande de restauration du point sélectionné.</summary>
        private System.Windows.Forms.Button checkpointRestore;
        /// <summary>Commande de résolution en gardant la version locale.</summary>
        private System.Windows.Forms.Button mergeOurs;
        /// <summary>Commande de résolution en gardant la version entrante.</summary>
        private System.Windows.Forms.Button mergeTheirs;
        /// <summary>Commande de résolution par contenu textuel.</summary>
        private System.Windows.Forms.Button mergeText;
        /// <summary>Commande de finalisation de la fusion.</summary>
        private System.Windows.Forms.Button mergeComplete;
        /// <summary>Commande d’abandon de la fusion.</summary>
        private System.Windows.Forms.Button mergeAbort;
        /// <summary>Grille du contenu en conflit.</summary>
        private System.Windows.Forms.DataGridView conflictDiff;
        /// <summary>Colonne du contenu local.</summary>
        private System.Windows.Forms.DataGridViewTextBoxColumn conflictOurs;
        /// <summary>Colonne du contenu entrant.</summary>
        private System.Windows.Forms.DataGridViewTextBoxColumn conflictTheirs;

        /// <summary>Onglet des pull requests et fonctions GitHub.</summary>
        private System.Windows.Forms.TabPage githubTab;
        /// <summary>Panneau de gestion des pull requests GitHub.</summary>
        private CodexVBE.GitHubPane githubPane;
        /// <summary>Disposition des commandes de revue.</summary>
        private System.Windows.Forms.FlowLayoutPanel reviewActions;
        /// <summary>Commande d’ouverture du module concerné dans le VBE.</summary>
        private System.Windows.Forms.Button openModule;
        /// <summary>Commande de restauration du module concerné.</summary>
        private System.Windows.Forms.Button restoreModule;
        /// <summary>Commande d’annulation de l’opération active.</summary>
        private System.Windows.Forms.Button cancelOperation;
        /// <summary>Indicateur de progression de l’opération.</summary>
        private System.Windows.Forms.ProgressBar operationProgress;
        /// <summary>Disposition du statut et de la progression.</summary>
        private System.Windows.Forms.TableLayoutPanel operationStatus;
        /// <summary>Champ des détails du commit sélectionné.</summary>
        private System.Windows.Forms.TextBox historyDetails;
        /// <summary>Contenu de base utilisé pour les comparaisons ou conflits.</summary>
        private System.Windows.Forms.TextBox baseContent;
        /// <summary>Onglet d’import de changements.</summary>
        private System.Windows.Forms.TabPage importTab;
        /// <summary>Résumé du contenu à importer.</summary>
        private System.Windows.Forms.TextBox importSummary;
        /// <summary>Commande de comparaison avec une entrée d’historique.</summary>
        private System.Windows.Forms.Button historyCompare;
        /// <summary>Commande d’aperçu avant import.</summary>
        private System.Windows.Forms.Button previewImport;
        /// <summary>Disposition des libellés de base et de résultat du conflit.</summary>
        private System.Windows.Forms.TableLayoutPanel conflictLayout;
        /// <summary>Libellé du contenu ancêtre du conflit.</summary>
        private System.Windows.Forms.Label ancestorLabel;
        /// <summary>Libellé du résultat de résolution.</summary>
        private System.Windows.Forms.Label resultLabel;
        /// <summary>Crée les contrôles et configure la disposition de la fenêtre Git.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.documentLabel = new System.Windows.Forms.Label();
            this.remoteLabel = new System.Windows.Forms.Label();
            this.branchLabel = new System.Windows.Forms.Label();
            this.messageLabel = new System.Windows.Forms.Label();
            this.help = new System.Windows.Forms.Label();
            this.remote = new System.Windows.Forms.TextBox();
            this.branch = new System.Windows.Forms.TextBox();
            this.commitMessage = new System.Windows.Forms.TextBox();
            this.connect = new CodexVBE.ThemedButton();
            this.commit = new CodexVBE.ThemedButton();
            this.fetch = new CodexVBE.ThemedButton();
            this.syncStatus = new System.Windows.Forms.Label();
            this.tabs = new CodexVBE.ThemedTabControl();
            this.changesTab = new System.Windows.Forms.TabPage();
            this.historyTab = new System.Windows.Forms.TabPage();
            this.changeSplit = new System.Windows.Forms.SplitContainer();
            this.diff = new CodexVBE.CodeDiffView();
            this.beforeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.afterColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.history = new System.Windows.Forms.ListBox();
            this.compare = new CodexVBE.ThemedButton();
            this.push = new CodexVBE.ThemedButton();
            this.pull = new CodexVBE.ThemedButton();
            this.restore = new CodexVBE.ThemedButton();
            this.actions = new System.Windows.Forms.FlowLayoutPanel();
            this.changes = new System.Windows.Forms.CheckedListBox();
            this.status = new System.Windows.Forms.Label();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.branchesTab = new System.Windows.Forms.TabPage();
            this.checkpointsTab = new System.Windows.Forms.TabPage();
            this.conflictsTab = new System.Windows.Forms.TabPage();
            this.branchActions = new System.Windows.Forms.FlowLayoutPanel();
            this.checkpointActions = new System.Windows.Forms.FlowLayoutPanel();
            this.conflictActions = new System.Windows.Forms.FlowLayoutPanel();
            this.branchList = new System.Windows.Forms.ListBox();
            this.checkpointList = new System.Windows.Forms.ListBox();
            this.conflictList = new System.Windows.Forms.ListBox();
            this.branchName = new CodexVBE.ThemedComboBox();
            this.branchRemote = new CodexVBE.ThemedButton();
            this.checkpointName = new System.Windows.Forms.TextBox();
            this.resolutionText = new System.Windows.Forms.TextBox();
            this.branchCreate = new CodexVBE.ThemedButton();
            this.branchTrack = new CodexVBE.ThemedButton();
            this.branchSwitch = new CodexVBE.ThemedButton();
            this.mergeBegin = new CodexVBE.ThemedButton();
            this.checkpointCreate = new CodexVBE.ThemedButton();
            this.checkpointRestore = new CodexVBE.ThemedButton();
            this.mergeOurs = new CodexVBE.ThemedButton();
            this.mergeTheirs = new CodexVBE.ThemedButton();
            this.mergeText = new CodexVBE.ThemedButton();
            this.mergeComplete = new CodexVBE.ThemedButton();
            this.mergeAbort = new CodexVBE.ThemedButton();
            this.conflictDiff = new System.Windows.Forms.DataGridView();
            this.conflictOurs = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.conflictTheirs = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.branchActions.Name = "branchActions";
            this.checkpointActions.Name = "checkpointActions";
            this.conflictActions.Name = "conflictActions";
            this.branchCreate.Name = "branchCreate";
            this.branchRemote.Name = "branchRemote";
            this.branchTrack.Name = "branchTrack";
            this.branchSwitch.Name = "branchSwitch";
            this.mergeBegin.Name = "mergeBegin";
            this.checkpointCreate.Name = "checkpointCreate";
            this.checkpointRestore.Name = "checkpointRestore";
            this.mergeOurs.Name = "mergeOurs";
            this.mergeTheirs.Name = "mergeTheirs";
            this.mergeText.Name = "mergeText";
            this.mergeComplete.Name = "mergeComplete";
            this.mergeAbort.Name = "mergeAbort";
            this.conflictOurs.Name = "conflictOurs";
            this.conflictTheirs.Name = "conflictTheirs";
            this.githubTab = new System.Windows.Forms.TabPage();
            this.githubPane = new CodexVBE.GitHubPane();
            this.reviewActions = new System.Windows.Forms.FlowLayoutPanel();
            this.openModule = new CodexVBE.ThemedButton();
            this.restoreModule = new CodexVBE.ThemedButton();
            this.cancelOperation = new CodexVBE.ThemedButton();
            this.operationProgress = new System.Windows.Forms.ProgressBar();
            this.operationStatus = new System.Windows.Forms.TableLayoutPanel();
            this.historyDetails = new System.Windows.Forms.TextBox();
            this.baseContent = new System.Windows.Forms.TextBox();
            this.importTab = new System.Windows.Forms.TabPage();
            this.importSummary = new System.Windows.Forms.TextBox();
            this.layout.SuspendLayout();
            this.actions.SuspendLayout();
            this.SuspendLayout();
            // layout
            this.layout.ColumnCount = 2;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 135F));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 10;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Padding = new System.Windows.Forms.Padding(20);
            this.layout.Controls.Add(this.documentLabel, 0, 0);
            this.layout.SetColumnSpan(this.documentLabel, 2);
            this.layout.Controls.Add(this.help, 0, 1);
            this.layout.SetColumnSpan(this.help, 2);
            this.layout.Controls.Add(this.remoteLabel, 0, 2);
            this.layout.Controls.Add(this.remote, 1, 2);
            this.layout.Controls.Add(this.branchLabel, 0, 3);
            this.layout.Controls.Add(this.branch, 1, 3);
            this.layout.Controls.Add(this.connect, 1, 4);
            this.layout.Controls.Add(this.messageLabel, 0, 5);
            this.layout.Controls.Add(this.commitMessage, 1, 5);
            this.layout.Controls.Add(this.syncStatus, 0, 6);
            this.layout.SetColumnSpan(this.syncStatus, 2);
            this.layout.Controls.Add(this.actions, 0, 7);
            this.layout.SetColumnSpan(this.actions, 2);
            this.layout.Controls.Add(this.tabs, 0, 8);
            this.layout.SetColumnSpan(this.tabs, 2);
            this.layout.Controls.Add(this.operationStatus, 0, 9);
            this.layout.SetColumnSpan(this.operationStatus, 2);
            this.layout.Name = "layout";
            // documentLabel
            this.documentLabel.Text = "GitHub · VBA sources";
            this.documentLabel.Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold);
            this.documentLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentLabel.AutoEllipsis = true;
            this.documentLabel.Name = "documentLabel";
            // help
            this.help.Text = "VBA changes become source files in the repository. No folder next to the macro.\r\nThe first pull replaces VBA with a restorable backup.\r\nSign in through Git Credential Manager; commit identity is configured in Git.";
            this.help.Dock = System.Windows.Forms.DockStyle.Fill;
            this.help.AutoSize = true;
            this.help.Name = "help";
            // fields
            this.remoteLabel.Text = "GitHub repository";
            this.remoteLabel.AutoSize = true;
            this.remoteLabel.Name = "remoteLabel";
            this.remote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.remote.Name = "remote";
            this.remote.TabIndex = 0;
            this.toolTips.SetToolTip(this.remote, "HTTPS URL, for example https://github.com/organization/macros.git. Do not paste a token.");
            this.branchLabel.Text = "Branch";
            this.branchLabel.AutoSize = true;
            this.branchLabel.Name = "branchLabel";
            this.branch.Text = "main";
            this.branch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.branch.Name = "branch";
            this.branch.TabIndex = 1;
            this.toolTips.SetToolTip(this.branch, "Branch to synchronize. Pushes are never forced.");
            this.connect.Text = "Link repository";
            this.connect.AutoSize = true;
            this.connect.Name = "connect";
            this.connect.TabIndex = 2;
            this.connect.Click += new System.EventHandler(this.Connect_Click);
            this.toolTips.SetToolTip(this.connect, "Save the link for this document on this computer. Does not publish code.");
            this.messageLabel.Text = "Commit message";
            this.messageLabel.AutoSize = true;
            this.messageLabel.Name = "messageLabel";
            this.commitMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commitMessage.Multiline = true;
            this.commitMessage.Name = "commitMessage";
            this.commitMessage.TabIndex = 3;
            // actions
            this.actions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.actions.AutoSize = true;
            this.actions.Name = "actions";
            this.actions.Controls.Add(this.compare);
            this.actions.Controls.Add(this.commit);
            this.actions.Controls.Add(this.fetch);
            this.actions.Controls.Add(this.push);
            this.actions.Controls.Add(this.pull);
            this.actions.Controls.Add(this.restore);
            this.compare.Text = "Compare";
            this.compare.AutoSize = true;
            this.compare.Enabled = false;
            this.compare.Name = "compare";
            this.compare.Click += new System.EventHandler(this.Compare_Click);
            this.toolTips.SetToolTip(this.compare, "Compare live VBA with the last synchronized state without network access.");
            this.commit.Text = "Commit selected";
            this.commit.AutoSize = true;
            this.commit.Enabled = false;
            this.commit.Name = "commit";
            this.commit.Click += new System.EventHandler(this.Commit_Click);
            this.toolTips.SetToolTip(this.commit, "Commit checked modules and their form resources. Push publishes them afterwards.");
            this.fetch.Text = "Fetch";
            this.fetch.AutoSize = true;
            this.fetch.Enabled = false;
            this.fetch.Name = "fetch";
            this.fetch.Click += new System.EventHandler(this.Fetch_Click);
            this.toolTips.SetToolTip(this.fetch, "Refresh remote commits and counts without changing VBA.");
            this.push.Text = "Push";
            this.push.AutoSize = true;
            this.push.Enabled = false;
            this.push.Name = "push";
            this.push.Click += new System.EventHandler(this.Push_Click);
            this.toolTips.SetToolTip(this.push, "Publish local commits on the chosen branch. Does not publish uncommitted changes.");
            this.pull.Text = "Pull and import";
            this.pull.AutoSize = true;
            this.pull.Enabled = false;
            this.pull.Name = "pull";
            this.pull.Click += new System.EventHandler(this.Pull_Click);
            this.toolTips.SetToolTip(this.pull, "Download and import sources. Refuses unsynchronized local changes and divergent histories.");
            this.restore.Text = "Restore VBA";
            this.restore.AutoSize = true;
            this.restore.Enabled = false;
            this.restore.Name = "restore";
            this.restore.Click += new System.EventHandler(this.Restore_Click);
            this.toolTips.SetToolTip(this.restore, "Restore the backup made before the last import if VBA has not changed since. Does not rewrite GitHub.");
            // changes and status
            this.changes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changes.HorizontalScrollbar = true;
            this.changes.IntegralHeight = false;
            this.changes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.changes.Name = "changes";
            this.changes.SelectedIndexChanged += new System.EventHandler(this.Changes_SelectedIndexChanged);
            this.syncStatus.Text = "Branch not linked";
            this.syncStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.syncStatus.Name = "syncStatus";
            this.tabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabs.Name = "tabs";
            this.tabs.Controls.Add(this.changesTab);
            this.tabs.Controls.Add(this.historyTab);
            this.tabs.Controls.Add(this.branchesTab);
            this.tabs.Controls.Add(this.checkpointsTab);
            this.tabs.Controls.Add(this.conflictsTab);
            this.branchesTab.Text = "Branches";
            this.branchesTab.Name = "branchesTab";
            this.branchesTab.Enabled = false;
            this.branchesTab.Controls.Add(this.branchList);
            this.branchesTab.Controls.Add(this.branchActions);
            this.branchList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.branchList.Name = "branchList";
            this.branchActions.Dock = System.Windows.Forms.DockStyle.Top;
            this.branchActions.AutoSize = true;
            this.branchActions.Controls.Add(this.branchName);
            this.branchActions.Controls.Add(this.branchCreate);
            this.branchActions.Controls.Add(this.branchRemote);
            this.branchActions.Controls.Add(this.branchTrack);
            this.branchActions.Controls.Add(this.branchSwitch);
            this.branchActions.Controls.Add(this.mergeBegin);
            this.branchName.Width = 170;
            this.branchName.Name = "branchName";
            this.branchName.AccessibleName = "Branch name to create or track";
            this.branchCreate.Text = "Create";
            this.branchCreate.AutoSize = true;
            this.branchCreate.Click += new System.EventHandler(this.BranchCreate_Click);
            this.branchRemote.Text = "Remote branches";
            this.branchRemote.AutoSize = true;
            this.branchRemote.Click += new System.EventHandler(this.RemoteBranches_Click);
            this.toolTips.SetToolTip(this.branchRemote, "List GitHub repository branches in the name selector without importing VBA.");
            this.branchTrack.Text = "Track remote";
            this.branchTrack.AutoSize = true;
            this.branchTrack.Click += new System.EventHandler(this.BranchTrack_Click);
            this.branchSwitch.Text = "Switch";
            this.branchSwitch.AutoSize = true;
            this.branchSwitch.Click += new System.EventHandler(this.BranchSwitch_Click);
            this.mergeBegin.Text = "Merge";
            this.mergeBegin.AutoSize = true;
            this.mergeBegin.Click += new System.EventHandler(this.MergeBegin_Click);
            this.toolTips.SetToolTip(this.branchCreate, "Create the named branch from the current commit without publishing it.");
            this.toolTips.SetToolTip(this.branchTrack, "Fetch the named remote branch without importing its VBA.");
            this.toolTips.SetToolTip(this.branchSwitch, "Import the selected branch with a checkpoint first. Uncommitted changes are refused.");
            this.toolTips.SetToolTip(this.mergeBegin, "Prepare a merge from the selected branch into the current branch.");
            this.checkpointsTab.Text = "Checkpoints";
            this.checkpointsTab.Name = "checkpointsTab";
            this.checkpointsTab.Enabled = false;
            this.checkpointsTab.Controls.Add(this.checkpointList);
            this.checkpointsTab.Controls.Add(this.checkpointActions);
            this.checkpointList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkpointList.Name = "checkpointList";
            this.checkpointList.HorizontalScrollbar = true;
            this.checkpointActions.Dock = System.Windows.Forms.DockStyle.Top;
            this.checkpointActions.AutoSize = true;
            this.checkpointActions.Controls.Add(this.checkpointName);
            this.checkpointActions.Controls.Add(this.checkpointCreate);
            this.checkpointActions.Controls.Add(this.checkpointRestore);
            this.checkpointName.Width = 240;
            this.checkpointName.Name = "checkpointName";
            this.checkpointName.AccessibleName = "Checkpoint name";
            this.checkpointCreate.Text = "Create checkpoint";
            this.checkpointCreate.AutoSize = true;
            this.checkpointCreate.Click += new System.EventHandler(this.CheckpointCreate_Click);
            this.checkpointRestore.Text = "Restore selection";
            this.checkpointRestore.AutoSize = true;
            this.checkpointRestore.Click += new System.EventHandler(this.CheckpointRestore_Click);
            this.toolTips.SetToolTip(this.checkpointCreate, "Save all current VBA in a private checkpoint without committing to the published branch.");
            this.toolTips.SetToolTip(this.checkpointRestore, "Restore the selected checkpoint after backing up the current state. Does not rewrite GitHub.");
            this.conflictsTab.Text = "Conflicts";
            this.conflictsTab.Name = "conflictsTab";
            this.conflictsTab.Enabled = false;
            this.conflictsTab.Controls.Add(this.resolutionText);
            this.conflictsTab.Controls.Add(this.conflictDiff);
            this.conflictsTab.Controls.Add(this.conflictList);
            this.conflictsTab.Controls.Add(this.conflictActions);
            this.conflictList.Dock = System.Windows.Forms.DockStyle.Top;
            this.conflictList.Height = 55;
            this.conflictList.Name = "conflictList";
            this.conflictList.SelectedIndexChanged += new System.EventHandler(this.ConflictList_SelectedIndexChanged);
            this.conflictDiff.Dock = System.Windows.Forms.DockStyle.Top;
            this.conflictDiff.Height = 110;
            this.conflictDiff.Name = "conflictDiff";
            this.conflictDiff.ReadOnly = true;
            this.conflictDiff.AllowUserToAddRows = false;
            this.conflictDiff.AllowUserToDeleteRows = false;
            this.conflictDiff.RowHeadersVisible = false;
            this.conflictDiff.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.conflictDiff.BackgroundColor = System.Drawing.Color.White;
            this.conflictDiff.Font = new System.Drawing.Font("Consolas", 9F);
            this.conflictDiff.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { this.conflictOurs, this.conflictTheirs });
            this.conflictOurs.HeaderText = "Current branch";
            this.conflictOurs.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.conflictTheirs.HeaderText = "Incoming branch";
            this.conflictTheirs.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.resolutionText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resolutionText.Multiline = true;
            this.resolutionText.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.resolutionText.WordWrap = false;
            this.resolutionText.Font = new System.Drawing.Font("Consolas", 9F);
            this.resolutionText.Name = "resolutionText";
            this.resolutionText.AccessibleName = "Full resolution content for the selected VBA file";
            this.conflictActions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.conflictActions.AutoSize = true;
            this.conflictActions.Controls.Add(this.mergeOurs);
            this.conflictActions.Controls.Add(this.mergeTheirs);
            this.conflictActions.Controls.Add(this.mergeText);
            this.conflictActions.Controls.Add(this.mergeComplete);
            this.conflictActions.Controls.Add(this.mergeAbort);
            this.mergeOurs.Text = "Keep local";
            this.mergeOurs.AutoSize = true;
            this.mergeOurs.Click += new System.EventHandler(this.MergeOurs_Click);
            this.mergeTheirs.Text = "Keep incoming";
            this.mergeTheirs.AutoSize = true;
            this.mergeTheirs.Click += new System.EventHandler(this.MergeTheirs_Click);
            this.mergeText.Text = "Use this text";
            this.mergeText.AutoSize = true;
            this.mergeText.Click += new System.EventHandler(this.MergeText_Click);
            this.mergeComplete.Text = "Complete merge";
            this.mergeComplete.AutoSize = true;
            this.mergeComplete.Click += new System.EventHandler(this.MergeComplete_Click);
            this.mergeAbort.Text = "Abort merge";
            this.mergeAbort.AutoSize = true;
            this.mergeAbort.Click += new System.EventHandler(this.MergeAbort_Click);
            this.toolTips.SetToolTip(this.mergeOurs, "Choose the entire current-branch file for the selected conflict.");
            this.toolTips.SetToolTip(this.mergeTheirs, "Choose the entire incoming-branch file for the selected conflict.");
            this.toolTips.SetToolTip(this.mergeText, "Replace the entire VBA text file with the text edited above.");
            this.toolTips.SetToolTip(this.mergeComplete, "Validate all files, create a merge commit, then import VBA with a backup.");
            this.toolTips.SetToolTip(this.mergeAbort, "Discard only the prepared merge. VBA and commits remain intact.");
            this.changesTab.Text = "Git changes";
            this.changesTab.Name = "changesTab";
            this.changesTab.Controls.Add(this.changeSplit);
            this.historyTab.Text = "History";
            this.historyTab.Name = "historyTab";
            this.historyTab.Controls.Add(this.history);
            this.history.Dock = System.Windows.Forms.DockStyle.Fill;
            this.history.HorizontalScrollbar = true;
            this.history.Name = "history";
            this.changeSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeSplit.Size = new System.Drawing.Size(860, 260);
            this.changeSplit.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.changeSplit.SplitterDistance = 200;
            this.changeSplit.Name = "changeSplit";
            this.changeSplit.Panel1.Controls.Add(this.changes);
            this.changeSplit.Panel2.Controls.Add(this.diff);
            this.diff.Dock = System.Windows.Forms.DockStyle.Fill;
            this.diff.Name = "diff";
            this.status.Text = "Link an existing GitHub repository to get started.";
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Padding = new System.Windows.Forms.Padding(0, 6, 8, 0);
            this.status.Name = "status";
                        this.githubTab.Name = "githubTab";
                        this.githubPane.Name = "githubPane";
                        this.reviewActions.Name = "reviewActions";
                        this.openModule.Name = "openModule";
                        this.restoreModule.Name = "restoreModule";
                        this.cancelOperation.Name = "cancelOperation";
                        this.operationProgress.Name = "operationProgress";
                        this.historyDetails.Name = "historyDetails";
                        this.baseContent.Name = "baseContent";
                        this.importTab.Name = "importTab";
                        this.importSummary.Name = "importSummary";
            this.githubTab.Text = "GitHub";
            this.githubPane.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubTab.Controls.Add(this.githubPane);
            this.tabs.Controls.Add(this.githubTab);
            this.importTab.Text = "Import summary";
            this.importSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.importSummary.Multiline = true;
            this.importSummary.ReadOnly = true;
            this.importSummary.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.importTab.Controls.Add(this.importSummary);
            this.tabs.Controls.Add(this.importTab);
            this.reviewActions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.reviewActions.AutoSize = true;
            this.changesTab.Controls.Add(this.reviewActions);
            this.openModule.Text = "Open in the VBE";
            this.openModule.AutoSize = true;
            this.openModule.Click += new System.EventHandler(this.OpenModule_Click);
            this.restoreModule.Text = "Restore selected module";
            this.restoreModule.AutoSize = true;
            this.restoreModule.Click += new System.EventHandler(this.RestoreModule_Click);
            this.reviewActions.Controls.Add(this.openModule);
            this.reviewActions.Controls.Add(this.restoreModule);
            this.cancelOperation.Text = "Cancel";
            this.cancelOperation.Enabled = false;
            this.cancelOperation.AutoSize = true;
            this.cancelOperation.Click += new System.EventHandler(this.CancelOperation_Click);
            this.operationProgress.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.operationProgress.Visible = false;
            this.operationStatus.Name = "operationStatus";
            this.operationStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.operationStatus.Margin = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.operationStatus.ColumnCount = 2;
            this.operationStatus.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.operationStatus.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.operationStatus.RowCount = 2;
            this.operationStatus.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.operationStatus.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 3F));
            this.operationStatus.Controls.Add(this.status, 0, 0);
            this.operationStatus.Controls.Add(this.cancelOperation, 1, 0);
            this.operationStatus.Controls.Add(this.operationProgress, 0, 1);
            this.operationStatus.SetColumnSpan(this.operationProgress, 2);
            this.operationProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.operationProgress.Margin = new System.Windows.Forms.Padding(0);
            this.cancelOperation.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            this.cancelOperation.Visible = false;
            this.history.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.history.SelectedIndexChanged += new System.EventHandler(this.HistoryChanged);
            this.historyDetails.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.historyDetails.Height = 100;
            this.historyDetails.Multiline = true;
            this.historyDetails.ReadOnly = true;
            this.historyDetails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.historyTab.Controls.Add(this.historyDetails);
            this.checkpointList.SelectedIndexChanged += new System.EventHandler(this.CheckpointChanged);
            this.baseContent.Dock = System.Windows.Forms.DockStyle.Top;
            this.baseContent.Height = 100;
            this.baseContent.Multiline = true;
            this.baseContent.ReadOnly = true;
            this.baseContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.baseContent.Font = new System.Drawing.Font("Consolas", 9F);
            this.baseContent.AccessibleName = "Common ancestor";
            this.conflictsTab.Controls.Add(this.baseContent);
            this.toolTips.SetToolTip(this.baseContent, "Common ancestor");
            this.toolTips.SetToolTip(this.history, "Select one commit for details, or two commits to compare their VBA sources.");
            this.toolTips.SetToolTip(this.restoreModule, "Restore only the selected module from the reviewed revision, with a checkpoint first.");
            this.changes.CheckOnClick = true;
            this.historyCompare = new CodexVBE.ThemedButton();
            this.historyCompare.Name = "historyCompare";
            this.historyCompare.Text = "Compare revisions";
            this.historyCompare.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.historyCompare.AutoSize = true;
            this.historyCompare.Click += new System.EventHandler(this.HistoryChanged);
            this.historyTab.Controls.Add(this.historyCompare);
            this.previewImport = new CodexVBE.ThemedButton();
            this.previewImport.Name = "previewImport";
            this.previewImport.Text = "Preview import";
            this.previewImport.AutoSize = true;
            this.previewImport.Click += new System.EventHandler(this.PreviewImport_Click);
            this.actions.Controls.Add(this.previewImport);
            this.actions.Controls.SetChildIndex(this.previewImport, 4);
            this.commit.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.push.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.pull.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.conflictLayout = new System.Windows.Forms.TableLayoutPanel();
            this.ancestorLabel = new System.Windows.Forms.Label();
            this.resultLabel = new System.Windows.Forms.Label();
            this.conflictLayout.Name = "conflictLayout";
            this.conflictLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictLayout.ColumnCount = 1;
            this.conflictLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.conflictLayout.RowCount = 7;
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.ancestorLabel.Name = "ancestorLabel";
            this.ancestorLabel.Text = "Common ancestor";
            this.ancestorLabel.AutoSize = true;
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Text = "Full resolution content for the selected VBA file";
            this.resultLabel.AutoSize = true;
            this.conflictList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.baseContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictDiff.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictLayout.Controls.Add(this.conflictList, 0, 0);
            this.conflictLayout.Controls.Add(this.ancestorLabel, 0, 1);
            this.conflictLayout.Controls.Add(this.baseContent, 0, 2);
            this.conflictLayout.Controls.Add(this.conflictDiff, 0, 3);
            this.conflictLayout.Controls.Add(this.resultLabel, 0, 4);
            this.conflictLayout.Controls.Add(this.resolutionText, 0, 5);
            this.conflictLayout.Controls.Add(this.conflictActions, 0, 6);
            this.conflictsTab.Controls.Add(this.conflictLayout);
            // window
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.BackColor = System.Drawing.Color.FromArgb(248, 250, 252);
            this.ForeColor = System.Drawing.Color.FromArgb(30, 41, 59);
            this.ClientSize = new System.Drawing.Size(960, 780);
            this.MinimumSize = new System.Drawing.Size(700, 580);
            this.Controls.Add(this.layout);
            this.Name = "GitWindow";
            this.Text = "GitHub · VBAi";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.actions.ResumeLayout(false);
            this.actions.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
