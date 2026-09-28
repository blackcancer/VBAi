namespace CodexVBE
{
    /// <summary>Panneau de recherche de dépôts et de gestion des pull requests GitHub.</summary>
    public sealed partial class GitHubPane
    {
        /// <summary>Conteneur des composants du panneau.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Onglets des pages de dépôt et de pull request.</summary>
        private System.Windows.Forms.TabControl pages;
        /// <summary>Page de recherche et sélection des dépôts.</summary>
        private System.Windows.Forms.TabPage repositoriesPage;
        /// <summary>Page de création et consultation des pull requests.</summary>
        private System.Windows.Forms.TabPage pullsPage;
        /// <summary>Disposition de la page des dépôts.</summary>
        private System.Windows.Forms.TableLayoutPanel repoLayout;
        /// <summary>Disposition des commandes de dépôt.</summary>
        private System.Windows.Forms.FlowLayoutPanel repoActions;
        /// <summary>Disposition des champs de création de dépôt.</summary>
        private System.Windows.Forms.FlowLayoutPanel createActions;
        /// <summary>Disposition principale des pull requests.</summary>
        private System.Windows.Forms.TableLayoutPanel pullLayout;
        /// <summary>Disposition des actions de pull request.</summary>
        private System.Windows.Forms.FlowLayoutPanel pullActions;
        /// <summary>Onglets de composition et de détail de pull request.</summary>
        private System.Windows.Forms.TabControl pullTabs;
        /// <summary>Onglet de création d’une pull request.</summary>
        private System.Windows.Forms.TabPage composeTab;
        /// <summary>Onglet des détails d’une pull request.</summary>
        private System.Windows.Forms.TabPage detailTab;
        /// <summary>Onglet des fichiers modifiés.</summary>
        private System.Windows.Forms.TabPage filesTab;
        /// <summary>Onglet des commentaires.</summary>
        private System.Windows.Forms.TabPage commentsTab;
        /// <summary>Onglet des vérifications.</summary>
        private System.Windows.Forms.TabPage checksTab;
        /// <summary>Disposition du formulaire de création.</summary>
        private System.Windows.Forms.TableLayoutPanel composeLayout;
        /// <summary>Disposition du commentaire et des détails.</summary>
        private System.Windows.Forms.TableLayoutPanel commentLayout;
        /// <summary>Champ de recherche des dépôts GitHub.</summary>
        private System.Windows.Forms.TextBox repositorySearch;
        /// <summary>Liste des dépôts accessibles.</summary>
        private System.Windows.Forms.ListBox repositoryList;
        /// <summary>Sélecteur de branche du dépôt sélectionné.</summary>
        private System.Windows.Forms.ComboBox repositoryBranch;
        /// <summary>Champ facultatif du compte d’organisation.</summary>
        private System.Windows.Forms.ComboBox organization;
        /// <summary>Champ du nom de dépôt à créer.</summary>
        private System.Windows.Forms.TextBox repositoryName;
        /// <summary>Option de création d’un dépôt privé.</summary>
        private System.Windows.Forms.CheckBox privateRepository;
        /// <summary>Commande de chargement des dépôts.</summary>
        private System.Windows.Forms.Button loadRepositories;
        /// <summary>Commande d’utilisation du dépôt choisi.</summary>
        private System.Windows.Forms.Button useRepository;
        /// <summary>Commande de création du dépôt renseigné.</summary>
        private System.Windows.Forms.Button createRepository;
        /// <summary>Liste des pull requests du dépôt.</summary>
        private System.Windows.Forms.ListBox pulls;
        /// <summary>Commande de chargement des pull requests.</summary>
        private System.Windows.Forms.Button loadPulls;
        /// <summary>Sélecteur de branche cible.</summary>
        private System.Windows.Forms.ComboBox targetBranch;
        /// <summary>Libellé de la branche source courante.</summary>
        private System.Windows.Forms.Label sourceLabel;
        /// <summary>Champ du titre de pull request.</summary>
        private System.Windows.Forms.TextBox pullTitle;
        /// <summary>Champ de description de pull request.</summary>
        private System.Windows.Forms.TextBox pullBody;
        /// <summary>Option de création en brouillon.</summary>
        private System.Windows.Forms.CheckBox draft;
        /// <summary>Commande de création de pull request.</summary>
        private System.Windows.Forms.Button createPull;
        /// <summary>Commande d’ouverture de la pull request sélectionnée.</summary>
        private System.Windows.Forms.Button openPull;
        /// <summary>Champ des détails de pull request.</summary>
        private System.Windows.Forms.TextBox pullDetails;
        /// <summary>Liste des fichiers de la pull request.</summary>
        private System.Windows.Forms.ListBox files;
        /// <summary>Liste des commentaires de revue.</summary>
        private System.Windows.Forms.ListBox comments;
        /// <summary>Champ du corps de commentaire.</summary>
        private System.Windows.Forms.TextBox commentBody;
        /// <summary>Liste des vérifications du commit.</summary>
        private System.Windows.Forms.TextBox checks;
        /// <summary>Libellé de statut du panneau.</summary>
        private System.Windows.Forms.Label status;
        /// <summary>Commande d’annulation de l’opération.</summary>
        private System.Windows.Forms.Button cancel;
        /// <summary>Disposition du pied et des informations d’état.</summary>
        private System.Windows.Forms.FlowLayoutPanel footer;
        /// <summary>Instructions de création et de revue.</summary>
        private System.Windows.Forms.ToolTip tips;
        /// <summary>Commande de chargement des brouillons.</summary>
        private System.Windows.Forms.Button loadDraft;
        /// <summary>Libellé de la branche cible.</summary>
        private System.Windows.Forms.Label targetLabel;
        /// <summary>Libellé du titre.</summary>
        private System.Windows.Forms.Label titleLabel;
        /// <summary>Libellé de la description.</summary>
        private System.Windows.Forms.Label bodyLabel;
        /// <summary>Libellé de recherche des dépôts.</summary>
        private System.Windows.Forms.Label searchLabel;
        /// <summary>Crée les contrôles et configure les pages du panneau GitHub.</summary>
        private void InitializeComponent() {
this.components = new System.ComponentModel.Container();
this.pages = new CodexVBE.ThemedTabControl();
this.repositoriesPage = new System.Windows.Forms.TabPage();
this.pullsPage = new System.Windows.Forms.TabPage();
this.repoLayout = new System.Windows.Forms.TableLayoutPanel();
this.repoActions = new System.Windows.Forms.FlowLayoutPanel();
this.createActions = new System.Windows.Forms.FlowLayoutPanel();
this.pullLayout = new System.Windows.Forms.TableLayoutPanel();
this.pullActions = new System.Windows.Forms.FlowLayoutPanel();
this.pullTabs = new CodexVBE.ThemedTabControl();
this.composeTab = new System.Windows.Forms.TabPage();
this.detailTab = new System.Windows.Forms.TabPage();
this.filesTab = new System.Windows.Forms.TabPage();
this.commentsTab = new System.Windows.Forms.TabPage();
this.checksTab = new System.Windows.Forms.TabPage();
this.composeLayout = new System.Windows.Forms.TableLayoutPanel();
this.commentLayout = new System.Windows.Forms.TableLayoutPanel();
this.repositorySearch = new System.Windows.Forms.TextBox();
this.repositoryList = new System.Windows.Forms.ListBox();
this.repositoryBranch = new CodexVBE.ThemedComboBox();
this.organization = new CodexVBE.ThemedComboBox();
this.repositoryName = new System.Windows.Forms.TextBox();
this.privateRepository = new System.Windows.Forms.CheckBox();
this.loadRepositories = new CodexVBE.ThemedButton();
this.useRepository = new CodexVBE.ThemedButton();
this.createRepository = new CodexVBE.ThemedButton();
this.pulls = new System.Windows.Forms.ListBox();
this.loadPulls = new CodexVBE.ThemedButton();
this.targetBranch = new CodexVBE.ThemedComboBox();
this.sourceLabel = new System.Windows.Forms.Label();
this.pullTitle = new System.Windows.Forms.TextBox();
this.pullBody = new System.Windows.Forms.TextBox();
this.draft = new System.Windows.Forms.CheckBox();
this.createPull = new CodexVBE.ThemedButton();
this.openPull = new CodexVBE.ThemedButton();
this.pullDetails = new System.Windows.Forms.TextBox();
this.files = new System.Windows.Forms.ListBox();
this.comments = new System.Windows.Forms.ListBox();
this.commentBody = new System.Windows.Forms.TextBox();
this.checks = new System.Windows.Forms.TextBox();
this.status = new System.Windows.Forms.Label();
this.cancel = new CodexVBE.ThemedButton();
this.footer = new System.Windows.Forms.FlowLayoutPanel();
this.tips = new System.Windows.Forms.ToolTip(this.components);
this.loadDraft = new CodexVBE.ThemedButton();
this.loadDraft.Name = "loadDraft";
this.loadDraft.Text = "Load prepared draft";
this.loadDraft.AutoSize = true;
this.loadDraft.Click += new System.EventHandler(this.LoadDraft_Click);
this.pullActions.Controls.Add(this.loadDraft);
this.SuspendLayout();
this.pages.Name = "pages";
this.repositoriesPage.Name = "repositoriesPage";
this.pullsPage.Name = "pullsPage";
this.repoLayout.Name = "repoLayout";
this.repoActions.Name = "repoActions";
this.createActions.Name = "createActions";
this.pullLayout.Name = "pullLayout";
this.pullActions.Name = "pullActions";
this.pullTabs.Name = "pullTabs";
this.composeTab.Name = "composeTab";
this.detailTab.Name = "detailTab";
this.filesTab.Name = "filesTab";
this.commentsTab.Name = "commentsTab";
this.checksTab.Name = "checksTab";
this.composeLayout.Name = "composeLayout";
this.commentLayout.Name = "commentLayout";
this.repositorySearch.Name = "repositorySearch";
this.repositoryList.Name = "repositoryList";
this.repositoryBranch.Name = "repositoryBranch";
this.organization.Name = "organization";
this.repositoryName.Name = "repositoryName";
this.privateRepository.Name = "privateRepository";
this.loadRepositories.Name = "loadRepositories";
this.useRepository.Name = "useRepository";
this.createRepository.Name = "createRepository";
this.pulls.Name = "pulls";
this.loadPulls.Name = "loadPulls";
this.targetBranch.Name = "targetBranch";
this.sourceLabel.Name = "sourceLabel";
this.pullTitle.Name = "pullTitle";
this.pullBody.Name = "pullBody";
this.draft.Name = "draft";
this.createPull.Name = "createPull";
this.openPull.Name = "openPull";
this.pullDetails.Name = "pullDetails";
this.files.Name = "files";
this.comments.Name = "comments";
this.commentBody.Name = "commentBody";
this.checks.Name = "checks";
this.status.Name = "status";
this.cancel.Name = "cancel";
this.footer.Name = "footer";
this.pages.Dock = System.Windows.Forms.DockStyle.Fill;
this.footer.Dock = System.Windows.Forms.DockStyle.Bottom;
this.footer.AutoSize = true;
this.status.AutoSize = true;
this.status.MaximumSize = new System.Drawing.Size(700, 0);
this.footer.Controls.Add(this.status);
this.cancel.Text = "Cancel";
this.cancel.AutoSize = true;
this.cancel.Click += new System.EventHandler(this.Cancel_Click);
this.footer.Controls.Add(this.cancel);
this.tips.SetToolTip(this.cancel, "Cancel");
this.cancel.Enabled = false;
this.repoLayout.ColumnCount = 1;
this.repoLayout.RowCount = 4;
this.repoLayout.Dock = System.Windows.Forms.DockStyle.Fill;
this.repoLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
this.repoLayout.Controls.Add(this.repositorySearch, 0, 0);
this.repositorySearch.Dock = System.Windows.Forms.DockStyle.Fill;
this.repoLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.repoLayout.Controls.Add(this.repoActions, 0, 1);
this.repoActions.Dock = System.Windows.Forms.DockStyle.Fill;
this.repoLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
this.repoLayout.Controls.Add(this.repositoryList, 0, 2);
this.repositoryList.Dock = System.Windows.Forms.DockStyle.Fill;
this.repoLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.repoLayout.Controls.Add(this.createActions, 0, 3);
this.createActions.Dock = System.Windows.Forms.DockStyle.Fill;
this.repositorySearch.AccessibleName = "Search repositories";
this.repositorySearch.TextChanged += new System.EventHandler(this.FilterRepositories);
this.loadRepositories.Text = "Load repositories";
this.loadRepositories.AutoSize = true;
this.loadRepositories.Click += new System.EventHandler(this.LoadRepositories_Click);
this.repoActions.Controls.Add(this.loadRepositories);
this.tips.SetToolTip(this.loadRepositories, "Load repositories");
this.repoActions.Controls.Add(this.repositoryBranch);
this.useRepository.Text = "Use repository";
this.useRepository.AutoSize = true;
this.useRepository.Click += new System.EventHandler(this.UseRepository_Click);
this.repoActions.Controls.Add(this.useRepository);
this.tips.SetToolTip(this.useRepository, "Use repository");
this.repoActions.AutoSize = true;
this.repositoryList.HorizontalScrollbar = true;
this.repositoryList.SelectedIndexChanged += new System.EventHandler(this.RepositoryChanged);
this.createActions.Controls.Add(this.repositoryName);
this.createActions.Controls.Add(this.organization);
this.createActions.Controls.Add(this.privateRepository);
this.repositoryName.AccessibleName = "New repository name";
this.organization.AccessibleName = "Repository owner";
this.organization.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
this.privateRepository.Text = "Private";
this.privateRepository.Checked = true;
this.privateRepository.AutoSize = true;
this.createRepository.Text = "Create repository";
this.createRepository.AutoSize = true;
this.createRepository.Click += new System.EventHandler(this.CreateRepository_Click);
this.createActions.Controls.Add(this.createRepository);
this.tips.SetToolTip(this.createRepository, "Create repository");
this.createActions.AutoSize = true;
this.repositoriesPage.Text = "Repositories";
this.repositoriesPage.Controls.Add(this.repoLayout);
this.pages.Controls.Add(this.repositoriesPage);
this.pullLayout.ColumnCount = 1;
this.pullLayout.RowCount = 3;
this.pullLayout.Dock = System.Windows.Forms.DockStyle.Fill;
this.pullLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.pullLayout.Controls.Add(this.pullActions, 0, 0);
this.pullActions.Dock = System.Windows.Forms.DockStyle.Fill;
this.pullLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 110F));
this.pullLayout.Controls.Add(this.pulls, 0, 1);
this.pulls.Dock = System.Windows.Forms.DockStyle.Fill;
this.pullLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
this.pullLayout.Controls.Add(this.pullTabs, 0, 2);
this.pullTabs.Dock = System.Windows.Forms.DockStyle.Fill;
this.loadPulls.Text = "Load pull requests";
this.loadPulls.AutoSize = true;
this.loadPulls.Click += new System.EventHandler(this.LoadPulls_Click);
this.pullActions.Controls.Add(this.loadPulls);
this.tips.SetToolTip(this.loadPulls, "Load pull requests");
this.openPull.Text = "Open on GitHub";
this.openPull.AutoSize = true;
this.openPull.Click += new System.EventHandler(this.OpenPull_Click);
this.pullActions.Controls.Add(this.openPull);
this.tips.SetToolTip(this.openPull, "Open on GitHub");
this.pullActions.AutoSize = true;
this.pulls.SelectedIndexChanged += new System.EventHandler(this.PullChanged);
this.pulls.HorizontalScrollbar = true;
this.composeLayout.ColumnCount = 1;
this.composeLayout.RowCount = 6;
this.composeLayout.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.composeLayout.Controls.Add(this.sourceLabel, 0, 0);
this.sourceLabel.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
this.composeLayout.Controls.Add(this.targetBranch, 0, 1);
this.targetBranch.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
this.composeLayout.Controls.Add(this.pullTitle, 0, 2);
this.pullTitle.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
this.composeLayout.Controls.Add(this.pullBody, 0, 3);
this.pullBody.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.composeLayout.Controls.Add(this.draft, 0, 4);
this.draft.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
this.composeLayout.Controls.Add(this.createPull, 0, 5);
this.createPull.Dock = System.Windows.Forms.DockStyle.Fill;
this.targetBranch.AccessibleName = "Target branch";
this.pullTitle.AccessibleName = "Pull request title";
this.pullBody.AccessibleName = "Pull request description";
this.draft.Text = "Draft";
this.draft.Checked = true;
this.draft.AutoSize = true;
this.createPull.Text = "Create pull request";
this.createPull.AutoSize = true;
this.createPull.Click += new System.EventHandler(this.CreatePull_Click);
this.pullBody.Multiline = true;
this.pullBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
this.pullBody.Dock = System.Windows.Forms.DockStyle.Fill;
this.pullDetails.Multiline = true;
this.pullDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both;
this.pullDetails.Dock = System.Windows.Forms.DockStyle.Fill;
this.pullDetails.ReadOnly = true;
this.checks.Multiline = true;
this.checks.ScrollBars = System.Windows.Forms.ScrollBars.Both;
this.checks.Dock = System.Windows.Forms.DockStyle.Fill;
this.checks.ReadOnly = true;
this.commentBody.Multiline = true;
this.commentBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
this.commentBody.Dock = System.Windows.Forms.DockStyle.Fill;
this.commentBody.ReadOnly = true;
this.commentLayout.ColumnCount = 1;
this.commentLayout.RowCount = 2;
this.commentLayout.Dock = System.Windows.Forms.DockStyle.Fill;
this.commentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45F));
this.commentLayout.Controls.Add(this.comments, 0, 0);
this.comments.Dock = System.Windows.Forms.DockStyle.Fill;
this.commentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 55F));
this.commentLayout.Controls.Add(this.commentBody, 0, 1);
this.commentBody.Dock = System.Windows.Forms.DockStyle.Fill;
this.comments.SelectedIndexChanged += new System.EventHandler(this.CommentChanged);
this.comments.DoubleClick += new System.EventHandler(this.OpenFile_Click);
this.files.DoubleClick += new System.EventHandler(this.OpenFile_Click);
this.files.HorizontalScrollbar = true;
this.comments.HorizontalScrollbar = true;
this.files.Dock = System.Windows.Forms.DockStyle.Fill;
this.composeTab.Text = "New pull request";
this.composeTab.Controls.Add(this.composeLayout);
this.pullTabs.Controls.Add(this.composeTab);
this.detailTab.Text = "Details";
this.detailTab.Controls.Add(this.pullDetails);
this.pullTabs.Controls.Add(this.detailTab);
this.filesTab.Text = "Files";
this.filesTab.Controls.Add(this.files);
this.pullTabs.Controls.Add(this.filesTab);
this.commentsTab.Text = "Comments";
this.commentsTab.Controls.Add(this.commentLayout);
this.pullTabs.Controls.Add(this.commentsTab);
this.checksTab.Text = "Checks";
this.checksTab.Controls.Add(this.checks);
this.pullTabs.Controls.Add(this.checksTab);
this.pullsPage.Text = "Pull requests";
this.pullsPage.Controls.Add(this.pullLayout);
this.pages.Controls.Add(this.pullsPage);
this.targetLabel = new System.Windows.Forms.Label();
this.titleLabel = new System.Windows.Forms.Label();
this.bodyLabel = new System.Windows.Forms.Label();
this.searchLabel = new System.Windows.Forms.Label();
this.targetLabel.Text = "Target branch";
this.titleLabel.Text = "Pull request title";
this.bodyLabel.Text = "Pull request description";
this.searchLabel.Text = "Search repositories";
this.targetLabel.AutoSize = true;
this.titleLabel.AutoSize = true;
this.bodyLabel.AutoSize = true;
this.searchLabel.AutoSize = true;
this.composeLayout.ColumnCount = 2;
this.composeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
this.composeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
this.composeLayout.SetColumn(this.targetBranch, 1);
this.composeLayout.SetColumn(this.pullTitle, 1);
this.composeLayout.SetColumn(this.pullBody, 1);
this.composeLayout.SetColumnSpan(this.sourceLabel, 2);
this.composeLayout.SetColumnSpan(this.createPull, 2);
this.composeLayout.Controls.Add(this.targetLabel, 0, 1);
this.composeLayout.Controls.Add(this.titleLabel, 0, 2);
this.composeLayout.Controls.Add(this.bodyLabel, 0, 3);
this.repoLayout.ColumnCount = 2;
this.repoLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
this.repoLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
this.repoLayout.SetColumn(this.repositorySearch, 1);
this.repoLayout.Controls.Add(this.searchLabel, 0, 0);
this.repoLayout.SetColumnSpan(this.repoActions, 2);
this.repoLayout.SetColumnSpan(this.repositoryList, 2);
this.repoLayout.SetColumnSpan(this.createActions, 2);
this.Controls.Add(this.pages);
this.Controls.Add(this.footer);
this.Name = "GitHubPane";
this.Size = new System.Drawing.Size(800, 550);
this.ResumeLayout(false);
this.PerformLayout();
} } }
