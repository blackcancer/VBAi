namespace VBAi
{
    /// <summary>Designer-generated repository search, selection, and creation controls.</summary>
    public sealed partial class GitHubRepositoriesView
    {
        /// <summary>Layout for repository search and selection.</summary>
        internal System.Windows.Forms.TableLayoutPanel repoLayout;
        /// <summary>Actions for loading and selecting a repository.</summary>
        internal System.Windows.Forms.FlowLayoutPanel repoActions;
        /// <summary>Actions for creating a repository.</summary>
        internal System.Windows.Forms.FlowLayoutPanel createActions;
        /// <summary>Repository filter input.</summary>
        internal VBAi.UiTextBox repositorySearch;
        /// <summary>Available repositories matching the filter.</summary>
        internal VBAi.UiListBox repositoryList;
        /// <summary>Branches available for the selected repository.</summary>
        internal VBAi.ThemedComboBox repositoryBranch;
        /// <summary>Organization that will own a new repository.</summary>
        internal VBAi.ThemedComboBox organization;
        /// <summary>Name for a new repository.</summary>
        internal VBAi.UiTextBox repositoryName;
        /// <summary>Whether the new repository is private.</summary>
        internal System.Windows.Forms.CheckBox privateRepository;
        /// <summary>Loads repositories from the connected account.</summary>
        internal VBAi.ThemedButton loadRepositories;
        /// <summary>Uses the selected repository for the current project.</summary>
        internal VBAi.ThemedButton useRepository;
        /// <summary>Creates the configured repository.</summary>
        internal VBAi.ThemedButton createRepository;
        /// <summary>Caption for the repository filter.</summary>
        internal System.Windows.Forms.Label searchLabel;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with repository actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges repository search, selection, and creation controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.repoLayout = new System.Windows.Forms.TableLayoutPanel();
            this.repoActions = new System.Windows.Forms.FlowLayoutPanel();
            this.createActions = new System.Windows.Forms.FlowLayoutPanel();
            this.repositorySearch = new VBAi.UiTextBox();
            this.repositoryList = new VBAi.UiListBox();
            this.repositoryBranch = new VBAi.ThemedComboBox();
            this.organization = new VBAi.ThemedComboBox();
            this.repositoryName = new VBAi.UiTextBox();
            this.privateRepository = new System.Windows.Forms.CheckBox();
            this.loadRepositories = new VBAi.ThemedButton();
            this.useRepository = new VBAi.ThemedButton();
            this.createRepository = new VBAi.ThemedButton();
            this.searchLabel = new System.Windows.Forms.Label();
            this.repoLayout.SuspendLayout();
            this.repoActions.SuspendLayout();
            this.createActions.SuspendLayout();
            this.SuspendLayout();
            this.repoLayout.Name = "repoLayout";
            this.repoActions.Name = "repoActions";
            this.createActions.Name = "createActions";
            this.repositorySearch.Name = "repositorySearch";
            this.repositoryList.Name = "repositoryList";
            this.repositoryBranch.Name = "repositoryBranch";
            this.organization.Name = "organization";
            this.repositoryName.Name = "repositoryName";
            this.privateRepository.Name = "privateRepository";
            this.loadRepositories.Name = "loadRepositories";
            this.useRepository.Name = "useRepository";
            this.createRepository.Name = "createRepository";
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
            this.loadRepositories.Text = "Load repositories";
            this.loadRepositories.AutoSize = true;
            this.repoActions.Controls.Add(this.loadRepositories);
            this.toolTips.SetToolTip(this.loadRepositories, "Load repositories");
            this.repoActions.Controls.Add(this.repositoryBranch);
            this.useRepository.Text = "Use repository";
            this.useRepository.AutoSize = true;
            this.repoActions.Controls.Add(this.useRepository);
            this.toolTips.SetToolTip(this.useRepository, "Use repository");
            this.repoActions.AutoSize = true;
            this.repositoryList.HorizontalScrollbar = true;
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
            this.createActions.Controls.Add(this.createRepository);
            this.toolTips.SetToolTip(this.createRepository, "Create repository");
            this.createActions.AutoSize = true;
            this.searchLabel.Text = "Search repositories";
            this.searchLabel.AutoSize = true;
            this.repoLayout.ColumnCount = 2;
            this.repoLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.repoLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.repoLayout.SetColumn(this.repositorySearch, 1);
            this.repoLayout.Controls.Add(this.searchLabel, 0, 0);
            this.repoLayout.SetColumnSpan(this.repoActions, 2);
            this.repoLayout.SetColumnSpan(this.repositoryList, 2);
            this.repoLayout.SetColumnSpan(this.createActions, 2);
            this.Controls.Add(this.repoLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubRepositoriesView";
            this.Size = new System.Drawing.Size(860, 500);
            this.repoLayout.Location = new System.Drawing.Point(0, 0);
            this.repoLayout.Size = new System.Drawing.Size(860, 500);
            this.repoLayout.TabIndex = 0;
            this.repoActions.Location = new System.Drawing.Point(3, 33);
            this.repoActions.Size = new System.Drawing.Size(854, 34);
            this.repoActions.TabIndex = 1;
            this.createActions.Location = new System.Drawing.Point(3, 463);
            this.createActions.Size = new System.Drawing.Size(854, 34);
            this.createActions.TabIndex = 3;
            this.repositorySearch.Location = new System.Drawing.Point(117, 3);
            this.repositorySearch.Size = new System.Drawing.Size(740, 23);
            this.repositorySearch.TabIndex = 0;
            this.repositoryList.Location = new System.Drawing.Point(3, 73);
            this.repositoryList.Size = new System.Drawing.Size(854, 384);
            this.repositoryList.TabIndex = 2;
            this.repositoryBranch.Location = new System.Drawing.Point(117, 3);
            this.repositoryBranch.Size = new System.Drawing.Size(121, 23);
            this.repositoryBranch.TabIndex = 1;
            this.organization.Location = new System.Drawing.Point(109, 3);
            this.organization.Size = new System.Drawing.Size(121, 23);
            this.organization.TabIndex = 1;
            this.repositoryName.Location = new System.Drawing.Point(3, 3);
            this.repositoryName.Size = new System.Drawing.Size(100, 23);
            this.repositoryName.TabIndex = 0;
            this.privateRepository.Location = new System.Drawing.Point(236, 3);
            this.privateRepository.Size = new System.Drawing.Size(61, 22);
            this.privateRepository.TabIndex = 2;
            this.loadRepositories.Location = new System.Drawing.Point(3, 3);
            this.loadRepositories.Size = new System.Drawing.Size(108, 28);
            this.loadRepositories.TabIndex = 0;
            this.useRepository.Location = new System.Drawing.Point(244, 3);
            this.useRepository.Size = new System.Drawing.Size(93, 28);
            this.useRepository.TabIndex = 2;
            this.createRepository.Location = new System.Drawing.Point(303, 3);
            this.createRepository.Size = new System.Drawing.Size(108, 28);
            this.createRepository.TabIndex = 3;
            this.searchLabel.Name = "searchLabel";
            this.searchLabel.Location = new System.Drawing.Point(3, 0);
            this.searchLabel.Size = new System.Drawing.Size(108, 21);
            this.searchLabel.TabIndex = 4;
            this.repoLayout.ResumeLayout(false);
            this.repoLayout.PerformLayout();
            this.repoActions.ResumeLayout(false);
            this.repoActions.PerformLayout();
            this.createActions.ResumeLayout(false);
            this.createActions.PerformLayout();
            this.loadRepositories.Symbol = VBAi.UiSymbol.Refresh;
            this.loadRepositories.IconOnly = true;
            this.loadRepositories.AutoSize = false;
            this.loadRepositories.MinimumSize = System.Drawing.Size.Empty;
            this.loadRepositories.Size = new System.Drawing.Size(32, 30);
            this.useRepository.Symbol = VBAi.UiSymbol.Check;
            this.createRepository.Symbol = VBAi.UiSymbol.Add;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
