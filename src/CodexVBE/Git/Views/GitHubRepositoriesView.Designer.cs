namespace CodexVBE
{
    public sealed partial class GitHubRepositoriesView
    {
        internal System.Windows.Forms.TableLayoutPanel repoLayout;
        internal System.Windows.Forms.FlowLayoutPanel repoActions;
        internal System.Windows.Forms.FlowLayoutPanel createActions;
        internal System.Windows.Forms.TextBox repositorySearch;
        internal System.Windows.Forms.ListBox repositoryList;
        internal System.Windows.Forms.ComboBox repositoryBranch;
        internal System.Windows.Forms.ComboBox organization;
        internal System.Windows.Forms.TextBox repositoryName;
        internal System.Windows.Forms.CheckBox privateRepository;
        internal System.Windows.Forms.Button loadRepositories;
        internal System.Windows.Forms.Button useRepository;
        internal System.Windows.Forms.Button createRepository;
        internal System.Windows.Forms.Label searchLabel;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.repoLayout = new System.Windows.Forms.TableLayoutPanel();
            this.repoActions = new System.Windows.Forms.FlowLayoutPanel();
            this.createActions = new System.Windows.Forms.FlowLayoutPanel();
            this.repositorySearch = new System.Windows.Forms.TextBox();
            this.repositoryList = new System.Windows.Forms.ListBox();
            this.repositoryBranch = new CodexVBE.ThemedComboBox();
            this.organization = new CodexVBE.ThemedComboBox();
            this.repositoryName = new System.Windows.Forms.TextBox();
            this.privateRepository = new System.Windows.Forms.CheckBox();
            this.loadRepositories = new CodexVBE.ThemedButton();
            this.useRepository = new CodexVBE.ThemedButton();
            this.createRepository = new CodexVBE.ThemedButton();
            this.searchLabel = new System.Windows.Forms.Label();
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
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
