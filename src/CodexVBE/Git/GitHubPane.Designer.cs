namespace CodexVBE
{
    public sealed partial class GitHubPane
    {
        private System.Windows.Forms.TabControl pages;
        private System.Windows.Forms.TabPage repositoriesPage;
        private System.Windows.Forms.TabPage pullsPage;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.FlowLayoutPanel footer;
        private System.Windows.Forms.ToolTip tips;
        private System.ComponentModel.IContainer components;
        private CodexVBE.GitHubRepositoriesView gitHubRepositoriesView;
        private CodexVBE.GitHubPullRequestsView gitHubPullRequestsView;
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.gitHubPullRequestsView = new CodexVBE.GitHubPullRequestsView();
            this.gitHubRepositoriesView = new CodexVBE.GitHubRepositoriesView();
            this.pages = new CodexVBE.ThemedTabControl();
            this.repositoriesPage = new System.Windows.Forms.TabPage();
            this.pullsPage = new System.Windows.Forms.TabPage();
            this.status = new System.Windows.Forms.Label();
            this.cancel = new CodexVBE.ThemedButton();
            this.footer = new System.Windows.Forms.FlowLayoutPanel();
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.pages.Name = "pages";
            this.repositoriesPage.Name = "repositoriesPage";
            this.pullsPage.Name = "pullsPage";
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
            this.repositoriesPage.Text = "Repositories";
            this.pages.Controls.Add(this.repositoriesPage);
            this.pullsPage.Text = "Pull requests";
            this.pages.Controls.Add(this.pullsPage);
            this.gitHubRepositoriesView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.repositoriesPage.Controls.Add(this.gitHubRepositoriesView);
            this.gitHubPullRequestsView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pullsPage.Controls.Add(this.gitHubPullRequestsView);
            this.Controls.Add(this.pages);
            this.Controls.Add(this.footer);
            this.Name = "GitHubPane";
            this.Size = new System.Drawing.Size(860, 500);
        }
    }
}
