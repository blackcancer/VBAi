namespace CodexVBE
{
    public sealed partial class GitHubPullRequestsView
    {
        internal System.Windows.Forms.TableLayoutPanel pullLayout;
        internal System.Windows.Forms.FlowLayoutPanel pullActions;
        internal System.Windows.Forms.TabControl pullTabs;
        internal System.Windows.Forms.TabPage composeTab;
        internal System.Windows.Forms.TabPage detailTab;
        internal System.Windows.Forms.TabPage filesTab;
        internal System.Windows.Forms.TabPage commentsTab;
        internal System.Windows.Forms.TabPage checksTab;
        internal System.Windows.Forms.ListBox pulls;
        internal System.Windows.Forms.Button loadPulls;
        internal System.Windows.Forms.Button openPull;
        internal System.Windows.Forms.Button loadDraft;
        internal CodexVBE.GitHubPullComposeView gitHubPullComposeView;
        internal CodexVBE.GitHubPullDetailsView gitHubPullDetailsView;
        internal CodexVBE.GitHubPullFilesView gitHubPullFilesView;
        internal CodexVBE.GitHubPullCommentsView gitHubPullCommentsView;
        internal CodexVBE.GitHubPullChecksView gitHubPullChecksView;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.gitHubPullComposeView = new CodexVBE.GitHubPullComposeView();
            this.gitHubPullDetailsView = new CodexVBE.GitHubPullDetailsView();
            this.gitHubPullFilesView = new CodexVBE.GitHubPullFilesView();
            this.gitHubPullCommentsView = new CodexVBE.GitHubPullCommentsView();
            this.gitHubPullChecksView = new CodexVBE.GitHubPullChecksView();
            this.pullLayout = new System.Windows.Forms.TableLayoutPanel();
            this.pullActions = new System.Windows.Forms.FlowLayoutPanel();
            this.pullTabs = new CodexVBE.ThemedTabControl();
            this.composeTab = new System.Windows.Forms.TabPage();
            this.detailTab = new System.Windows.Forms.TabPage();
            this.filesTab = new System.Windows.Forms.TabPage();
            this.commentsTab = new System.Windows.Forms.TabPage();
            this.checksTab = new System.Windows.Forms.TabPage();
            this.pulls = new System.Windows.Forms.ListBox();
            this.loadPulls = new CodexVBE.ThemedButton();
            this.openPull = new CodexVBE.ThemedButton();
            this.loadDraft = new CodexVBE.ThemedButton();
            this.SuspendLayout();
            this.loadDraft.Name = "loadDraft";
            this.loadDraft.Text = "Load prepared draft";
            this.loadDraft.AutoSize = true;
            this.pullActions.Controls.Add(this.loadDraft);
            this.pullLayout.Name = "pullLayout";
            this.pullActions.Name = "pullActions";
            this.pullTabs.Name = "pullTabs";
            this.composeTab.Name = "composeTab";
            this.detailTab.Name = "detailTab";
            this.filesTab.Name = "filesTab";
            this.commentsTab.Name = "commentsTab";
            this.checksTab.Name = "checksTab";
            this.pulls.Name = "pulls";
            this.loadPulls.Name = "loadPulls";
            this.openPull.Name = "openPull";
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
            this.pullActions.Controls.Add(this.loadPulls);
            this.toolTips.SetToolTip(this.loadPulls, "Load pull requests");
            this.openPull.Text = "Open on GitHub";
            this.openPull.AutoSize = true;
            this.pullActions.Controls.Add(this.openPull);
            this.toolTips.SetToolTip(this.openPull, "Open on GitHub");
            this.pullActions.AutoSize = true;
            this.pulls.HorizontalScrollbar = true;
            this.composeTab.Text = "New pull request";
            this.pullTabs.Controls.Add(this.composeTab);
            this.detailTab.Text = "Details";
            this.pullTabs.Controls.Add(this.detailTab);
            this.filesTab.Text = "Files";
            this.pullTabs.Controls.Add(this.filesTab);
            this.commentsTab.Text = "Comments";
            this.pullTabs.Controls.Add(this.commentsTab);
            this.checksTab.Text = "Checks";
            this.pullTabs.Controls.Add(this.checksTab);
            this.Controls.Add(this.pullLayout);
            this.gitHubPullComposeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeTab.Controls.Add(this.gitHubPullComposeView);
            this.gitHubPullDetailsView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.detailTab.Controls.Add(this.gitHubPullDetailsView);
            this.gitHubPullFilesView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.filesTab.Controls.Add(this.gitHubPullFilesView);
            this.gitHubPullCommentsView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentsTab.Controls.Add(this.gitHubPullCommentsView);
            this.gitHubPullChecksView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checksTab.Controls.Add(this.gitHubPullChecksView);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullRequestsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
