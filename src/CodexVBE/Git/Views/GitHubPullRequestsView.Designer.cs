namespace CodexVBE
{
    /// <summary>Designer-generated pull-request list, composition, and review pages.</summary>
    public sealed partial class GitHubPullRequestsView
    {
        /// <summary>Root layout for pull list and review tabs.</summary>
        internal System.Windows.Forms.TableLayoutPanel pullLayout;
        /// <summary>Actions for loading and opening pull requests.</summary>
        internal System.Windows.Forms.FlowLayoutPanel pullActions;
        /// <summary>Composition and review page selector.</summary>
        internal CodexVBE.ThemedTabControl pullTabs;
        /// <summary>Pull-request composition page.</summary>
        internal System.Windows.Forms.TabPage composeTab;
        /// <summary>Pull-request detail page.</summary>
        internal System.Windows.Forms.TabPage detailTab;
        /// <summary>Changed-files page.</summary>
        internal System.Windows.Forms.TabPage filesTab;
        /// <summary>Review-comments page.</summary>
        internal System.Windows.Forms.TabPage commentsTab;
        /// <summary>Continuous-integration checks page.</summary>
        internal System.Windows.Forms.TabPage checksTab;
        /// <summary>Pull requests for the selected repository.</summary>
        internal System.Windows.Forms.ListBox pulls;
        /// <summary>Loads pull requests from GitHub.</summary>
        internal CodexVBE.ThemedButton loadPulls;
        /// <summary>Opens the selected pull request.</summary>
        internal CodexVBE.ThemedButton openPull;
        /// <summary>Loads a saved pull-request draft.</summary>
        internal CodexVBE.ThemedButton loadDraft;
        /// <summary>Pull-request composition editor.</summary>
        internal CodexVBE.GitHubPullComposeView gitHubPullComposeView;
        /// <summary>Selected pull-request summary and metadata.</summary>
        internal CodexVBE.GitHubPullDetailsView gitHubPullDetailsView;
        /// <summary>Changed files for the selected pull request.</summary>
        internal CodexVBE.GitHubPullFilesView gitHubPullFilesView;
        /// <summary>Review comments for the selected pull request.</summary>
        internal CodexVBE.GitHubPullCommentsView gitHubPullCommentsView;
        /// <summary>CI check results for the selected pull request.</summary>
        internal CodexVBE.GitHubPullChecksView gitHubPullChecksView;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with pull-request actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges pull-request pages and nested views.</summary>
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
            this.pullLayout.SuspendLayout();
            this.pullActions.SuspendLayout();
            this.pullTabs.SuspendLayout();
            this.composeTab.SuspendLayout();
            this.detailTab.SuspendLayout();
            this.filesTab.SuspendLayout();
            this.commentsTab.SuspendLayout();
            this.checksTab.SuspendLayout();
            this.gitHubPullComposeView.SuspendLayout();
            this.gitHubPullDetailsView.SuspendLayout();
            this.gitHubPullFilesView.SuspendLayout();
            this.gitHubPullCommentsView.SuspendLayout();
            this.gitHubPullChecksView.SuspendLayout();
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
            this.pullLayout.Location = new System.Drawing.Point(0, 0);
            this.pullLayout.Size = new System.Drawing.Size(860, 500);
            this.pullLayout.TabIndex = 0;
            this.pullActions.Location = new System.Drawing.Point(3, 3);
            this.pullActions.Size = new System.Drawing.Size(854, 34);
            this.pullActions.TabIndex = 0;
            this.pullTabs.Location = new System.Drawing.Point(3, 153);
            this.pullTabs.Size = new System.Drawing.Size(854, 344);
            this.pullTabs.TabIndex = 2;
            this.composeTab.Location = new System.Drawing.Point(0, 0);
            this.composeTab.Size = new System.Drawing.Size(200, 100);
            this.composeTab.TabIndex = 0;
            this.detailTab.Location = new System.Drawing.Point(0, 0);
            this.detailTab.Size = new System.Drawing.Size(200, 100);
            this.detailTab.TabIndex = 1;
            this.filesTab.Location = new System.Drawing.Point(0, 0);
            this.filesTab.Size = new System.Drawing.Size(200, 100);
            this.filesTab.TabIndex = 2;
            this.commentsTab.Location = new System.Drawing.Point(0, 0);
            this.commentsTab.Size = new System.Drawing.Size(200, 100);
            this.commentsTab.TabIndex = 3;
            this.checksTab.Location = new System.Drawing.Point(0, 0);
            this.checksTab.Size = new System.Drawing.Size(200, 100);
            this.checksTab.TabIndex = 4;
            this.pulls.Location = new System.Drawing.Point(3, 43);
            this.pulls.Size = new System.Drawing.Size(854, 104);
            this.pulls.TabIndex = 1;
            this.loadPulls.Location = new System.Drawing.Point(132, 3);
            this.loadPulls.Size = new System.Drawing.Size(114, 28);
            this.loadPulls.TabIndex = 1;
            this.openPull.Location = new System.Drawing.Point(252, 3);
            this.openPull.Size = new System.Drawing.Size(104, 28);
            this.openPull.TabIndex = 2;
            this.loadDraft.Location = new System.Drawing.Point(3, 3);
            this.loadDraft.Size = new System.Drawing.Size(123, 28);
            this.loadDraft.TabIndex = 0;
            this.gitHubPullComposeView.Name = "gitHubPullComposeView";
            this.gitHubPullComposeView.Location = new System.Drawing.Point(0, 0);
            this.gitHubPullComposeView.Size = new System.Drawing.Size(200, 100);
            this.gitHubPullComposeView.TabIndex = 0;
            this.gitHubPullDetailsView.Name = "gitHubPullDetailsView";
            this.gitHubPullDetailsView.Location = new System.Drawing.Point(0, 0);
            this.gitHubPullDetailsView.Size = new System.Drawing.Size(200, 100);
            this.gitHubPullDetailsView.TabIndex = 0;
            this.gitHubPullFilesView.Name = "gitHubPullFilesView";
            this.gitHubPullFilesView.Location = new System.Drawing.Point(0, 0);
            this.gitHubPullFilesView.Size = new System.Drawing.Size(200, 100);
            this.gitHubPullFilesView.TabIndex = 0;
            this.gitHubPullCommentsView.Name = "gitHubPullCommentsView";
            this.gitHubPullCommentsView.Location = new System.Drawing.Point(0, 0);
            this.gitHubPullCommentsView.Size = new System.Drawing.Size(200, 100);
            this.gitHubPullCommentsView.TabIndex = 0;
            this.gitHubPullChecksView.Name = "gitHubPullChecksView";
            this.gitHubPullChecksView.Location = new System.Drawing.Point(0, 0);
            this.gitHubPullChecksView.Size = new System.Drawing.Size(200, 100);
            this.gitHubPullChecksView.TabIndex = 0;
            this.pullLayout.ResumeLayout(false);
            this.pullLayout.PerformLayout();
            this.pullActions.ResumeLayout(false);
            this.pullActions.PerformLayout();
            this.pullTabs.ResumeLayout(false);
            this.pullTabs.PerformLayout();
            this.composeTab.ResumeLayout(false);
            this.composeTab.PerformLayout();
            this.detailTab.ResumeLayout(false);
            this.detailTab.PerformLayout();
            this.filesTab.ResumeLayout(false);
            this.filesTab.PerformLayout();
            this.commentsTab.ResumeLayout(false);
            this.commentsTab.PerformLayout();
            this.checksTab.ResumeLayout(false);
            this.checksTab.PerformLayout();
            this.gitHubPullComposeView.ResumeLayout(false);
            this.gitHubPullComposeView.PerformLayout();
            this.gitHubPullDetailsView.ResumeLayout(false);
            this.gitHubPullDetailsView.PerformLayout();
            this.gitHubPullFilesView.ResumeLayout(false);
            this.gitHubPullFilesView.PerformLayout();
            this.gitHubPullCommentsView.ResumeLayout(false);
            this.gitHubPullCommentsView.PerformLayout();
            this.gitHubPullChecksView.ResumeLayout(false);
            this.gitHubPullChecksView.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
