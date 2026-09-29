namespace VBAi
{
    /// <summary>Control references from embedded repository, pull-request, detail, and review views.</summary>
    public sealed partial class GitHubPane
    {
        /// <summary>Root layout for repository search and selection.</summary>
        private System.Windows.Forms.TableLayoutPanel repoLayout;
        /// <summary>Actions for loading or selecting a repository.</summary>
        private System.Windows.Forms.FlowLayoutPanel repoActions;
        /// <summary>Actions for creating a repository.</summary>
        private System.Windows.Forms.FlowLayoutPanel createActions;
        /// <summary>Repository filter input.</summary>
        private System.Windows.Forms.TextBox repositorySearch;
        /// <summary>Filtered repository choices.</summary>
        private System.Windows.Forms.ListBox repositoryList;
        /// <summary>Branches available for the selected repository.</summary>
        private System.Windows.Forms.ComboBox repositoryBranch;
        /// <summary>Organization that owns a new repository.</summary>
        private System.Windows.Forms.ComboBox organization;
        /// <summary>Name for a new repository.</summary>
        private System.Windows.Forms.TextBox repositoryName;
        /// <summary>Whether a new repository should be private.</summary>
        private System.Windows.Forms.CheckBox privateRepository;
        /// <summary>Loads repositories available to the connected account.</summary>
        private System.Windows.Forms.Button loadRepositories;
        /// <summary>Selects the highlighted repository for the current Git project.</summary>
        private System.Windows.Forms.Button useRepository;
        /// <summary>Creates the configured repository.</summary>
        private System.Windows.Forms.Button createRepository;
        /// <summary>Caption for repository search input.</summary>
        private System.Windows.Forms.Label searchLabel;
        /// <summary>Root layout for pull-request lists and detail pages.</summary>
        private System.Windows.Forms.TableLayoutPanel pullLayout;
        /// <summary>Actions for listing and opening pull requests.</summary>
        private System.Windows.Forms.FlowLayoutPanel pullActions;
        /// <summary>Tabs for pull composition, details, files, comments, and checks.</summary>
        private System.Windows.Forms.TabControl pullTabs;
        /// <summary>Pull-request composition page.</summary>
        private System.Windows.Forms.TabPage composeTab;
        /// <summary>Pull-request details page.</summary>
        private System.Windows.Forms.TabPage detailTab;
        /// <summary>Changed files page.</summary>
        private System.Windows.Forms.TabPage filesTab;
        /// <summary>Review comments page.</summary>
        private System.Windows.Forms.TabPage commentsTab;
        /// <summary>CI checks page.</summary>
        private System.Windows.Forms.TabPage checksTab;
        /// <summary>Pull requests for the selected repository.</summary>
        private System.Windows.Forms.ListBox pulls;
        /// <summary>Loads pull requests for the selected repository.</summary>
        private System.Windows.Forms.Button loadPulls;
        /// <summary>Opens the selected pull request detail pages.</summary>
        private System.Windows.Forms.Button openPull;
        /// <summary>Loads a saved pull-request draft.</summary>
        private System.Windows.Forms.Button loadDraft;
        /// <summary>Layout for creating a pull request.</summary>
        private System.Windows.Forms.TableLayoutPanel composeLayout;
        /// <summary>Target branch for a new pull request.</summary>
        private System.Windows.Forms.ComboBox targetBranch;
        /// <summary>Caption describing the source branch.</summary>
        private System.Windows.Forms.Label sourceLabel;
        /// <summary>New pull-request title input.</summary>
        private System.Windows.Forms.TextBox pullTitle;
        /// <summary>New pull-request description input.</summary>
        private System.Windows.Forms.TextBox pullBody;
        /// <summary>Whether the new pull request is created as a draft.</summary>
        private System.Windows.Forms.CheckBox draft;
        /// <summary>Creates the configured pull request.</summary>
        private System.Windows.Forms.Button createPull;
        /// <summary>Caption for the pull-request target branch.</summary>
        private System.Windows.Forms.Label targetLabel;
        /// <summary>Caption for the pull-request title.</summary>
        private System.Windows.Forms.Label titleLabel;
        /// <summary>Caption for the pull-request description.</summary>
        private System.Windows.Forms.Label bodyLabel;
        /// <summary>Details for the selected pull request.</summary>
        private System.Windows.Forms.TextBox pullDetails;
        /// <summary>Files changed by the selected pull request.</summary>
        private System.Windows.Forms.ListBox files;
        /// <summary>Layout for review comments and their text.</summary>
        private System.Windows.Forms.TableLayoutPanel commentLayout;
        /// <summary>Review comments for the selected pull request.</summary>
        private System.Windows.Forms.ListBox comments;
        /// <summary>Selected review comment text.</summary>
        private System.Windows.Forms.TextBox commentBody;
        /// <summary>Checks reported for the selected pull request.</summary>
        private System.Windows.Forms.TextBox checks;
        /// <summary>Associe les contrôles du volet à ceux des vues spécialisées et branche leurs interactions.</summary>
        private void BindViews()
        {
            repoLayout = gitHubRepositoriesView.repoLayout;
            repoActions = gitHubRepositoriesView.repoActions;
            createActions = gitHubRepositoriesView.createActions;
            repositorySearch = gitHubRepositoriesView.repositorySearch;
            repositoryList = gitHubRepositoriesView.repositoryList;
            repositoryBranch = gitHubRepositoriesView.repositoryBranch;
            organization = gitHubRepositoriesView.organization;
            repositoryName = gitHubRepositoriesView.repositoryName;
            privateRepository = gitHubRepositoriesView.privateRepository;
            loadRepositories = gitHubRepositoriesView.loadRepositories;
            useRepository = gitHubRepositoriesView.useRepository;
            createRepository = gitHubRepositoriesView.createRepository;
            searchLabel = gitHubRepositoriesView.searchLabel;
            pullLayout = gitHubPullRequestsView.pullLayout;
            pullActions = gitHubPullRequestsView.pullActions;
            pullTabs = gitHubPullRequestsView.pullTabs;
            composeTab = gitHubPullRequestsView.composeTab;
            detailTab = gitHubPullRequestsView.detailTab;
            filesTab = gitHubPullRequestsView.filesTab;
            commentsTab = gitHubPullRequestsView.commentsTab;
            checksTab = gitHubPullRequestsView.checksTab;
            pulls = gitHubPullRequestsView.pulls;
            loadPulls = gitHubPullRequestsView.loadPulls;
            openPull = gitHubPullRequestsView.openPull;
            loadDraft = gitHubPullRequestsView.loadDraft;
            composeLayout = gitHubPullRequestsView.gitHubPullComposeView.composeLayout;
            targetBranch = gitHubPullRequestsView.gitHubPullComposeView.targetBranch;
            sourceLabel = gitHubPullRequestsView.gitHubPullComposeView.sourceLabel;
            pullTitle = gitHubPullRequestsView.gitHubPullComposeView.pullTitle;
            pullBody = gitHubPullRequestsView.gitHubPullComposeView.pullBody;
            draft = gitHubPullRequestsView.gitHubPullComposeView.draft;
            createPull = gitHubPullRequestsView.gitHubPullComposeView.createPull;
            targetLabel = gitHubPullRequestsView.gitHubPullComposeView.targetLabel;
            titleLabel = gitHubPullRequestsView.gitHubPullComposeView.titleLabel;
            bodyLabel = gitHubPullRequestsView.gitHubPullComposeView.bodyLabel;
            pullDetails = gitHubPullRequestsView.gitHubPullDetailsView.pullDetails;
            files = gitHubPullRequestsView.gitHubPullFilesView.files;
            commentLayout = gitHubPullRequestsView.gitHubPullCommentsView.commentLayout;
            comments = gitHubPullRequestsView.gitHubPullCommentsView.comments;
            commentBody = gitHubPullRequestsView.gitHubPullCommentsView.commentBody;
            checks = gitHubPullRequestsView.gitHubPullChecksView.checks;
            this.loadDraft.Click += new System.EventHandler(this.LoadDraft_Click);
            this.repositorySearch.TextChanged += new System.EventHandler(this.FilterRepositories);
            this.loadRepositories.Click += new System.EventHandler(this.LoadRepositories_Click);
            this.useRepository.Click += new System.EventHandler(this.UseRepository_Click);
            this.repositoryList.SelectedIndexChanged += new System.EventHandler(this.RepositoryChanged);
            this.createRepository.Click += new System.EventHandler(this.CreateRepository_Click);
            this.loadPulls.Click += new System.EventHandler(this.LoadPulls_Click);
            this.openPull.Click += new System.EventHandler(this.OpenPull_Click);
            this.pulls.SelectedIndexChanged += new System.EventHandler(this.PullChanged);
            this.createPull.Click += new System.EventHandler(this.CreatePull_Click);
            this.comments.SelectedIndexChanged += new System.EventHandler(this.CommentChanged);
            this.comments.DoubleClick += new System.EventHandler(this.OpenFile_Click);
            this.files.DoubleClick += new System.EventHandler(this.OpenFile_Click);
        }
    }
}
