using System;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPane : UserControl
    {
        private string account, remote, branch;
        private CancellationTokenSource cancellation;
        private GitHubRepositoryInfo[] repositories = new GitHubRepositoryInfo[0];
        internal Action<string, string> RepositorySelected;
        internal Action<string, int> OpenModule;
        internal Func<GitPullDraft> LoadDraft;
        private GitHubPull selectedPull;
        internal bool Busy { get { return cancellation != null; } }
        public GitHubPane() { InitializeComponent(); UiText.Apply(this, components); }
        internal void Configure(string selectedAccount, string url, string activeBranch)
        { account = selectedAccount; remote = url; branch = activeBranch; sourceLabel.Text = UiText.Get("Source branch") + ": " + branch; }
        private async Task Run(Func<GitHubApi, CancellationToken, Task> action, bool cancelable = true)
        {
            if (cancellation != null) return;
            using (var pending = new CancellationTokenSource())
            using (var api = new GitHubApi(account))
            {
                cancellation = pending; pages.Enabled = false; cancel.Enabled = cancelable; status.Text = UiText.Get("Connecting to GitHub…");
                try { await action(api, pending.Token); status.Text = UiText.Get("Operation complete."); }
                catch (OperationCanceledException) { status.Text = UiText.Get("Operation cancelled."); }
                catch (Exception ex) { status.Text = ex is HttpRequestException ? UiText.Get("Unable to reach GitHub. Check your connection, then retry.") : ex.Message; }
                finally { cancellation = null; if (!IsDisposed) { pages.Enabled = true; cancel.Enabled = false; } }
            }
        }
        private async void LoadRepositories_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                repositories = await api.Repositories(ct); FilterRepositories(sender, e);
                organization.Items.Clear(); organization.Items.Add(UiText.Get("Personal account"));
                organization.Items.AddRange((await api.Organizations(ct)).Select(x => x.login).ToArray()); organization.SelectedIndex = 0;
            });
        }
        private void FilterRepositories(object sender, EventArgs e)
        {
            repositoryList.Items.Clear(); repositoryList.Items.AddRange(repositories.Where(x => (x.full_name ?? "").IndexOf(repositorySearch.Text, StringComparison.OrdinalIgnoreCase) >= 0).ToArray());
        }
        private async void RepositoryChanged(object sender, EventArgs e)
        {
            var repo = repositoryList.SelectedItem as GitHubRepositoryInfo;
            if (repo == null) return;
            repositoryBranch.Items.Clear(); repositoryBranch.Text = repo.default_branch ?? "main";
            await Run(async (api, ct) => { repositoryBranch.Items.AddRange((await api.Branches(repo.clone_url, ct)).Select(x => x.name).ToArray()); });
        }
        private void UseRepository_Click(object sender, EventArgs e)
        {
            var repo = repositoryList.SelectedItem as GitHubRepositoryInfo;
            if (repo == null) return;
            try { MacroGitRepository.ValidateBranch(repositoryBranch.Text); RepositorySelected?.Invoke(repo.clone_url, repositoryBranch.Text); }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        private async void CreateRepository_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                var created = await api.CreateRepository(repositoryName.Text.Trim(), organization.SelectedIndex > 0 ? organization.Text : null, privateRepository.Checked, ct);
                repositories = repositories.Concat(new[] { created }).ToArray(); repositorySearch.Clear(); FilterRepositories(sender, e);
                repositoryList.SelectedItem = created; repositoryBranch.Text = created.default_branch ?? "main";
                RepositorySelected?.Invoke(created.clone_url, repositoryBranch.Text);
            }, false);
        }
        private async void LoadPulls_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                pulls.Items.Clear(); pulls.Items.AddRange(await api.Pulls(remote, ct));
                targetBranch.Items.Clear(); targetBranch.Items.AddRange((await api.Branches(remote, ct)).Select(x => x.name).ToArray());
                if (targetBranch.Items.Contains("main")) targetBranch.SelectedItem = "main";
            });
        }
        private async void PullChanged(object sender, EventArgs e)
        {
            var pull = pulls.SelectedItem as GitHubPull; if (pull == null) return;
            selectedPull = pull;
            await Run(async (api, ct) => {
                string path = GitHubApi.RepositoryPath(remote);
                var details = await api.Request<GitHubPull>(HttpMethod.Get, path + "/pulls/" + pull.number, null, ct);
                pullDetails.Text = details.title + Environment.NewLine + details.state + (details.merged ? " · " + UiText.Get("Merged") : "") + Environment.NewLine + details.body;
                files.Items.Clear(); files.Items.AddRange(await api.List<GitHubFile>(path + "/pulls/" + pull.number + "/files", ct));
                comments.Items.Clear();
                comments.Items.AddRange(await api.List<GitHubComment>(path + "/issues/" + pull.number + "/comments", ct));
                comments.Items.AddRange(await api.List<GitHubComment>(path + "/pulls/" + pull.number + "/comments", ct));
                try { checks.Text = await api.Checks(remote, details.head.sha, ct); }
                catch (InvalidOperationException ex) { checks.Text = ex.Message; }
            });
        }
        private async void CreatePull_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                var created = await api.CreatePull(remote, branch, targetBranch.Text, pullTitle.Text, pullBody.Text, draft.Checked, ct);
                pulls.Items.Insert(0, created); selectedPull = created; pullDetails.Text = created.html_url;
            }, false);
        }
        private void OpenPull_Click(object sender, EventArgs e)
        {
            if (selectedPull == null) return;
            try { SafeLinks.Open(selectedPull.html_url); } catch (Exception ex) { status.Text = ex.Message; }
        }
        private void OpenFile_Click(object sender, EventArgs e)
        {
            try {
                string path = (files.SelectedItem as GitHubFile)?.filename;
                int line = 1;
                if (sender == comments) { var comment = comments.SelectedItem as GitHubComment; path = comment?.path; line = comment?.line ?? 1; }
                if (path == null || !path.StartsWith("vba/", StringComparison.Ordinal) || path.Substring(4).Contains("/")) return;
                string name = System.IO.Path.GetFileNameWithoutExtension(path);
                OpenModule?.Invoke(name, line);
            } catch (Exception ex) { status.Text = ex.Message; }
        }
        private void CommentChanged(object sender, EventArgs e) { commentBody.Text = (comments.SelectedItem as GitHubComment)?.body ?? ""; }
        private void LoadDraft_Click(object sender, EventArgs e)
        {
            try {
                var prepared = LoadDraft?.Invoke();
                if (prepared == null) { status.Text = UiText.Get("No prepared draft for this branch."); return; }
                targetBranch.Text = prepared.Target; pullTitle.Text = prepared.Title; pullBody.Text = prepared.Body;
                pullTabs.SelectedTab = composeTab;
            } catch (Exception ex) { status.Text = ex.Message; }
        }
        private void Cancel_Click(object sender, EventArgs e) { cancellation?.Cancel(); }
        protected override void Dispose(bool disposing) { if (disposing) { cancellation?.Cancel(); components?.Dispose(); } base.Dispose(disposing); }
    }
}
