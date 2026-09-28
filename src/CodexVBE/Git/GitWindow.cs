using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class GitWindow : Form
    {
        private VbaGitProject project;
        private MacroGitRepository repository;
        private string cache;
        private FileStream cacheLock;
        private bool running;
        private string account;
        private VbaGitSnapshot displayedLive;
        private VbaGitSnapshot displayedBaseline;
        private string displayedBranch;
        public GitWindow() { InitializeComponent(); BindViews(); Icon = VbeWindowIcons.Icon("github"); UiText.Apply(this, components); InitializeReview(); }

        internal GitWindow(VbaGitProject project, string scope, string label, string account = null) : this()
        {
            this.account = account;
            this.project = project;
            githubPane.Configure(account, "", "main");
            cache = MacroGitRepository.ScopeDirectory(scope);
            documentLabel.Text = label;
            Directory.CreateDirectory(cache);
            string file = Path.Combine(cache, "binding.json");
            if (File.Exists(file))
            {
                var binding = new JavaScriptSerializer().Deserialize<Binding>(File.ReadAllText(file));
                remote.Text = binding.Remote; branch.Text = binding.Branch;
                githubPane.Configure(account, remote.Text, branch.Text);
            }
        }

        private sealed class Binding { public string Remote { get; set; } public string Branch { get; set; } }

        private async void Connect_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                string url = MacroGitRepository.ValidateRemote(remote.Text);
                string bindingId;
                using (var hash = System.Security.Cryptography.SHA256.Create())
                    bindingId = BitConverter.ToString(hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(url + "\n" + branch.Text.Trim()))).Replace("-", "");
                var selected = new MacroGitRepository(Path.Combine(cache, bindingId + ".git"), branch.Text.Trim(), account);
                selected.Cancellation = operationCancellation.Token; selected.Progress = ReportProgress; cancelOperation.Enabled = true;
                try { await Task.Run(() => { selected.Initialize(url); selected.Fetch(); }); }
                finally { selected.Cancellation = System.Threading.CancellationToken.None; selected.Progress = null; }
                repository = selected;
                string bindingFile = Path.Combine(cache, "binding.json");
                string temporaryBinding = Path.Combine(cache, "binding.pending");
                File.WriteAllText(temporaryBinding, new JavaScriptSerializer().Serialize(
                    // Keep the original cache key stable when Initialize restores another active branch.
                    new Binding { Remote = url, Branch = branch.Text.Trim() }));
                if (File.Exists(bindingFile)) File.Replace(temporaryBinding, bindingFile, null);
                else File.Move(temporaryBinding, bindingFile);
                githubPane.Configure(account, remote.Text, repository.Branch);
                await Compare();
                tabs.SelectedTab = changesTab;
            });
        }

        private async void Compare_Click(object sender, EventArgs e) { await Perform(Compare, true); }
        private async Task Compare()
        {
            var live = project.Capture();
            displayedBranch = repository.Branch;
            var baseline = await Task.Run(() => repository.Read(repository.Resolve(MacroGitRepository.Baseline)));
            displayedLive = live; displayedBaseline = baseline;
            reviewCommit = null;
            PopulateChanges(live, baseline);
            history.Items.Clear(); history.Items.AddRange(await Task.Run(() => repository.Commits()));
            branch.Text = repository.Branch;
            branchList.Items.Clear(); branchList.Items.AddRange(await Task.Run(() => repository.Branches()));
            checkpointList.Items.Clear(); checkpointList.Items.AddRange(await Task.Run(() => repository.Checkpoints()));
            conflictList.Items.Clear();
            if (repository.PendingMerge != null) conflictList.Items.AddRange(repository.PendingMerge.Conflicts);
            syncStatus.Text = repository.Branch + "   ·   " + await Task.Run(() => repository.SynchronizationStatus());
            status.Text = repository.RecoveryPending ? UiText.Get("Import interrupted: restore VBA before continuing. The backup is preserved in the cache.") :
                baseline == null ? UiText.Get("First link: commit then push to publish, or pull to import the repository with a backup first.") :
                changes.Items.Count == 0 ? UiText.Get("VBA matches the last synchronized state.") : changes.Items.Count + UiText.Get(" file(s) changed since the last synchronization.");
        }

        private async void Commit_Click(object sender, EventArgs e) { await RunGitAction("commit_selected", text: commitMessage.Text); }
        private async void Push_Click(object sender, EventArgs e) { await RunGitAction("push"); }
        private async void Fetch_Click(object sender, EventArgs e) { await RunGitAction("fetch"); }
        private async Task RunGitAction(string action, string name = null, string text = null, string choice = null, string path = null)
        {
            await Perform(async () => {
                var operations = new MacroGitOperations(project, repository) { ImportPreview = ShowImportSummary };
                if (action == "commit_selected" && !repository.RecoveryPending && displayedLive != null && !project.Capture().SameAs(displayedLive))
                    throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                repository.Progress = ReportProgress;
                repository.Cancellation = action == "fetch" ? operationCancellation.Token : System.Threading.CancellationToken.None;
                cancelOperation.Enabled = action == "fetch";
                string[] selected = changes.CheckedItems.Cast<ModuleChange>().Where(x => x.Name != null).Select(x => x.Name).ToArray();
                object result = await operations.ExecuteAsync(action, name: name, text: text, choice: choice, path: path, modules: selected, references: changes.CheckedItems.Cast<ModuleChange>().Any(x => x.Name == null));
                repository.Cancellation = System.Threading.CancellationToken.None;
                await Compare();
                status.Text = (action == "commit" || action == "commit_selected") ? UiText.Get("Local commit created. Use Push to publish it.") :
                    action == "push" ? UiText.Get("Push complete.") : action == "pull" ? UiText.Get("Pull and import complete. Check and save the document.") :
                    action == "rollback" || action == "checkpoint_restore" ? UiText.Get("VBA restored. Check and save the document.") :
                    repository.PendingMerge != null ? UiText.Get("Merge prepared: resolve conflicts, then click Complete merge.") : UiText.Get("Operation complete: ") + action;
                if (repository.PendingMerge != null) tabs.SelectedTab = conflictsTab;
            });
        }
        private void Changes_SelectedIndexChanged(object sender, EventArgs e)
        {
            var item = changes.SelectedItem as ModuleChange;
            if (item == null) { diff.ShowDiff("", ""); return; }
            diff.ShowDiff(Source(displayedBaseline, item.Name), Source(displayedLive, item.Name));
        }

        private async void Pull_Click(object sender, EventArgs e) { await RunGitAction("pull"); }
        private async void Restore_Click(object sender, EventArgs e) { await RunGitAction("rollback"); }
        private async void CheckpointCreate_Click(object sender, EventArgs e) { await RunGitAction("checkpoint_create", checkpointName.Text); }
        private async void CheckpointRestore_Click(object sender, EventArgs e)
        {
            var selected = checkpointList.SelectedItem as GitCheckpoint;
            if (selected != null) await RunGitAction("checkpoint_restore", selected.Id);
        }
        private async void BranchCreate_Click(object sender, EventArgs e) { await RunGitAction("branch_create", branchName.Text); }
        private async void BranchTrack_Click(object sender, EventArgs e) { await RunGitAction("branch_track", branchName.Text); }
        private async void RemoteBranches_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                string[] names = await Task.Run(() => repository.RemoteBranches());
                branchName.Items.Clear(); branchName.Items.AddRange(names);
                status.Text = names.Length + UiText.Get(" remote branch(es). Choose a name, then Track remote.");
            });
            if (!running && branchName.Items.Count > 0) branchName.DroppedDown = true;
        }
        private async void BranchSwitch_Click(object sender, EventArgs e)
        { if (branchList.SelectedItem != null) await RunGitAction("branch_switch", branchList.SelectedItem.ToString()); }
        private async void MergeBegin_Click(object sender, EventArgs e)
        { if (branchList.SelectedItem != null) await RunGitAction("merge_begin", branchList.SelectedItem.ToString()); }
        private async void MergeOurs_Click(object sender, EventArgs e) { await ResolveMerge("ours"); }
        private async void MergeTheirs_Click(object sender, EventArgs e) { await ResolveMerge("theirs"); }
        private async void MergeText_Click(object sender, EventArgs e) { await ResolveMerge("text"); }
        private async Task ResolveMerge(string choice)
        { if (conflictList.SelectedItem != null) await RunGitAction("merge_resolve", text: resolutionText.Text, choice: choice, path: conflictList.SelectedItem.ToString()); }
        private async void MergeComplete_Click(object sender, EventArgs e) { await RunGitAction("merge_complete", text: commitMessage.Text); }
        private async void MergeAbort_Click(object sender, EventArgs e) { await RunGitAction("merge_abort"); }
        private async void ConflictList_SelectedIndexChanged(object sender, EventArgs e)
        {
            conflictDiff.Rows.Clear(); resolutionText.Clear();
            if (running || repository == null || conflictList.SelectedItem == null) return;
            string path = conflictList.SelectedItem.ToString();
            await Perform(async () => {
                var content = await Task.Run(() => repository.ConflictContent(path));
                if (conflictList.SelectedItem?.ToString() != path) return;
                baseContent.Text = content.Base;
                resolutionText.Text = content.Ours.StartsWith("[", StringComparison.Ordinal) ? "" : content.Ours;
                foreach (var row in CodeChange.BuildRows(content.Ours, content.Theirs).Take(2000))
                {
                    int index = conflictDiff.Rows.Add(row.Kind == CodeDiffKind.Added ? "" : row.Text, row.Kind == CodeDiffKind.Removed ? "" : row.Text);
                    if (row.Kind == CodeDiffKind.Added) conflictDiff.Rows[index].Cells[1].Tag = "vba-added";
                    if (row.Kind == CodeDiffKind.Removed) conflictDiff.Rows[index].Cells[0].Tag = "vba-removed";
                }
                status.Text = UiText.Get("Conflict: ") + path + UiText.Get(". Choose a version or edit the full content below.");
            });
        }
        private async Task Perform(Func<Task> action, bool refresh = false)
        {
            if (running || project == null) return;
            running = true; operationCancellation = new System.Threading.CancellationTokenSource(); operationProgress.Visible = true; cancelOperation.Visible = true; UpdateButtons(); status.Text = UiText.Get("Operation in progress…");
            try {
                // Lock only the transaction so the chat agent can use Git while this window is idle.
                if (cache != null) cacheLock = new FileStream(Path.Combine(cache, "session.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                if (repository != null && !string.IsNullOrWhiteSpace(remote.Text)) {
                    string repositoryUrl = MacroGitRepository.ValidateRemote(remote.Text);
                    await Task.Run(() => repository.Initialize(repositoryUrl));
                    if (!refresh && displayedBranch != null && displayedBranch != repository.Branch)
                        throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes.") + " " + UiText.Get("Compare returns to live VBA."));
                }
                await action();
            }
            catch (OperationCanceledException) { status.Text = UiText.Get("Operation cancelled."); }
            catch (Exception ex) { status.Text = ex.Message + " · " + UiText.Get("Check the connection, account and Git state, then retry."); }
            finally { cacheLock?.Dispose(); cacheLock = null; if (repository != null) { repository.Cancellation = System.Threading.CancellationToken.None; repository.Progress = null; } operationCancellation.Dispose(); operationCancellation = null; operationProgress.Visible = false; cancelOperation.Visible = false; cancelOperation.Enabled = false; running = false; UpdateButtons(); }
        }
        private void UpdateButtons()
        {
            connect.Enabled = !running && repository == null;
            remote.ReadOnly = branch.ReadOnly = repository != null || running;
            compare.Enabled = commit.Enabled = fetch.Enabled = push.Enabled = pull.Enabled = restore.Enabled = !running && repository != null;
            commit.Enabled = !running && repository != null && reviewCommit == null;
            githubPane.Enabled = !running;
            previewImport.Enabled = openModule.Enabled = restoreModule.Enabled = !running && repository != null;
            restoreModule.Enabled = restoreModule.Enabled && reviewCommit != null;
            if (repository != null) githubPane.Configure(account, remote.Text, repository.Branch);
            commitMessage.Enabled = !running;
            branchesTab.Enabled = checkpointsTab.Enabled = conflictsTab.Enabled = !running && repository != null;
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running || githubPane.Busy) { e.Cancel = true; return; }
            base.OnFormClosing(e);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { cacheLock?.Dispose(); components?.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
