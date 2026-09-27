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
        public GitWindow() { InitializeComponent(); Icon = VbeWindowIcons.Icon("github"); UiText.Apply(this, components); }

        internal GitWindow(VbaGitProject project, string scope, string label, string account = null) : this()
        {
            this.account = account;
            this.project = project;
            cache = MacroGitRepository.ScopeDirectory(scope);
            documentLabel.Text = label;
            Directory.CreateDirectory(cache);
            cacheLock = new FileStream(Path.Combine(cache, "session.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            try
            {
                string file = Path.Combine(cache, "binding.json");
                if (File.Exists(file))
                {
                    var binding = new JavaScriptSerializer().Deserialize<Binding>(File.ReadAllText(file));
                    remote.Text = binding.Remote; branch.Text = binding.Branch;
                }
            }
            catch { cacheLock.Dispose(); throw; }
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
                await Task.Run(() => { selected.Initialize(url); selected.Fetch(); });
                repository = selected;
                string bindingFile = Path.Combine(cache, "binding.json");
                string temporaryBinding = Path.Combine(cache, "binding.pending");
                File.WriteAllText(temporaryBinding, new JavaScriptSerializer().Serialize(
                    // Keep the original cache key stable when Initialize restores another active branch.
                    new Binding { Remote = url, Branch = branch.Text.Trim() }));
                if (File.Exists(bindingFile)) File.Replace(temporaryBinding, bindingFile, null);
                else File.Move(temporaryBinding, bindingFile);
                await Compare();
            });
        }

        private async void Compare_Click(object sender, EventArgs e) { await Perform(Compare); }
        private async Task Compare()
        {
            var live = project.Capture();
            var baseline = await Task.Run(() => repository.Read(repository.Resolve(MacroGitRepository.Baseline)));
            displayedLive = live; displayedBaseline = baseline;
            changes.Items.Clear(); changes.Items.AddRange(live.Changes(baseline));
            history.Items.Clear(); history.Items.AddRange(await Task.Run(() => repository.History()));
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

        private async void Commit_Click(object sender, EventArgs e) { await RunGitAction("commit", text: commitMessage.Text); }
        private async void Push_Click(object sender, EventArgs e) { await RunGitAction("push"); }
        private async void Fetch_Click(object sender, EventArgs e) { await RunGitAction("fetch"); }
        private async Task RunGitAction(string action, string name = null, string text = null, string choice = null, string path = null)
        {
            await Perform(async () => {
                var operations = new MacroGitOperations(project, repository);
                object result = await operations.ExecuteAsync(action, name: name, text: text, choice: choice, path: path);
                await Compare();
                status.Text = action == "commit" ? UiText.Get("Local commit created. Use Push to publish it.") :
                    action == "push" ? UiText.Get("Push complete.") : action == "pull" ? UiText.Get("Pull and import complete. Check and save the document.") :
                    action == "rollback" || action == "checkpoint_restore" ? UiText.Get("VBA restored. Check and save the document.") :
                    repository.PendingMerge != null ? UiText.Get("Merge prepared: resolve conflicts, then click Complete merge.") : UiText.Get("Operation complete: ") + action;
                if (repository.PendingMerge != null) tabs.SelectedTab = conflictsTab;
            });
        }
        private void Changes_SelectedIndexChanged(object sender, EventArgs e)
        {
            diff.Rows.Clear();
            if (changes.SelectedItem == null || displayedLive == null) return;
            string name = changes.SelectedItem.ToString().Substring(2);
            if (name.EndsWith(".frx", StringComparison.Ordinal)) { diff.Rows.Add(UiText.Get("Binary resource"), UiText.Get("Binary resource")); return; }
            byte[] old = null, current = null;
            displayedBaseline?.Serialize().TryGetValue(name, out old);
            displayedLive.Serialize().TryGetValue(name, out current);
            string before = old == null ? "" : VbaGitSnapshot.Utf8.GetString(old);
            string after = current == null ? "" : VbaGitSnapshot.Utf8.GetString(current);
            string[] left = CodeRollback.Lines(before), right = CodeRollback.Lines(after);
            int x = 0, y = 0;
            foreach (var hunk in CodeRollback.Hunks(before, after))
            {
                while (x < hunk.BeforeStart && y < hunk.AfterStart && diff.Rows.Count < 2000)
                    AddDiffRow(left[x], right[y], x++, y++, false);
                for (int i = 0; i < Math.Max(hunk.Before.Length, hunk.After.Length) && diff.Rows.Count < 2000; i++)
                {
                    bool a = i < hunk.Before.Length, b = i < hunk.After.Length;
                    AddDiffRow(a ? hunk.Before[i] : null, b ? hunk.After[i] : null, x, y, true);
                    if (a) x++; if (b) y++;
                }
                if (diff.Rows.Count >= 2000) break;
            }
            while (x < left.Length && y < right.Length && diff.Rows.Count < 2000) AddDiffRow(left[x], right[y], x++, y++, false);
            if (x < left.Length || y < right.Length) diff.Rows.Add(UiText.Get("Preview limited to 2,000 lines"), UiText.Get("Preview limited to 2,000 lines"));
        }

        private void AddDiffRow(string before, string after, int oldLine, int newLine, bool changed)
        {
            int row = diff.Rows.Add(before == null ? "" : (oldLine + 1) + "  " + before, after == null ? "" : (newLine + 1) + "  " + after);
            if (!changed) return;
            if (before != null) diff.Rows[row].Cells[0].Style.BackColor = System.Drawing.Color.MistyRose;
            if (after != null) diff.Rows[row].Cells[1].Style.BackColor = System.Drawing.Color.Honeydew;
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
                resolutionText.Text = content.Ours.StartsWith("[", StringComparison.Ordinal) ? "" : content.Ours;
                foreach (var row in CodeChange.BuildRows(content.Ours, content.Theirs).Take(2000))
                {
                    int index = conflictDiff.Rows.Add(row.Kind == CodeDiffKind.Added ? "" : row.Text, row.Kind == CodeDiffKind.Removed ? "" : row.Text);
                    if (row.Kind == CodeDiffKind.Added) conflictDiff.Rows[index].Cells[1].Style.BackColor = System.Drawing.Color.Honeydew;
                    if (row.Kind == CodeDiffKind.Removed) conflictDiff.Rows[index].Cells[0].Style.BackColor = System.Drawing.Color.MistyRose;
                }
                status.Text = UiText.Get("Conflict: ") + path + UiText.Get(". Choose a version or edit the full content below.");
            });
        }
        private async Task Perform(Func<Task> action)
        {
            if (running || project == null) return;
            running = true; UpdateButtons(); status.Text = UiText.Get("Operation in progress…");
            try { await action(); }
            catch (Exception ex) { status.Text = ex.Message; }
            finally { running = false; UpdateButtons(); }
        }
        private void UpdateButtons()
        {
            connect.Enabled = !running && repository == null;
            remote.ReadOnly = branch.ReadOnly = repository != null || running;
            compare.Enabled = commit.Enabled = fetch.Enabled = push.Enabled = pull.Enabled = restore.Enabled = !running && repository != null;
            commitMessage.Enabled = !running;
            branchesTab.Enabled = checkpointsTab.Enabled = conflictsTab.Enabled = !running && repository != null;
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running) { e.Cancel = true; return; }
            base.OnFormClosing(e);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { cacheLock?.Dispose(); components?.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
