using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class GitWindow
    {
        private CancellationTokenSource operationCancellation;
        private string reviewCommit;
        private sealed class ModuleChange
        {
            internal string Name, Label;
            public override string ToString() { return Label; }
        }
        private void InitializeReview()
        {
            tabs.SelectedIndexChanged += (s, e) => AdjustReviewLayout();
            Resize += (s, e) => AdjustReviewLayout();
            AdjustReviewLayout();
            githubPane.RepositorySelected = (url, selectedBranch) => {
                if (repository != null) { status.Text = UiText.Get("This document is already linked. Reopen GitHub to choose another repository."); return; }
                remote.Text = url; branch.Text = selectedBranch; tabs.SelectedTab = changesTab;
                status.Text = UiText.Get("Repository selected. Click Link repository to connect this document.");
            };
            githubPane.LoadDraft = () => repository?.PullDraft();
            githubPane.OpenModule = (name, line) => project?.OpenModule(name, line);
        }
        private void AdjustReviewLayout()
        {
            bool online = tabs.SelectedTab == githubTab;
            float scale = DeviceDpi / 96F;
            bool binding = !online && repository == null;
            remote.Visible = remoteLabel.Visible = branch.Visible = branchLabel.Visible = connect.Visible = binding;
            help.Visible = binding && ClientSize.Height >= 650 * scale;
            layout.RowStyles[1].SizeType = help.Visible ? SizeType.AutoSize : SizeType.Absolute; layout.RowStyles[1].Height = 0;
            layout.RowStyles[2].Height = layout.RowStyles[3].Height = binding ? 34 * scale : 0;
            layout.RowStyles[4].Height = binding ? 36 * scale : 0;
            bool editing = !online && repository != null && tabs.SelectedTab == changesTab;
            commitMessage.Visible = messageLabel.Visible = editing; layout.RowStyles[5].Height = editing ? 54 * scale : 0;
            syncStatus.Visible = actions.Visible = !online;
            layout.RowStyles[6].Height = online ? 0 : 28 * scale;
            layout.RowStyles[7].SizeType = online ? SizeType.Absolute : SizeType.AutoSize; layout.RowStyles[7].Height = 0;
            layout.RowStyles[9].Height = 64 * scale;
        }
        private static string Source(VbaGitSnapshot snapshot, string name)
        {
            if (snapshot == null) return "";
            if (name == null) return snapshot.Manifest.References;
            var module = snapshot.Manifest.Components.FirstOrDefault(x => x.Name == name);
            return module == null ? "" : VbaGitSnapshot.Utf8.GetString(snapshot.Files[module.FileName]);
        }
        private void PopulateChanges(VbaGitSnapshot target, VbaGitSnapshot baseline)
        {
            displayedLive = target; displayedBaseline = baseline;
            changes.Items.Clear(); diff.ShowDiff("", "");
            if (target == null) return;
            var changed = target.Changes(baseline).Select(x => x.Substring(2)).ToArray();
            foreach (string name in target.Manifest.Components.Concat(baseline?.Manifest.Components ?? new VbaGitComponent[0]).Select(x => x.Name).Distinct().OrderBy(x => x))
            {
                var old = baseline?.Manifest.Components.FirstOrDefault(x => x.Name == name);
                var next = target.Manifest.Components.FirstOrDefault(x => x.Name == name);
                if (reviewCommit == null && !changed.Any(x => x == old?.FileName || x == next?.FileName || x == name + ".frx")) continue;
                changes.Items.Add(new ModuleChange { Name = name, Label = (old == null ? "+ " : next == null ? "− " : "~ ") + name +
                    ((old?.HasResources ?? false) || (next?.HasResources ?? false) ? " · " + UiText.Get("Form resources included") : "") }, reviewCommit == null);
            }
            if (baseline == null || target.Manifest.References != baseline.Manifest.References)
                changes.Items.Add(new ModuleChange { Label = UiText.Get("VBA references") }, reviewCommit == null);
            if (changes.Items.Count > 0) changes.SelectedIndex = 0;
        }
        private void ShowImportSummary(string text)
        {
            importSummary.Text = text; tabs.SelectedTab = importTab;
        }
        private void ReportProgress(string text)
        {
            if (IsHandleCreated && !IsDisposed) BeginInvoke(new Action(() => { if (running) status.Text = text; }));
        }
        private void CancelOperation_Click(object sender, EventArgs e) { operationCancellation?.Cancel(); }
        private async void PreviewImport_Click(object sender, EventArgs e)
        {
            if (repository == null) return;
            await Perform(async () => {
                cancelOperation.Enabled = true; repository.Cancellation = operationCancellation.Token; repository.Progress = ReportProgress;
                var target = await Task.Run(() => repository.Read(repository.Fetch()));
                if (target == null) throw new InvalidOperationException(UiText.Get("The target contains no VBA sources."));
                ShowImportSummary(target.ImportSummary(project.Capture()));
                status.Text = UiText.Get("Preview only. Pull imports these changes with a checkpoint.");
            });
        }
        private void OpenModule_Click(object sender, EventArgs e)
        {
            try { var item = changes.SelectedItem as ModuleChange; if (item?.Name != null) project?.OpenModule(item.Name); }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        private async void RestoreModule_Click(object sender, EventArgs e)
        {
            var item = changes.SelectedItem as ModuleChange;
            if (item?.Name == null || reviewCommit == null || repository == null) { status.Text = UiText.Get("Select a revision in History or Checkpoints, then a module to restore."); return; }
            await RunGitAction("module_restore", name: reviewCommit, path: item.Name);
        }
        private async void HistoryChanged(object sender, EventArgs e)
        {
            if (running || repository == null) return;
            var selected = history.SelectedItems.Cast<GitCommitInfo>().ToArray();
            if (selected.Length == 0) return;
            await Perform(async () => {
                var current = selected[0]; reviewCommit = current.Id;
                historyDetails.Text = await Task.Run(() => repository.CommitDetails(current.Id));
                string previous = selected.Length > 1 ? selected[1].Id : await Task.Run(() => repository.ParentCommit(current.Id));
                var target = await Task.Run(() => repository.Read(current.Id));
                var baseline = await Task.Run(() => repository.Read(previous));
                PopulateChanges(target, baseline); if (sender == historyCompare) tabs.SelectedTab = changesTab;
                status.Text = UiText.Get("Reviewing revision") + " " + current.Id.Substring(0, 8) + " · " + UiText.Get("Compare returns to live VBA.");
            });
        }
        private async void CheckpointChanged(object sender, EventArgs e)
        {
            if (running || repository == null) return;
            var checkpoint = checkpointList.SelectedItem as GitCheckpoint; if (checkpoint == null) return;
            await Perform(async () => {
                reviewCommit = checkpoint.Commit;
                PopulateChanges(await Task.Run(() => repository.Read(checkpoint.Commit)), project.Capture());
                tabs.SelectedTab = changesTab; status.Text = UiText.Get("Reviewing checkpoint") + " · " + checkpoint.Label;
            });
        }
    }
}
