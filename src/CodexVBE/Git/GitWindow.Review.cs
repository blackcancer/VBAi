using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Affiche l’historique, les changements VBA et les prévisualisations d’import Git.</summary>
    internal sealed partial class GitWindow
    {
        /// <summary>Source d’annulation de l’opération de revue ou d’import courante.</summary>
        private CancellationTokenSource operationCancellation;
        /// <summary>Commit ou point de contrôle actuellement affiché pour revue.</summary>
        private string reviewCommit;
    /// <summary>Module et libellé de changement présentés dans la liste de revue.</summary>
        private sealed class ModuleChange
        {
            /// <summary>Nom du module concerné, ou null pour un changement global.</summary>
            internal string Name;
            /// <summary>Préfixe et nom affichés pour le changement.</summary>
            internal string Label;
            /// <summary>Retourne le libellé affiché dans la liste.</summary>
            /// <returns>Libellé de la ligne de revue.</returns>
            public override string ToString() { return Label; }
        }
        /// <summary>Relie les événements de navigation de revue au layout et aux services Git.</summary>
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
        /// <summary>Affiche les commandes et espaces adaptés à l’onglet actif et à l’état de liaison.</summary>
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
        /// <summary>Retourne le code d’un composant dans le snapshot, ou l’empreinte des références.</summary>
        /// <param name="snapshot">Snapshot contenant le composant à lire.</param>
        /// <param name="name">Nom du module, ou null pour lire les références.</param>
        /// <returns>Code source normalisé, références sérialisées, ou chaîne vide si absent.</returns>
        private static string Source(VbaGitSnapshot snapshot, string name)
        {
            if (snapshot == null) return "";
            if (name == null) return snapshot.Manifest.References;
            var module = snapshot.Manifest.Components.FirstOrDefault(x => x.Name == name);
            return module == null ? "" : VbaGitSnapshot.Utf8.GetString(snapshot.Files[module.FileName]);
        }
        /// <summary>Construit la liste des modules et références différentes entre deux snapshots.</summary>
        /// <param name="target">Snapshot courant à présenter.</param>
        /// <param name="baseline">Snapshot précédent utilisé pour comparer les fichiers.</param>
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
        /// <summary>Affiche le résumé de prévisualisation dans l’onglet d’import.</summary>
        /// <param name="text">Résumé d’import à afficher.</param>
        private void ShowImportSummary(string text)
        {
            importSummary.Text = text; tabs.SelectedTab = importTab;
        }
        /// <summary>Transmet à l’interface l’avancement Git si la fenêtre reste disponible.</summary>
        /// <param name="text">Résumé d’import à afficher.</param>
        private void ReportProgress(string text)
        {
            if (IsHandleCreated && !IsDisposed) BeginInvoke(new Action(() => { if (running) status.Text = text; }));
        }
        /// <summary>Annule l’opération Git active.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void CancelOperation_Click(object sender, EventArgs e) { operationCancellation?.Cancel(); }
        /// <summary>Récupère et compare l’état distant sans importer ses changements.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
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
        /// <summary>Ouvre dans le VBE le module sélectionné dans la revue.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void OpenModule_Click(object sender, EventArgs e)
        {
            try { var item = changes.SelectedItem as ModuleChange; if (item?.Name != null) project?.OpenModule(item.Name); }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        /// <summary>Restaure le module sélectionné depuis le commit de revue.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void RestoreModule_Click(object sender, EventArgs e)
        {
            var item = changes.SelectedItem as ModuleChange;
            if (item?.Name == null || reviewCommit == null || repository == null) { status.Text = UiText.Get("Select a revision in History or Checkpoints, then a module to restore."); return; }
            await RunGitAction("module_restore", name: reviewCommit, path: item.Name);
        }
        /// <summary>Charge un commit et son parent ou la seconde sélection pour comparaison.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
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
        /// <summary>Affiche les différences entre le point de contrôle et le projet vivant.</summary>
        /// <param name="sender">Contrôle à l’origine de l’action.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
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
