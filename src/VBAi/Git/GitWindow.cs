using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Fenêtre de synchronisation Git du document VBA, avec revue, branches, points de contrôle et conflits.</summary>
    internal sealed partial class GitWindow : Form
    {
        /// <summary>Projet VBA dont les modules sont suivis par Git.</summary>
        private VbaGitProject project;
        /// <summary>Dépôt et branche associés au document.</summary>
        private MacroGitRepository repository;
        /// <summary>Dossier local de cache et de session pour la portée du document.</summary>
        private string cache;
        /// <summary>Verrou exclusif empêchant les opérations concurrentes sur le cache.</summary>
        private FileStream cacheLock;
        /// <summary>Indique qu’une opération Git est en cours dans la fenêtre.</summary>
        private bool running;
        /// <summary>Compte GitHub utilisé par les commandes de revue.</summary>
        private string account;
        /// <summary>Instantané vivant du projet VBA à comparer.</summary>
        private VbaGitSnapshot displayedLive;
        /// <summary>Instantané de référence de la branche synchronisée.</summary>
        private VbaGitSnapshot displayedBaseline;
        /// <summary>Branche affichée lors de la dernière comparaison.</summary>
        private string displayedBranch;
        /// <summary>Résout le cache local associé au document.</summary>
        internal static Func<string, string> CacheDirectory = MacroGitRepository.ResolveScopeDirectory;
        /// <summary>Initialise le dépôt et lit son état distant avec les commandes Git natives.</summary>
        internal Func<MacroGitRepository, string, Task> ConnectRepository = (selected, url) => Task.Run(() => { selected.Initialize(url); selected.Fetch(); });
        /// <summary>Crée la fenêtre Git et initialise la revue et les ressources visuelles.</summary>
        public GitWindow() { InitializeComponent(); BindViews(); Icon = VbeWindowIcons.Icon("github"); UiText.Apply(this, components); InitializeReview(); }

        /// <summary>Crée la fenêtre et l’associe au projet, à sa portée de cache et au compte GitHub choisi.</summary>
        /// <param name="project">Projet VBA à suivre.</param>
        /// <param name="scope">Clé de portée stable du document.</param>
        /// <param name="label">Libellé affiché du document.</param>
        /// <param name="account">Compte GitHub facultatif.</param>
        internal GitWindow(VbaGitProject project, string scope, string label, string account = null) : this()
        {
            this.account = account;
            this.project = project;
            githubPane.Configure(account, "", "main");
            cache = CacheDirectory(scope);
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

        /// <summary>Configuration locale de l’URL distante et de la branche suivie par un document.</summary>
        private sealed class Binding
        {
            /// <summary>Obtient ou définit l’URL distante validée.</summary>
            /// <value>URL HTTPS du dépôt.</value>
            public string Remote { get; set; }
            /// <summary>Obtient ou définit le nom de la branche locale associée.</summary>
            /// <value>Nom de branche suivi.</value>
            public string Branch { get; set; }
        }

        /// <summary>Valide la configuration distante, initialise le dépôt et compare le VBA courant au dernier état synchronisé.</summary>
        /// <param name="sender">Commande de connexion.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Connect_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                string url = MacroGitRepository.ValidateRemote(remote.Text);
                string bindingId;
                using (var hash = System.Security.Cryptography.SHA256.Create())
                    bindingId = BitConverter.ToString(hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(url + "\n" + branch.Text.Trim()))).Replace("-", "");
                var selected = new MacroGitRepository(Path.Combine(cache, bindingId + ".git"), branch.Text.Trim(), account);
                selected.Cancellation = operationCancellation.Token; selected.Progress = ReportProgress; cancelOperation.Enabled = true;
                try { await ConnectRepository(selected, url); }
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

        /// <summary>Actualise la comparaison du projet local avec le dépôt.</summary>
        /// <param name="sender">Commande de comparaison.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Compare_Click(object sender, EventArgs e) { await Perform(Compare, true); }
        /// <summary>Lit les instantanés VBA et Git et remplit les changements, l’historique, les branches et les conflits.</summary>
        /// <returns>Tâche terminée après l’actualisation de l’interface.</returns>
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

        /// <summary>Crée un commit local des éléments sélectionnés.</summary>
        /// <param name="sender">Commande de commit.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Commit_Click(object sender, EventArgs e) { await RunGitAction("commit_selected", text: commitMessage.Text); }
        /// <summary>Publie la branche locale vers le dépôt distant.</summary>
        /// <param name="sender">Commande Push.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Push_Click(object sender, EventArgs e) { await RunGitAction("push"); }
        /// <summary>Récupère les objets distants et actualise la comparaison.</summary>
        /// <param name="sender">Commande Fetch.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Fetch_Click(object sender, EventArgs e) { await RunGitAction("fetch"); }
        /// <summary>Exécute une opération Git sur les éléments cochés, puis relit le dépôt et le projet.</summary>
        /// <param name="action">Nom de l’opération métier à exécuter.</param>
        /// <param name="name">Nom facultatif de branche ou de point de contrôle.</param>
        /// <param name="text">Texte facultatif de commit ou de résolution.</param>
        /// <param name="choice">Choix facultatif de résolution de fusion.</param>
        /// <param name="path">Chemin facultatif du conflit ciblé.</param>
        /// <returns>Tâche terminée après l’opération et la comparaison de suivi.</returns>
        private async Task RunGitAction(string action, string name = null, string text = null, string choice = null, string path = null)
        {
            if (running || project == null) return;
            string[] selected = changes.CheckedItems.Cast<ModuleChange>().Where(x => x.Name != null).Select(x => x.Name).ToArray();
            bool references = changes.CheckedItems.Cast<ModuleChange>().Any(x => x.Name == null);
            GitModalSession.Request request = null;
            Func<Task<bool>> execute = async () => {
            try { await Perform(async () => {
                var operations = new MacroGitOperations(project, repository) { ImportPreview = ShowImportSummary };
                if (action == "commit_selected" && !repository.RecoveryPending && displayedLive != null && !project.Capture().SameAs(displayedLive))
                    throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                repository.Progress = ReportProgress;
                repository.Cancellation = action == "fetch" ? operationCancellation.Token : System.Threading.CancellationToken.None;
                cancelOperation.Enabled = action == "fetch";
                string expectedState = null;
                if (GitModalSession.RequiresHandoff(action))
                {
                    if (Modal && modalSession == null) throw new InvalidOperationException("A modal import requires its owning Git session.");
                    if (modalSession != null)
                    {
                        var live = project.Capture();
                        expectedState = await Task.Run(() => operations.Revision(live));
                        request = new GitModalSession.Request(action, name, text, choice, path, selected, references, expectedState, async () => {
                            var observed = project.Capture();
                            if (expectedState != await Task.Run(() => operations.Revision(observed)))
                                throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                        });
                        await AdmitImport(request);
                        operations.ImportOwnerPreflight = () => modalSession.RequireImportOwner(request);
                    }
                }
                object result = await operations.ExecuteAsync(action, expectedState, name: name, text: text, choice: choice, path: path,
                    modules: request == null ? selected : request.Modules, references: references);
                repository.Cancellation = System.Threading.CancellationToken.None;
                await Compare();
                status.Text = (action == "commit" || action == "commit_selected") ? UiText.Get("Local commit created. Use Push to publish it.") :
                    action == "push" ? UiText.Get("Push complete.") : action == "pull" ? UiText.Get("Pull and import complete. Check and save the document.") :
                    action == "rollback" || action == "checkpoint_restore" ? UiText.Get("VBA restored. Check and save the document.") :
                    repository.PendingMerge != null ? UiText.Get("Merge prepared: resolve conflicts, then click Complete merge.") : UiText.Get("Operation complete: ") + action;
                if (repository.PendingMerge != null) tabs.SelectedTab = conflictsTab;
            }); }
            catch (Exception error)
            {
                operationFailure = operationFailure == null || ReferenceEquals(operationFailure, error) ? error :
                    new AggregateException("The Git operation and its cleanup both failed.", operationFailure, error);
                status.Text = operationFailure.Message + " [" + GitFailureDiagnostic.Describe(operationFailure) + "] · " +
                    UiText.Get("Check the connection, account and Git state, then retry.");
            }
            return true;
            };
            try
            {
                if (modalSession != null) await VbeUiTask.Run(execute);
                else await execute();
            }
            catch (Exception error)
            {
                operationFailure = operationFailure == null ? error : new AggregateException(operationFailure, error);
                throw;
            }
            finally
            {
                if (request != null && (request.Phase == "Executing" || request.Phase == "Refused"))
                    request.Complete(operationFailure, PublishHandoffState);
            }
        }
        /// <summary>Affiche le diff correspondant au changement sélectionné.</summary>
        /// <param name="sender">Liste des changements.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Changes_SelectedIndexChanged(object sender, EventArgs e)
        {
            var item = changes.SelectedItem as ModuleChange;
            if (item == null) { diff.ShowDiff("", ""); return; }
            diff.ShowDiff(Source(displayedBaseline, item.Name), Source(displayedLive, item.Name));
        }

        /// <summary>Récupère les changements distants et importe les modules avec sauvegarde préalable.</summary>
        /// <param name="sender">Commande Pull.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Pull_Click(object sender, EventArgs e) { await RunGitAction("pull"); }
        /// <summary>Restaure le dernier état VBA synchronisé.</summary>
        /// <param name="sender">Commande de restauration.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Restore_Click(object sender, EventArgs e) { await RunGitAction("rollback"); }
        /// <summary>Crée un point de contrôle nommé.</summary>
        /// <param name="sender">Commande de création.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void CheckpointCreate_Click(object sender, EventArgs e) { await RunGitAction("checkpoint_create", checkpointName.Text); }
        /// <summary>Restaure le point de contrôle sélectionné.</summary>
        /// <param name="sender">Commande de restauration.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void CheckpointRestore_Click(object sender, EventArgs e)
        {
            var selected = checkpointList.SelectedItem as GitCheckpoint;
            if (selected != null) await RunGitAction("checkpoint_restore", selected.Id);
        }
        /// <summary>Crée une branche locale avec le nom saisi.</summary>
        /// <param name="sender">Commande de création de branche.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void BranchCreate_Click(object sender, EventArgs e) { await RunGitAction("branch_create", branchName.Text); }
        /// <summary>Configure le suivi de la branche distante indiquée.</summary>
        /// <param name="sender">Commande de suivi.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void BranchTrack_Click(object sender, EventArgs e) { await RunGitAction("branch_track", branchName.Text); }
        /// <summary>Charge les branches distantes et les propose dans le sélecteur de branche.</summary>
        /// <param name="sender">Commande de liste des branches distantes.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void RemoteBranches_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                string[] names = await Task.Run(() => repository.RemoteBranches());
                branchName.Items.Clear(); branchName.Items.AddRange(names);
                status.Text = names.Length + UiText.Get(" remote branch(es). Choose a name, then Track remote.");
            });
            if (!running && branchName.Items.Count > 0) branchName.DroppedDown = true;
        }
        /// <summary>Bascule vers la branche locale sélectionnée.</summary>
        /// <param name="sender">Commande de bascule.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void BranchSwitch_Click(object sender, EventArgs e)
        { if (branchList.SelectedItem != null) await RunGitAction("branch_switch", branchList.SelectedItem.ToString()); }
        /// <summary>Prépare une fusion avec la branche sélectionnée.</summary>
        /// <param name="sender">Commande de démarrage de fusion.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeBegin_Click(object sender, EventArgs e)
        { if (branchList.SelectedItem != null) await RunGitAction("merge_begin", branchList.SelectedItem.ToString()); }
        /// <summary>Résout le conflit sélectionné en conservant la version locale.</summary>
        /// <param name="sender">Commande de résolution locale.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeOurs_Click(object sender, EventArgs e) { await ResolveMerge("ours"); }
        /// <summary>Résout le conflit sélectionné en conservant la version entrante.</summary>
        /// <param name="sender">Commande de résolution entrante.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeTheirs_Click(object sender, EventArgs e) { await ResolveMerge("theirs"); }
        /// <summary>Résout le conflit sélectionné avec le texte fourni par l’utilisateur.</summary>
        /// <param name="sender">Commande de résolution textuelle.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeText_Click(object sender, EventArgs e) { await ResolveMerge("text"); }
        /// <summary>Applique le choix de résolution au conflit sélectionné.</summary>
        /// <param name="choice">Version à conserver ou mode textuel.</param>
        /// <returns>Tâche terminée après la mise à jour du dépôt.</returns>
        private async Task ResolveMerge(string choice)
        { if (conflictList.SelectedItem != null) await RunGitAction("merge_resolve", text: resolutionText.Text, choice: choice, path: conflictList.SelectedItem.ToString()); }
        /// <summary>Finalise la fusion avec le message de commit courant.</summary>
        /// <param name="sender">Commande de finalisation.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeComplete_Click(object sender, EventArgs e) { await RunGitAction("merge_complete", text: commitMessage.Text); }
        /// <summary>Abandonne la fusion en cours.</summary>
        /// <param name="sender">Commande d’abandon.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void MergeAbort_Click(object sender, EventArgs e) { await RunGitAction("merge_abort"); }
        /// <summary>Charge les contenus de conflit du fichier sélectionné et peuple la comparaison à deux côtés.</summary>
        /// <param name="sender">Liste des conflits.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Exécute une transaction exclusive, gère son annulation, son statut et l’état des commandes.</summary>
        /// <param name="action">Opération asynchrone à exécuter.</param>
        /// <param name="refresh">Autorise le changement de branche observé lors d’une actualisation.</param>
        /// <returns>Tâche achevée après libération des verrous et ressources d’opération.</returns>
        private async Task Perform(Func<Task> action, bool refresh = false)
        {
            if (running || project == null) return;
            running = true; operationCancellation = new System.Threading.CancellationTokenSource(); operationProgress.Visible = true; cancelOperation.Visible = true; UpdateButtons(); status.Text = UiText.Get("Operation in progress…");
            operationFailure = null;
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
            catch (OperationCanceledException error) { operationFailure = error; status.Text = UiText.Get("Operation cancelled."); }
            catch (Exception ex) { operationFailure = ex; status.Text = ex.Message + " [" + GitFailureDiagnostic.Describe(ex) + "] · " + UiText.Get("Check the connection, account and Git state, then retry."); }
            finally { FinishOperation(); }
        }
        /// <summary>Recalcule l’activation des commandes selon l’opération, le dépôt et l’état de revue.</summary>
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
        /// <summary>Empêche la fermeture pendant une opération Git ou une opération GitHub.</summary>
        /// <param name="e">Annulation et données de l’événement de fermeture.</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if ((running && !(leavingForImport && modalRequest != null && modalRequest.Phase == "AwaitingModalReturn" &&
                DialogResult == DialogResult.OK && e.CloseReason == CloseReason.None)) || githubPane.Busy) { e.Cancel = true; return; }
            base.OnFormClosing(e);
        }
        /// <summary>Libère le verrou de cache et les composants du formulaire.</summary>
        /// <param name="disposing"><see langword="true"/> si l’appel provient de <see cref="IDisposable.Dispose()"/>.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) { cacheLock?.Dispose(); components?.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
