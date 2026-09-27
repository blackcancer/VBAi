using System;
using System.IO;
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
        public GitWindow() { InitializeComponent(); }

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
                    new Binding { Remote = url, Branch = selected.Branch }));
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
            syncStatus.Text = repository.Branch + "   ·   " + await Task.Run(() => repository.SynchronizationStatus());
            status.Text = repository.RecoveryPending ? "Import interrompu : restaurez le VBA avant de poursuivre. La sauvegarde est conservée dans le cache." :
                baseline == null ? "Première liaison : commit puis push pour publier, ou pull pour importer le dépôt avec sauvegarde préalable." :
                changes.Items.Count == 0 ? "Le VBA correspond au dernier état synchronisé." : changes.Items.Count + " fichier(s) modifié(s) depuis la dernière synchronisation.";
        }

        private async void Commit_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                RequireReady();
                var live = project.Capture();
                string message = commitMessage.Text.Trim();
                if (message.Length == 0) throw new InvalidOperationException("Saisissez un message de commit.");
                string commit = await Task.Run(() => {
                    string local = repository.Resolve(repository.Head);
                    string incoming = repository.Resolve("refs/remotes/origin/selected");
                    var remoteState = repository.Read(incoming);
                    if (local == null && remoteState != null && !live.SameAs(remoteState))
                        throw new InvalidOperationException("Le dépôt contient déjà un autre état VBA. Faites d’abord un pull avec sauvegarde locale.");
                    string parent = local ?? incoming;
                    var parentState = repository.Read(parent);
                    string next = live.SameAs(parentState) ? parent : repository.Commit(live, parent, message);
                    repository.SetRef(repository.Head, next);
                    return next;
                });
                // The host may have changed while Git I/O was pending. The baseline is the captured state, never a fresh capture.
                await Task.Run(() => repository.SetRef(MacroGitRepository.Baseline, commit));
                await Compare();
                status.Text = "Commit local créé · " + commit.Substring(0, 8) + ". Utilisez Push pour le publier sur GitHub.";
            });
        }

        private async void Push_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                RequireReady();
                await Task.Run(() => {
                    string local = repository.Resolve(repository.Head) ?? throw new InvalidOperationException("Créez d’abord un commit local.");
                    string incoming = repository.Fetch();
                    if (incoming != null) repository.RequireFastForward(incoming, local);
                    repository.Push(local);
                });
                await Compare();
                status.Text = "Push terminé. Les commits locaux sont publiés ; les modifications non commitées restent dans le VBA.";
            });
        }

        private async void Fetch_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                await Task.Run(() => repository.Fetch());
                await Compare();
                status.Text = "Fetch terminé. Aucun code VBA n’a été modifié.";
            });
        }

        private void Changes_SelectedIndexChanged(object sender, EventArgs e)
        {
            diff.Rows.Clear();
            if (changes.SelectedItem == null || displayedLive == null) return;
            string name = changes.SelectedItem.ToString().Substring(2);
            if (name.EndsWith(".frx", StringComparison.Ordinal)) { diff.Rows.Add("Ressource binaire", "Ressource binaire"); return; }
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
            if (x < left.Length || y < right.Length) diff.Rows.Add("Aperçu limité à 2 000 lignes", "Aperçu limité à 2 000 lignes");
        }

        private void AddDiffRow(string before, string after, int oldLine, int newLine, bool changed)
        {
            int row = diff.Rows.Add(before == null ? "" : (oldLine + 1) + "  " + before, after == null ? "" : (newLine + 1) + "  " + after);
            if (!changed) return;
            if (before != null) diff.Rows[row].Cells[0].Style.BackColor = System.Drawing.Color.MistyRose;
            if (after != null) diff.Rows[row].Cells[1].Style.BackColor = System.Drawing.Color.Honeydew;
        }

        private async void Pull_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                RequireReady();
                var live = project.Capture();
                string incoming = null;
                var target = await Task.Run(() => {
                    var baseline = repository.Read(repository.Resolve(MacroGitRepository.Baseline));
                    if (baseline != null && !live.SameAs(baseline))
                        throw new InvalidOperationException("Le VBA contient des modifications locales. Commitez-les avant de lancer le pull.");
                    incoming = repository.Fetch();
                    if (incoming == null) throw new InvalidOperationException("La branche distante n’existe pas encore.");
                    repository.RequireFastForward(repository.Resolve(repository.Head), incoming);
                    return repository.Read(incoming) ?? throw new InvalidOperationException("La branche ne contient pas de dossier vba avec manifeste CodexVBA.");
                });
                if (!live.SameAs(project.Capture())) throw new InvalidOperationException("Le VBA a changé pendant le téléchargement. Relancez la comparaison.");
                if (live.SameAs(target))
                {
                    await Task.Run(() => { repository.SetRef(repository.Head, incoming); repository.SetRef(MacroGitRepository.Baseline, incoming); });
                    await Compare(); status.Text = "Déjà à jour. Aucun import nécessaire."; return;
                }
                await Task.Run(() => repository.PrepareRecovery(live));
                bool mutationStarted = false;
                try { project.Apply(target, live, () => mutationStarted = true); }
                finally
                {
                    if (mutationStarted)
                    {
                        var actual = project.Capture();
                        await Task.Run(() => repository.RecordImportedState(actual));
                    }
                    else await Task.Run(() => repository.CompleteRecovery());
                }
                await Task.Run(() => {
                    repository.SetRef(repository.Head, incoming);
                    repository.SetRef(MacroGitRepository.Baseline, incoming);
                    repository.CompleteRecovery();
                });
                await Compare();
                status.Text = "Pull et import terminés. Vérifiez/compilez le VBA, puis enregistrez le document dans son application. Restaurer reste disponible.";
            });
        }

        private async void Restore_Click(object sender, EventArgs e)
        {
            await Perform(async () => {
                var live = project.Capture();
                string backup = null;
                var target = await Task.Run(() => {
                    var after = repository.Read(repository.Resolve(MacroGitRepository.AfterImport));
                    if (after == null || !live.SameAs(after))
                        throw new InvalidOperationException("L’état actuel diffère de la relecture après import, ou cette relecture manque. La sauvegarde reste dans le cache ; restauration automatique bloquée pour préserver les modifications suivantes.");
                    backup = repository.Resolve(MacroGitRepository.Backup);
                    return repository.Read(backup) ?? throw new InvalidOperationException("Aucune sauvegarde avant import.");
                });
                bool wasPending = repository.RecoveryPending;
                File.WriteAllText(repository.RecoveryFile, backup);
                bool mutationStarted = false;
                try { project.Apply(target, live, () => mutationStarted = true); }
                finally
                {
                    if (mutationStarted)
                    {
                        var actual = project.Capture();
                        await Task.Run(() => repository.RecordImportedState(actual));
                    }
                    else if (!wasPending) await Task.Run(() => repository.CompleteRecovery());
                }
                await Task.Run(() => repository.CompleteRecovery());
                await Compare();
                status.Text = "VBA restauré. Enregistrez le document dans son application. GitHub conserve son historique ; un prochain push publiera la restauration.";
            });
        }

        private void RequireReady()
        {
            if (repository.RecoveryPending) throw new InvalidOperationException("Un import a été interrompu. Restaurez le VBA avant une nouvelle synchronisation.");
        }
        private async Task Perform(Func<Task> action)
        {
            if (running || project == null) return;
            running = true; UpdateButtons(); status.Text = "Opération en cours…";
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
