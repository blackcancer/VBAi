using System;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
        /// <summary>Panneau WinForms de sélection de dépôt et de gestion des demandes de fusion GitHub.</summary>
    public sealed partial class GitHubPane : UserControl
    {
        /// <summary>Crée un client API authentifié pour le compte demandé.</summary>
        internal Func<string, GitHubApi> ApiFactory = CreateApi;
        /// <summary>Ouvre un lien externe après validation de son adresse.</summary>
        internal Action<string> OpenExternalLink = SafeLinks.Open;
        /// <summary>Construit un client API GitHub associé à un compte.</summary>
        /// <param name="account">Identifiant du compte d’authentification.</param>
        /// <returns>Client API GitHub configuré pour ce compte.</returns>
        private static GitHubApi CreateApi(string account) { return new GitHubApi(account); }
        /// <summary>Compte GitHub actif, dépôt distant sélectionné et branche source locale.</summary>
        private string account, remote, branch;
        /// <summary>Source d’annulation de l’opération GitHub actuellement exécutée.</summary>
        private CancellationTokenSource cancellation;
        /// <summary>Dépôts chargés pour filtrage et sélection dans l’interface.</summary>
        private GitHubRepositoryInfo[] repositories = new GitHubRepositoryInfo[0];
        /// <summary>Notifie le conteneur du dépôt distant et de la branche choisis.</summary>
        internal Action<string, string> RepositorySelected;
        /// <summary>Ouvre un module VBA à la ligne demandée depuis un fichier ou commentaire.</summary>
        internal Action<string, int> OpenModule;
        /// <summary>Fournit le brouillon de demande de fusion préparé pour la branche courante.</summary>
        internal Func<GitPullDraft> LoadDraft;
        /// <summary>Demande de fusion sélectionnée pour les actions de détail et d’ouverture.</summary>
        private GitHubPull selectedPull;
        /// <summary>Indique si une opération asynchrone GitHub est en cours.</summary>
        /// <value><see langword="true"/> lorsqu’une source d’annulation est active.</value>
        internal bool Busy { get { return cancellation != null; } }
        /// <summary>Crée le panneau et applique les textes localisés.</summary>
        public GitHubPane() { InitializeComponent(); BindViews(); UiText.Apply(this, components); }
        /// <summary>Configure le compte, le dépôt distant et la branche locale affichés.</summary>
        /// <param name="selectedAccount">Identifiant du compte GitHub à utiliser.</param>
        /// <param name="url">URL du dépôt distant sélectionné.</param>
        /// <param name="activeBranch">Branche locale servant de source aux demandes de fusion.</param>
        internal void Configure(string selectedAccount, string url, string activeBranch)
        { account = selectedAccount; remote = url; branch = activeBranch; sourceLabel.Text = UiText.Get("Source branch") + ": " + branch; }
        /// <summary>Exécute une opération avec client GitHub, état d’attente, erreurs et annulation.</summary>
        /// <param name="action">Opération asynchrone à exécuter avec le client GitHub.</param>
        /// <param name="cancelable">Indique si le bouton d’annulation doit être activé pendant l’opération.</param>
        /// <returns>Une tâche qui se termine après l’opération et la mise à jour de l’état visuel.</returns>
        private async Task Run(Func<GitHubApi, CancellationToken, Task> action, bool cancelable = true)
        {
            if (cancellation != null) return;
            using (var pending = new CancellationTokenSource())
            using (var api = ApiFactory(account))
            {
                cancellation = pending; pages.Enabled = false; cancel.Enabled = cancelable; status.Text = UiText.Get("Connecting to GitHub…");
                try { await action(api, pending.Token); status.Text = UiText.Get("Operation complete."); }
                catch (OperationCanceledException) { status.Text = UiText.Get("Operation cancelled."); }
                catch (Exception ex) { status.Text = ex is HttpRequestException ? UiText.Get("Unable to reach GitHub. Check your connection, then retry.") : ex.Message; }
                finally { cancellation = null; if (!IsDisposed) { pages.Enabled = true; cancel.Enabled = false; } }
            }
        }
        /// <summary>Charge les dépôts et organisations du compte puis met à jour les listes.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void LoadRepositories_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                repositories = await api.Repositories(ct); FilterRepositories(sender, e);
                organization.Items.Clear(); organization.Items.Add(UiText.Get("Personal account"));
                organization.Items.AddRange((await api.Organizations(ct)).Select(x => x.login).ToArray()); organization.SelectedIndex = 0;
            });
        }
        /// <summary>Filtre la liste des dépôts sur le texte de recherche courant.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void FilterRepositories(object sender, EventArgs e)
        {
            repositoryList.Items.Clear(); repositoryList.Items.AddRange(repositories.Where(x => (x.full_name ?? "").IndexOf(repositorySearch.Text, StringComparison.OrdinalIgnoreCase) >= 0).ToArray());
        }
        /// <summary>Charge les branches du dépôt sélectionné et initialise la branche par défaut.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void RepositoryChanged(object sender, EventArgs e)
        {
            var repo = repositoryList.SelectedItem as GitHubRepositoryInfo;
            if (repo == null) return;
            repositoryBranch.Items.Clear(); repositoryBranch.Text = repo.default_branch ?? "main";
            await Run(async (api, ct) => { repositoryBranch.Items.AddRange((await api.Branches(repo.clone_url, ct)).Select(x => x.name).ToArray()); });
        }
        /// <summary>Valide la branche choisie et notifie le conteneur du dépôt à utiliser.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void UseRepository_Click(object sender, EventArgs e)
        {
            var repo = repositoryList.SelectedItem as GitHubRepositoryInfo;
            if (repo == null) return;
            try { MacroGitRepository.ValidateBranch(repositoryBranch.Text); RepositorySelected?.Invoke(repo.clone_url, repositoryBranch.Text); }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        /// <summary>Crée un dépôt GitHub puis l’ajoute à la sélection courante.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void CreateRepository_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                var created = await api.CreateRepository(repositoryName.Text.Trim(), organization.SelectedIndex > 0 ? organization.Text : null, privateRepository.Checked, ct);
                repositories = repositories.Concat(new[] { created }).ToArray(); repositorySearch.Clear(); FilterRepositories(sender, e);
                repositoryList.SelectedItem = created; repositoryBranch.Text = created.default_branch ?? "main";
                RepositorySelected?.Invoke(created.clone_url, repositoryBranch.Text);
            }, false);
        }
        /// <summary>Charge les demandes de fusion du dépôt et ses branches cibles.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void LoadPulls_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                pulls.Items.Clear(); pulls.Items.AddRange(await api.Pulls(remote, ct));
                targetBranch.Items.Clear(); targetBranch.Items.AddRange((await api.Branches(remote, ct)).Select(x => x.name).ToArray());
                if (targetBranch.Items.Contains("main")) targetBranch.SelectedItem = "main";
            });
        }
        /// <summary>Charge les détails, fichiers, commentaires et vérifications de la demande choisie.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
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
        /// <summary>Crée une demande de fusion depuis la branche locale vers la cible choisie.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private async void CreatePull_Click(object sender, EventArgs e)
        {
            await Run(async (api, ct) => {
                var created = await api.CreatePull(remote, branch, targetBranch.Text, pullTitle.Text, pullBody.Text, draft.Checked, ct);
                pulls.Items.Insert(0, created); selectedPull = created; pullDetails.Text = created.html_url;
            }, false);
        }
        /// <summary>Ouvre dans le navigateur l’URL de la demande de fusion sélectionnée.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void OpenPull_Click(object sender, EventArgs e)
        {
            if (selectedPull == null) return;
            try { OpenExternalLink(selectedPull.html_url); } catch (Exception ex) { status.Text = ex.Message; }
        }
        /// <summary>Ouvre dans l’éditeur un module VBA lié au fichier ou commentaire sélectionné.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
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
        /// <summary>Affiche le corps du commentaire actuellement sélectionné.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
                /// <param name="e">Données de l’événement de sélection.</param>
private void CommentChanged(object sender, EventArgs e) { commentBody.Text = (comments.SelectedItem as GitHubComment)?.body ?? ""; }
        /// <summary>Remplit le formulaire avec le brouillon préparé pour la branche.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void LoadDraft_Click(object sender, EventArgs e)
        {
            try {
                var prepared = LoadDraft?.Invoke();
                if (prepared == null) { status.Text = UiText.Get("No prepared draft for this branch."); return; }
                targetBranch.Text = prepared.Target; pullTitle.Text = prepared.Title; pullBody.Text = prepared.Body;
                pullTabs.SelectedTab = composeTab;
            } catch (Exception ex) { status.Text = ex.Message; }
        }
        /// <summary>Demande l’annulation de l’opération GitHub en cours.</summary>
        /// <param name="sender">Contrôle à l’origine de l’événement.</param>
                /// <param name="e">Données de l’événement de clic.</param>
private void Cancel_Click(object sender, EventArgs e) { cancellation?.Cancel(); }
        /// <summary>Annule l’opération en cours et libère les composants du panneau.</summary>
                /// <param name="disposing">Indique si les ressources gérées doivent être libérées.</param>
protected override void Dispose(bool disposing) { if (disposing) { cancellation?.Cancel(); components?.Dispose(); } base.Dispose(disposing); }
    }
}
