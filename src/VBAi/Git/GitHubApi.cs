using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Réponse HTTP refusée, sans corps distant ou secret dans le message.</summary>
    internal sealed class GitHubApiFailure : InvalidOperationException
    {
        /// <summary>Obtient le code HTTP à l’origine du refus structuré.</summary>
        /// <value>Code numérique renvoyé par GitHub.</value>
        internal int Status { get; }
        /// <summary>Crée une erreur API en conservant le statut HTTP pour le traitement de repli.</summary>
        /// <param name="status">Code HTTP renvoyé par GitHub.</param>
        /// <param name="message">Message d’erreur présenté à l’appelant.</param>
        internal GitHubApiFailure(int status, string message) : base(message) { Status = status; }
    }
    /// <summary>Issue créée sur GitHub.</summary>
    internal sealed class GitHubIssue
    {
        /// <summary>Obtient ou définit le numéro attribué à l’issue.</summary>
        /// <value>Numéro de l’issue dans le dépôt.</value>
        public int number { get; set; }
        /// <summary>Obtient ou définit l’URL Web de l’issue.</summary>
        /// <value>Adresse de l’issue sur GitHub.</value>
        public string html_url { get; set; }
    }
    /// <summary>Informations de dépôt renvoyées par l’API GitHub.</summary>
    internal sealed class GitHubRepositoryInfo
    {
        /// <summary>Obtient ou définit le nom complet propriétaire/dépôt.</summary>
        /// <value>le nom complet propriétaire/dépôt.</value>
        public string full_name { get; set; }
        /// <summary>Obtient ou définit l’URL de clonage HTTPS.</summary>
        /// <value>l’URL de clonage HTTPS.</value>
        public string clone_url { get; set; }
        /// <summary>Obtient ou définit le nom de la branche par défaut.</summary>
        /// <value>le nom de la branche par défaut.</value>
        public string default_branch { get; set; }
        /// <summary>Retourne le nom complet du dépôt.</summary>
        /// <returns>Valeur de <see cref="full_name"/>.</returns>
        public override string ToString() { return full_name; }
    }
    /// <summary>Informations sur une pull request GitHub.</summary>
    internal sealed class GitHubPull
    {
        /// <summary>Obtient ou définit le numéro de la pull request.</summary>
        /// <value>le numéro de la pull request.</value>
        public int number { get; set; }
        /// <summary>Obtient ou définit son titre.</summary>
        /// <value>son titre.</value>
        public string title { get; set; }
        /// <summary>Obtient ou définit sa description.</summary>
        /// <value>sa description.</value>
        public string body { get; set; }
        /// <summary>Obtient ou définit son état GitHub.</summary>
        /// <value>son état GitHub.</value>
        public string state { get; set; }
        /// <summary>Obtient ou définit si la pull request est un brouillon.</summary>
        /// <value>si la pull request est un brouillon.</value>
        public bool draft { get; set; }
        /// <summary>Obtient ou définit si elle a été fusionnée.</summary>
        /// <value>si elle a été fusionnée.</value>
        public bool merged { get; set; }
        /// <summary>Obtient ou définit son URL Web.</summary>
        /// <value>son URL Web.</value>
        public string html_url { get; set; }
        /// <summary>Obtient ou définit les informations de la tête de branche.</summary>
        /// <value>les informations de la tête de branche.</value>
        public GitHubHead head { get; set; }
        /// <summary>Retourne un libellé compact avec numéro, état et titre, en signalant les brouillons.</summary>
        /// <returns>Libellé localisé de la pull request.</returns>
        public override string ToString() { return "#" + number + " · " + state + (draft ? " · " + UiText.Get("Draft") : "") + " · " + title; }
    }
    /// <summary>Informations sur la révision de tête d’une pull request.</summary>
    internal sealed class GitHubHead
    {
        /// <summary>Obtient ou définit le SHA de la révision.</summary>
        /// <value>le SHA de la révision.</value>
        public string sha { get; set; }
    }
    /// <summary>Fichier modifié associé à une pull request.</summary>
    internal sealed class GitHubFile
    {
        /// <summary>Obtient ou définit le chemin du fichier.</summary>
        /// <value>le chemin du fichier.</value>
        public string filename { get; set; }
        /// <summary>Obtient ou définit l’état de modification du fichier.</summary>
        /// <value>l’état de modification du fichier.</value>
        public string status { get; set; }
        /// <summary>Retourne l’état suivi du chemin.</summary>
        /// <returns>Libellé compact du fichier.</returns>
        public override string ToString() { return status + " · " + filename; }
    }
    /// <summary>Commentaire GitHub avec son contenu et son emplacement facultatif.</summary>
    internal sealed class GitHubComment
    {
        /// <summary>Obtient ou définit le contenu du commentaire.</summary>
        /// <value>le contenu du commentaire.</value>
        public string body { get; set; }
        /// <summary>Obtient ou définit le chemin du fichier commenté.</summary>
        /// <value>le chemin du fichier commenté.</value>
        public string path { get; set; }
        /// <summary>Obtient ou définit le numéro de ligne, s’il est fourni.</summary>
        /// <value>le numéro de ligne, s’il est fourni.</value>
        public int? line { get; set; }
        /// <summary>Retourne le contenu précédé de l’emplacement lorsqu’un chemin est fourni.</summary>
        /// <returns>Libellé du commentaire.</returns>
        public override string ToString() { return (path == null ? "" : path + ":" + line + " · ") + body; }
    }
    /// <summary>État d’une exécution de vérification GitHub.</summary>
    internal sealed class GitHubCheck
    {
        /// <summary>Obtient ou définit le nom de la vérification.</summary>
        /// <value>le nom de la vérification.</value>
        public string name { get; set; }
        /// <summary>Obtient ou définit son état courant.</summary>
        /// <value>son état courant.</value>
        public string status { get; set; }
        /// <summary>Obtient ou définit sa conclusion éventuelle.</summary>
        /// <value>sa conclusion éventuelle.</value>
        public string conclusion { get; set; }
    }
    /// <summary>Enveloppe de la liste d’exécutions de vérification d’une révision.</summary>
    internal sealed class GitHubChecks
    {
        /// <summary>Obtient ou définit les vérifications renvoyées par l’API.</summary>
        /// <value>les vérifications renvoyées par l’API.</value>
        public GitHubCheck[] check_runs { get; set; }
    }
    /// <summary>État global de statut d’une révision.</summary>
    internal sealed class GitHubStatus
    {
        /// <summary>Obtient ou définit l’état agrégé.</summary>
        /// <value>l’état agrégé.</value>
        public string state { get; set; }
    }
    /// <summary>Organisation GitHub associée au compte.</summary>
    internal sealed class GitHubOrganization
    {
        /// <summary>Obtient ou définit le nom de connexion.</summary>
        /// <value>le nom de connexion.</value>
        public string login { get; set; }
    }
    /// <summary>Branche d’un dépôt GitHub.</summary>
    internal sealed class GitHubBranch
    {
        /// <summary>Obtient ou définit le nom de branche.</summary>
        /// <value>le nom de branche.</value>
        public string name { get; set; }
    }

    /// <summary>Appelle l’API GitHub avec les identifiants du gestionnaire GCM de l’utilisateur.</summary>
    internal sealed class GitHubApi : IDisposable
    {
        /// <summary>Démarre un processus de saisie d’identifiants sans préambule sur son entrée standard.</summary>
        internal static Func<Process, bool> StartCredentialProcess = ProcessInput.StartWithoutPreamble;
        /// <summary>Writes and closes credential input; permits deterministic input pipe failure tests.</summary>
        internal static Func<Process, string, Task> WriteCredentialInput = async (process, input) =>
        {
            await process.StandardInput.WriteAsync(input);
            process.StandardInput.Close();
        };
        /// <summary>Client HTTP utilisé pour les requêtes API.</summary>
        private readonly HttpClient client;
        /// <summary>Fournisseur asynchrone du jeton d’accès.</summary>
        private readonly Func<CancellationToken, Task<string>> credential;
        /// <summary>Sérialiseur des charges utiles JSON de l’API.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        /// <summary>Crée le client API et configure l’obtention du jeton depuis GCM si aucun fournisseur n’est injecté.</summary>
        /// <param name="account">Compte GitHub facultatif utilisé pour sélectionner les identifiants GCM.</param>
        /// <param name="handler">Gestionnaire HTTP facultatif, utile pour fournir un transport adapté.</param>
        /// <param name="credential">Fournisseur asynchrone facultatif du jeton d’accès.</param>
        internal GitHubApi(string account, HttpMessageHandler handler = null, Func<CancellationToken, Task<string>> credential = null)
        {
            client = new HttpClient(handler ?? new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(45) };
            this.credential = credential ?? (ct => ReadCredential(account, ct));
        }
        /// <summary>Libère le client HTTP détenu par cette instance.</summary>
        public void Dispose() { client.Dispose(); }
        /// <summary>Convertit une URL de dépôt validée en chemin de ressource GitHub.</summary>
        /// <param name="url">URL distante du dépôt.</param>
        /// <returns>Chemin API commençant par <c>/repos/</c>.</returns>
        internal static string RepositoryPath(string url)
        {
            var path = new Uri(MacroGitRepository.ValidateRemote(url)).AbsolutePath.Trim('/');
            if (path.EndsWith(".git", StringComparison.OrdinalIgnoreCase)) path = path.Substring(0, path.Length - 4);
            return "/repos/" + path;
        }
                /// <summary>Crée une issue avec un titre et un rapport explicites, sans étiquette nécessitant des droits supplémentaires.</summary>
                /// <param name="url">URL distante du dépôt cible.</param>
                /// <param name="title">Titre non vide de l’issue, limité à 180 caractères.</param>
                /// <param name="body">Rapport facultatif, limité à 60 000 caractères.</param>
                /// <param name="ct">Jeton d’annulation de la requête.</param>
                /// <returns>Issue créée avec son numéro et son URL.</returns>
                /// <exception cref="ArgumentException">Le titre est invalide ou le rapport dépasse la limite.</exception>
                /// <exception cref="InvalidOperationException">GitHub refuse la création ou est indisponible.</exception>
        internal Task<GitHubIssue> CreateIssue(string url, string title, string body, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(title) || title.Length > 180 || (body?.Length ?? 0) > 60000)
                throw new ArgumentException("Invalid issue title or body.");
            return Request<GitHubIssue>(HttpMethod.Post, RepositoryPath(url) + "/issues", new { title, body }, ct, true);
        }

        /// <summary>Envoie une requête authentifiée et désérialise sa réponse JSON.</summary>
        /// <typeparam name="T">Type du contenu JSON attendu.</typeparam>
        /// <param name="method">Méthode HTTP.</param>
        /// <param name="path">Chemin relatif de l’API, soumis à une validation contre les chemins externes.</param>
        /// <param name="data">Objet facultatif sérialisé comme corps JSON.</param>
        /// <param name="ct">Jeton d’annulation de la requête.</param>
        /// <returns>Contenu de la réponse désérialisé en <typeparamref name="T"/>.</returns>
        /// <exception cref="ArgumentException">Le chemin ne respecte pas le format API relatif attendu.</exception>
        /// <exception cref="InvalidOperationException">GitHub refuse la requête ou est indisponible.</exception>
        /// <exception cref="OperationCanceledException">La requête a été annulée.</exception>
        /// <param name="structuredFailure">Expose le statut HTTP pour la stratégie de repli des rapports.</param>
        internal async Task<T> Request<T>(HttpMethod method, string path, object data, CancellationToken ct, bool structuredFailure = false)
        {
            if (!path.StartsWith("/", StringComparison.Ordinal) || path.StartsWith("//", StringComparison.Ordinal) || path.Contains("\\") || path.Contains("..")) throw new ArgumentException("Invalid GitHub API path.");
            using (var request = new HttpRequestMessage(method, "https://api.github.com" + path))
            {
                request.Headers.UserAgent.ParseAdd("VBAi/1.0");
                request.Headers.Accept.ParseAdd("application/vnd.github+json");
                request.Headers.Add("X-GitHub-Api-Version", "2026-03-10");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await credential(ct));
                if (data != null) request.Content = new StringContent(json.Serialize(data), Encoding.UTF8, "application/json");
                using (var response = await client.SendAsync(request, ct))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        int code = (int)response.StatusCode;
                        string help = code == 401 ? "Sign in to GitHub again in Settings." : code == 403 || code == 429 ? "Check repository permissions or wait for the GitHub rate limit to reset." :
                            code == 404 ? "Check the repository name and account access." : code == 422 ? "Check the branch, title and repository name; a pull request may already exist." : "GitHub is unavailable. Try again later.";
                        if (structuredFailure) throw new GitHubApiFailure(code, "GitHub " + code + " · " + UiText.Get(help));
                        throw new InvalidOperationException("GitHub " + code + " · " + UiText.Get(help));
                    }
                    string body = await response.Content.ReadAsStringAsync();
                    ct.ThrowIfCancellationRequested();
                    return json.Deserialize<T>(body);
                }
            }
        }
        /// <summary>Récupère toutes les pages d’une ressource de liste GitHub, par lots de cent éléments.</summary>
        /// <typeparam name="T">Type de chaque élément de la liste.</typeparam>
        /// <param name="path">Chemin API de la liste.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Éléments concaténés dans leur ordre de page.</returns>
        internal async Task<T[]> List<T>(string path, CancellationToken ct)
        {
            var result = new List<T>();
            for (int page = 1; ; page++)
            {
                ct.ThrowIfCancellationRequested();
                var items = await Request<T[]>(HttpMethod.Get, path + (path.Contains("?") ? "&" : "?") + "per_page=100&page=" + page, null, ct);
                result.AddRange(items);
                if (items.Length < 100) return result.ToArray();
            }
        }
        /// <summary>Liste les dépôts accessibles au compte, triés par mise à jour.</summary>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Dépôts dont l’utilisateur est propriétaire, collaborateur ou membre de l’organisation.</returns>
        internal Task<GitHubRepositoryInfo[]> Repositories(CancellationToken ct) { return List<GitHubRepositoryInfo>("/user/repos?sort=updated&affiliation=owner,collaborator,organization_member", ct); }
        /// <summary>Liste les organisations visibles pour le compte.</summary>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Organisations renvoyées par GitHub.</returns>
        internal Task<GitHubOrganization[]> Organizations(CancellationToken ct) { return List<GitHubOrganization>("/user/orgs", ct); }
        /// <summary>Liste les branches du dépôt identifié par son URL distante.</summary>
        /// <param name="url">URL distante du dépôt.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Branches du dépôt.</returns>
        internal Task<GitHubBranch[]> Branches(string url, CancellationToken ct) { return List<GitHubBranch>(RepositoryPath(url) + "/branches", ct); }
        /// <summary>Crée un dépôt utilisateur ou organisation sans initialiser de commit.</summary>
        /// <param name="name">Nom du dépôt, composé de 1 à 100 caractères autorisés.</param>
        /// <param name="organization">Compte d’organisation facultatif ; vide pour un dépôt utilisateur.</param>
        /// <param name="isPrivate">Indique si le dépôt doit être privé.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Informations sur le dépôt créé.</returns>
        /// <exception cref="ArgumentException">Le nom du dépôt ou le compte d’organisation est invalide.</exception>
        internal Task<GitHubRepositoryInfo> CreateRepository(string name, string organization, bool isPrivate, CancellationToken ct)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(name ?? "", "^[A-Za-z0-9_.-]{1,100}$")) throw new ArgumentException(UiText.Get("Invalid repository name."));
            if (!string.IsNullOrEmpty(organization) && !GitHubAccountService.ValidAccount(organization)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            return Request<GitHubRepositoryInfo>(HttpMethod.Post, string.IsNullOrEmpty(organization) ? "/user/repos" : "/orgs/" + organization + "/repos", new { name, @private = isPrivate, auto_init = false }, ct);
        }
        /// <summary>Liste les pull requests ouvertes et fermées du dépôt, triées par mise à jour.</summary>
        /// <param name="url">URL distante du dépôt.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Pull requests du dépôt.</returns>
        internal Task<GitHubPull[]> Pulls(string url, CancellationToken ct) { return List<GitHubPull>(RepositoryPath(url) + "/pulls?state=all&sort=updated", ct); }
        /// <summary>Crée une pull request entre deux branches différentes.</summary>
        /// <param name="url">URL distante du dépôt.</param>
        /// <param name="head">Branche source.</param>
        /// <param name="target">Branche cible.</param>
        /// <param name="title">Titre non vide de la pull request.</param>
        /// <param name="body">Description facultative.</param>
        /// <param name="draft">Indique si elle doit être créée comme brouillon.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>Pull request renvoyée par GitHub.</returns>
        /// <exception cref="ArgumentException">Une branche est invalide, le titre est vide ou source et cible sont identiques.</exception>
        internal Task<GitHubPull> CreatePull(string url, string head, string target, string title, string body, bool draft, CancellationToken ct)
        {
            MacroGitRepository.ValidateBranch(head); MacroGitRepository.ValidateBranch(target);
            if (string.IsNullOrWhiteSpace(title) || head == target) throw new ArgumentException(UiText.Get("Enter a title and a different target branch."));
            return Request<GitHubPull>(HttpMethod.Post, RepositoryPath(url) + "/pulls", new { head, @base = target, title, body, draft }, ct);
        }
        /// <summary>Récupère les exécutions de vérification et le statut global d’un commit.</summary>
        /// <param name="url">URL distante du dépôt.</param>
        /// <param name="sha">SHA-1 du commit sous forme de quarante caractères hexadécimaux minuscules.</param>
        /// <param name="ct">Jeton d’annulation.</param>
        /// <returns>État global suivi des conclusions ou états des vérifications.</returns>
        /// <exception cref="ArgumentException">Le SHA n’a pas le format attendu.</exception>
        internal async Task<string> Checks(string url, string sha, CancellationToken ct)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(sha ?? "", "^[a-f0-9]{40}$")) throw new ArgumentException("Invalid commit.");
            var lines = new List<string>();
            for (int page = 1; ; page++)
            {
                var checks = await Request<GitHubChecks>(HttpMethod.Get, RepositoryPath(url) + "/commits/" + sha + "/check-runs?per_page=100&page=" + page, null, ct);
                lines.AddRange(checks.check_runs.Select(x => x.name + ": " + (x.conclusion ?? x.status)));
                if (checks.check_runs.Length < 100) break;
            }
            var status = await Request<GitHubStatus>(HttpMethod.Get, RepositoryPath(url) + "/commits/" + sha + "/status", null, ct);
            return UiText.Get("Commit status") + ": " + status.state + Environment.NewLine + string.Join(Environment.NewLine, lines);
        }
        /// <summary>Demande à Git Credential Manager le jeton HTTPS GitHub associé au compte.</summary>
        /// <param name="account">Compte facultatif utilisé pour sélectionner l’identifiant.</param>
        /// <param name="ct">Jeton d’annulation ; une limite interne de trente secondes est également appliquée.</param>
        /// <returns>Jeton d’accès lu depuis la réponse du gestionnaire d’identifiants.</returns>
        /// <exception cref="ArgumentException">Le compte fourni est invalide.</exception>
        /// <exception cref="InvalidOperationException">GCM ne fournit pas de jeton utilisable.</exception>
        /// <exception cref="OperationCanceledException">L’opération est annulée ou dépasse son délai.</exception>
        internal static async Task<string> ReadCredential(string account, CancellationToken ct)
        {
            if (!string.IsNullOrEmpty(account) && !GitHubAccountService.ValidAccount(account)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            var start = new ProcessStartInfo("git.exe", "credential-manager get") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true };
            start.EnvironmentVariables["GCM_INTERACTIVE"] = "never";
            start.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            using (var process = new Process { StartInfo = start })
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(30));
                StartCredentialProcess(process);
                using (timeout.Token.Register(() => { try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { } }))
                {
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    try
                    {
                        await WriteCredentialInput(process, "protocol=https\nhost=github.com\n" + (string.IsNullOrEmpty(account) ? "" : "username=" + account + "\n") + "\n");
                    }
                    catch (System.IO.IOException) when (timeout.IsCancellationRequested)
                    {
                        // Killing the cancelled child can break the input pipe before the final cancellation check.
                        throw new OperationCanceledException(timeout.Token);
                    }
                    await Task.Run(() => process.WaitForExit());
                    await error; string result = await output;
                    timeout.Token.ThrowIfCancellationRequested();
                    string token = result.Split('\n').FirstOrDefault(x => x.StartsWith("password=", StringComparison.Ordinal));
                    if (process.ExitCode != 0 || token == null) throw new InvalidOperationException(UiText.Get("Sign in to GitHub again in Settings."));
                    return token.Substring(9).TrimEnd('\r');
                }
            }
        }
    }
}
