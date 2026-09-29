using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace VBAi
{
    /// <summary>Brouillon local destiné à la préparation d’une demande de fusion.</summary>
    internal sealed class GitPullDraft
    {
        /// <summary>Branche locale depuis laquelle la demande sera créée.</summary>
        /// <value>Branche source capturée au moment de la préparation.</value>
        public string Branch { get; set; }
        /// <summary>Branche cible de la demande.</summary>
        /// <value>Nom de branche validé comme cible.</value>
        public string Target { get; set; }
        /// <summary>Titre de la demande de fusion.</summary>
        /// <value>Titre non vide de la demande.</value>
        public string Title { get; set; }
        /// <summary>Description facultative de la demande.</summary>
        /// <value>Corps de la demande, limité en longueur lors de l’enregistrement.</value>
        public string Body { get; set; }
    }
    /// <summary>Métadonnées affichées pour un commit local.</summary>
    internal sealed class GitCommitInfo
    {
        /// <summary>Identifiant Git complet du commit.</summary>
        /// <value>Identifiant Git complet du commit.</value>
        public string Id { get; set; }
        /// <summary>Nom d’auteur fourni par Git.</summary>
        /// <value>Nom d’auteur fourni par Git.</value>
        public string Author { get; set; }
        /// <summary>Date du commit au format ISO.</summary>
        /// <value>Date du commit au format ISO.</value>
        public string Date { get; set; }
        /// <summary>Sujet du commit.</summary>
        /// <value>Sujet du commit.</value>
        public string Subject { get; set; }
        /// <summary>Retourne un résumé compact du commit.</summary>
        /// <returns>Identifiant court, date et sujet du commit.</returns>
        public override string ToString() { return Id.Substring(0, 8) + " · " + Date + " · " + Subject; }
    }
    /// <summary>Expose l’historique et le brouillon de pull request du dépôt Git bare.</summary>
    internal sealed partial class MacroGitRepository
    {
        /// <summary>Jeton utilisé pour annuler les opérations Git asynchrones.</summary>
        /// <value>Jeton d’annulation assigné pour l’opération active.</value>
        internal CancellationToken Cancellation { get; set; }
        /// <summary>Notifie l’interface de la progression de l’opération Git.</summary>
        /// <value>Action de progression, ou null si aucune interface n’est abonnée.</value>
        internal Action<string> Progress { get; set; }
        /// <summary>URL HTTPS normalisée du dépôt origin.</summary>
        /// <value>URL validée obtenue depuis la configuration Git active.</value>
        internal string RemoteUrl { get { return ValidateRemote(Text("remote", "get-url", "origin")); } }
        /// <summary>Valide un préfixe de commit et le résout en commit existant.</summary>
        /// <param name="id">Préfixe hexadécimal fourni par l’historique Git.</param>
        /// <returns>Identifiant complet du commit résolu.</returns>
        internal string VerifiedCommit(string id)
        {
            if (!Regex.IsMatch(id ?? "", "^[a-fA-F0-9]{7,40}$")) throw new ArgumentException(UiText.Get("Select a commit from history."));
            return Resolve(id + "^{commit}") ?? throw new ArgumentException(UiText.Get("Commit not found."));
        }
        /// <summary>Retourne jusqu’aux deux cents commits de la branche active.</summary>
        /// <returns>Commits de la branche active dans l’ordre du plus récent au plus ancien.</returns>
        internal GitCommitInfo[] Commits()
        {
            string head = Resolve(Head);
            if (head == null) return new GitCommitInfo[0];
            return Text("log", "-200", "--date=iso-strict", "--format=%H%x09%an%x09%ad%x09%s", head)
                .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(line => {
                    var fields = line.TrimEnd('\r').Split(new[] { '\t' }, 4);
                    return new GitCommitInfo { Id = fields[0], Author = fields[1], Date = fields[2], Subject = fields[3] };
                }).ToArray();
        }
        /// <summary>Retourne les métadonnées complètes d’un commit validé.</summary>
        /// <param name="id">Préfixe hexadécimal fourni par l’historique Git.</param>
        /// <returns>Sortie détaillée de git show sans diff.</returns>
        internal string CommitDetails(string id) { return Text("show", "--no-patch", "--format=fuller", VerifiedCommit(id)); }
        /// <summary>Retourne le premier parent du commit, ou null pour une racine.</summary>
        /// <param name="id">Préfixe hexadécimal fourni par l’historique Git.</param>
        /// <returns>Identifiant du premier parent, ou null si le commit est une racine.</returns>
        internal string ParentCommit(string id) { return Resolve(VerifiedCommit(id) + "^1"); }
        /// <summary>Valide et enregistre atomiquement le brouillon pour la branche active.</summary>
        /// <param name="target">Nom de la branche cible de la demande.</param>
        /// <param name="title">Titre non vide, limité à 256 caractères.</param>
        /// <param name="body">Description de la demande, limitée à 60 000 caractères.</param>
        internal void SavePullDraft(string target, string title, string body)
        {
            ValidateBranch(target);
            if (string.IsNullOrWhiteSpace(title) || title.Length > 256 || (body ?? "").Length > 60000) throw new ArgumentException(UiText.Get("Invalid pull request draft."));
            string file = System.IO.Path.Combine(directory, "codex-pr-draft.json");
            string pending = file + ".pending";
            System.IO.File.WriteAllText(pending, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(new GitPullDraft { Branch = Branch, Target = target, Title = title, Body = body }));
            if (System.IO.File.Exists(file)) System.IO.File.Replace(pending, file, null);
            else System.IO.File.Move(pending, file);
        }
        /// <summary>Charge le brouillon seulement s’il appartient encore à la branche active.</summary>
        /// <returns>Brouillon stocké pour la branche active, ou null.</returns>
        internal GitPullDraft PullDraft()
        {
            string file = System.IO.Path.Combine(directory, "codex-pr-draft.json");
            if (!System.IO.File.Exists(file)) return null;
            var draft = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<GitPullDraft>(System.IO.File.ReadAllText(file));
            return draft.Branch == Branch ? draft : null;
        }
    }
}
