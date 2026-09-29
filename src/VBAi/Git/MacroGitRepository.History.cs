using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Point de contrôle enregistré dans les références privées du dépôt.</summary>
    internal sealed class GitCheckpoint
    {
        /// <summary>Identifiant unique de ce point de contrôle.</summary>
        /// <value>Identifiant unique de ce point de contrôle.</value>
        public string Id { get; set; }
        /// <summary>Libellé affiché dans la liste des points de contrôle.</summary>
        /// <value>Libellé affiché dans la liste des points de contrôle.</value>
        public string Label { get; set; }
        /// <summary>Commit Git capturé par ce point de contrôle.</summary>
        /// <value>Commit Git capturé par ce point de contrôle.</value>
        public string Commit { get; set; }
        /// <summary>Retourne le libellé suivi de l’identifiant du point de contrôle.</summary>
        /// <returns>Libellé, séparateur et identifiant du point de contrôle.</returns>
        public override string ToString() { return Label + " · " + Id; }
    }
    /// <summary>Plan de fusion calculé depuis les deux commits et l’arbre de résultat.</summary>
    internal sealed class GitMergePlan
    {
        /// <summary>Branche locale au début de la fusion.</summary>
        /// <value>Branche locale au début de la fusion.</value>
        public string Branch { get; set; }
        /// <summary>Commit local retenu comme côté courant.</summary>
        /// <value>Commit local retenu comme côté courant.</value>
        public string Ours { get; set; }
        /// <summary>Commit de la branche fusionnée.</summary>
        /// <value>Commit de la branche fusionnée.</value>
        public string Theirs { get; set; }
        /// <summary>Arbre Git résultant, y compris les résolutions déjà enregistrées.</summary>
        /// <value>Arbre Git résultant, y compris les résolutions déjà enregistrées.</value>
        public string Tree { get; set; }
        /// <summary>Chemins encore en conflit dans cet arbre.</summary>
        /// <value>Chemins encore en conflit dans cet arbre.</value>
        public string[] Conflicts { get; set; }
    }
    /// <summary>Contenu textuel de base, local et entrant d’un conflit.</summary>
    internal sealed class GitConflictContent
    {
        /// <summary>Texte de la base commune, si disponible.</summary>
        /// <value>Texte de la base commune, si disponible.</value>
        public string Base { get; set; }
        /// <summary>Chemin du fichier concerné dans le paquet VBA.</summary>
        /// <value>Chemin du fichier concerné dans le paquet VBA.</value>
        public string Path { get; set; }
        /// <summary>Commit local retenu comme côté courant.</summary>
        /// <value>Commit local retenu comme côté courant.</value>
        public string Ours { get; set; }
        /// <summary>Commit de la branche fusionnée.</summary>
        /// <value>Commit de la branche fusionnée.</value>
        public string Theirs { get; set; }
    }
    /// <summary>Opérations de branches, points de contrôle et fusion sur le dépôt bare.</summary>
    internal sealed partial class MacroGitRepository
    {
        /// <summary>Valide un nom de branche selon les règles acceptées par le dépôt.</summary>
        /// <param name="name">Nom de branche à vérifier.</param>
        internal static void ValidateBranch(string name) { new MacroGitRepository(Path.GetTempPath(), name); }
        /// <summary>Retourne les branches locales du dépôt.</summary>
        /// <returns>Noms des branches locales.</returns>
        internal string[] Branches() { return Text("for-each-ref", "--format=%(refname:strip=2)", "refs/heads/").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); }
        /// <summary>Crée une branche locale au commit HEAD courant.</summary>
        /// <param name="name">Nom de branche à valider.</param>
        internal void CreateBranch(string name)
        {
            ValidateBranch(name);
            string head = Resolve(Head) ?? throw new InvalidOperationException(UiText.Get("Create a commit first."));
            Text("update-ref", "refs/heads/" + name, head, new string('0', 40));
        }
        /// <summary>Résout le commit de la branche locale demandée.</summary>
        /// <param name="name">Nom de branche à valider.</param>
        /// <returns>Identifiant du commit de la branche.</returns>
        internal string BranchCommit(string name)
        {
            ValidateBranch(name);
            return Resolve("refs/heads/" + name) ?? throw new InvalidOperationException(UiText.Get("Unknown local branch. Create it or fetch it from the remote repository."));
        }
        /// <summary>Sélectionne une branche existante comme branche active du dépôt.</summary>
        /// <param name="name">Nom de branche à valider.</param>
        internal void SelectBranch(string name)
        {
            BranchCommit(name);
            Text("config", "codex.activeBranch", name);
            Branch = name;
            string previous = Resolve("refs/remotes/origin/selected");
            if (previous != null) Text("update-ref", "-d", "refs/remotes/origin/selected", previous);
        }
        /// <summary>Retourne les noms des branches annoncées par origin.</summary>
        /// <returns>Noms des branches distantes annoncées par origin.</returns>
        internal string[] RemoteBranches()
        {
            return Text("ls-remote", "--heads", "origin").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Substring(x.IndexOf("refs/heads/", StringComparison.Ordinal) + 11)).ToArray();
        }
        /// <summary>Récupère une branche distante qui n’existe pas encore localement.</summary>
        /// <param name="name">Nom de branche à valider.</param>
        internal void TrackBranch(string name)
        {
            ValidateBranch(name);
            if (Resolve("refs/heads/" + name) != null) throw new InvalidOperationException(UiText.Get("This local branch already exists."));
            Text("fetch", "--no-tags", "origin", "refs/heads/" + name + ":refs/heads/" + name);
        }
        /// <summary>Crée un commit de point de contrôle et une référence privée associée.</summary>
        /// <param name="live">Snapshot courant des fichiers VBA.</param>
        /// <param name="label">Libellé lisible du point de contrôle.</param>
        /// <returns>Point de contrôle nouvellement enregistré.</returns>
        internal GitCheckpoint Checkpoint(VbaGitSnapshot live, string label)
        {
            if (string.IsNullOrWhiteSpace(label) || label.Length > 160 || label.Any(char.IsControl))
                throw new ArgumentException(UiText.Get("A checkpoint name of 1 to 160 characters is required."));
            string id = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            string commit = Commit(live, null, label);
            SetRef("refs/codex/checkpoints/" + id, commit);
            return new GitCheckpoint { Id = id, Commit = commit, Label = label };
        }
        /// <summary>Retourne les points de contrôle privés les plus récents.</summary>
        /// <returns>Points de contrôle triés du plus récent au plus ancien.</returns>
        internal GitCheckpoint[] Checkpoints()
        {
            return Text("for-each-ref", "--sort=-refname", "--format=%(refname:strip=3)%09%(objectname)%09%(subject)", "refs/codex/checkpoints/")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(line => {
                    var parts = line.Split(new[] { '\t' }, 3); return new GitCheckpoint { Id = parts[0], Commit = parts[1], Label = parts[2] };
                }).ToArray();
        }
        /// <summary>Résout un identifiant de point de contrôle après validation de son format.</summary>
        /// <param name="id">Identifiant du point de contrôle.</param>
        /// <returns>Identifiant du commit du point de contrôle.</returns>
        internal string CheckpointCommit(string id)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(id ?? "", @"^\d{17}-[0-9a-f]{8}$")) throw new ArgumentException("Identifiant de checkpoint invalide.");
            return Resolve("refs/codex/checkpoints/" + id) ?? throw new InvalidOperationException("Checkpoint absent.");
        }
        /// <summary>Chemin du fichier d’état du plan de fusion.</summary>
        /// <value>Chemin du fichier d’état du plan de fusion.</value>
        private string MergeFile { get { return Path.Combine(directory, "codex-merge.json"); } }
        /// <summary>Plan de fusion enregistré, ou null si aucune fusion n’est en cours.</summary>
        /// <value>Plan de fusion enregistré, ou null si aucune fusion n’est en cours.</value>
        internal GitMergePlan PendingMerge { get { return File.Exists(MergeFile) ? new JavaScriptSerializer().Deserialize<GitMergePlan>(File.ReadAllText(MergeFile)) : null; } }
        /// <summary>Persiste atomiquement le plan et conserve son arbre par une référence Git.</summary>
        /// <param name="plan">Plan de fusion à vérifier ou finaliser.</param>
        private void SaveMerge(GitMergePlan plan)
        {
            string temporary = MergeFile + ".pending";
            File.WriteAllText(temporary, new JavaScriptSerializer().Serialize(plan), VbaGitSnapshot.Utf8);
            if (File.Exists(MergeFile)) File.Replace(temporary, MergeFile, null); else File.Move(temporary, MergeFile);
            // Keep generated merge trees reachable during cache maintenance.
            SetRef("refs/codex/merge-tree", plan.Tree);
        }
        /// <summary>Calcule et enregistre un plan de fusion sans modifier la branche active.</summary>
        /// <param name="branch">Nom de la branche source ou cible.</param>
        /// <returns>Plan de fusion calculé et persisté.</returns>
        internal GitMergePlan BeginMerge(string branch)
        {
            if (PendingMerge != null) throw new InvalidOperationException(UiText.Get("A merge is already in progress."));
            string ours = Resolve(Head) ?? throw new InvalidOperationException(UiText.Get("No local commit."));
            string theirs = BranchCommit(branch);
            var output = Run(new[] { "merge-tree", "--write-tree", "--name-only", "-z", ours, theirs }, null, true, true);
            if (output.ExitCode != 0 && output.ExitCode != 1) throw new InvalidOperationException(UiText.Get("Unable to merge: check common ancestors and Git 2.38 or later."));
            var records = Encoding.UTF8.GetString(output.Bytes).Split('\0');
            var plan = new GitMergePlan { Branch = Branch, Ours = ours, Theirs = theirs, Tree = records[0].Trim(),
                Conflicts = output.ExitCode == 0 ? new string[0] : records.Skip(1).TakeWhile(x => x.Length != 0).ToArray() };
            if (!System.Text.RegularExpressions.Regex.IsMatch(plan.Tree, "^[0-9a-f]{40}$")) throw new InvalidOperationException(UiText.Get("Invalid Git merge result."));
            SaveMerge(plan); return plan;
        }
        /// <summary>Vérifie que le plan correspond toujours à la branche et au HEAD actifs.</summary>
        /// <param name="plan">Plan de fusion à vérifier ou finaliser.</param>
        internal void AssertMerge(GitMergePlan plan)
        {
            if (plan == null || plan.Branch != Branch || plan.Ours != Resolve(Head)) throw new InvalidOperationException(UiText.Get("The branch changed since the merge started."));
        }
        /// <summary>Retourne les contenus de base, local et entrant pour un chemin en conflit.</summary>
        /// <param name="path">Chemin du fichier dans le paquet VBA.</param>
        /// <returns>Version de base, locale et entrante du conflit.</returns>
        internal GitConflictContent ConflictContent(string path)
        {
            var plan = PendingMerge; AssertMerge(plan);
            if (!plan.Conflicts.Contains(path)) throw new ArgumentException(UiText.Get("Path not found in conflicts."));
            string common = Text("merge-base", plan.Ours, plan.Theirs);
            return new GitConflictContent { Path = path, Base = ConflictText(common, path), Ours = ConflictText(plan.Ours, path), Theirs = ConflictText(plan.Theirs, path) };
        }
        /// <summary>Lit le contenu texte d’un fichier dans un arbre si sa taille et son encodage le permettent.</summary>
        /// <param name="tree">Identifiant de commit ou arbre Git à lire.</param>
        /// <param name="path">Chemin du fichier dans le paquet VBA.</param>
        /// <returns>Texte décodé, ou indication si la ressource est binaire ou trop volumineuse.</returns>
        private string ConflictText(string tree, string path)
        {
            string entry = TreeAtPath(tree, path);
            if (entry == null) return "";
            var parts = entry.Split(' ', '\t');
            if (parts[1] != "blob" || path.EndsWith(".frx", StringComparison.OrdinalIgnoreCase)) return UiText.Get("[Binary resource or directory: choose local or incoming]");
            if (long.Parse(Text("cat-file", "-s", parts[2])) > 65536) return UiText.Get("[Preview unavailable above 64 KiB: choose local or incoming]");
            try { return VbaGitSnapshot.Utf8.GetString(Run(new[] { "cat-file", "blob", parts[2] }).Bytes); }
            catch (DecoderFallbackException) { return UiText.Get("[Non-UTF-8 content: choose local or incoming]"); }
        }
        /// <summary>Résout un conflit avec la version locale, entrante ou un texte VBA fourni.</summary>
        /// <param name="path">Chemin du fichier dans le paquet VBA.</param>
        /// <param name="choice">Résolution choisie : ours, theirs ou text.</param>
        /// <param name="text">Contenu texte utilisé pour résoudre un conflit VBA.</param>
        /// <returns>Plan actualisé après résolution du conflit.</returns>
        internal GitMergePlan ResolveConflict(string path, string choice, string text = null)
        {
            var plan = PendingMerge; AssertMerge(plan);
            if (!plan.Conflicts.Contains(path)) throw new InvalidOperationException(UiText.Get("This file is not listed in conflicts."));
            string entry;
            if (choice == "ours" || choice == "theirs")
            {
                var selected = TreeAtPath(choice == "ours" ? plan.Ours : plan.Theirs, path);
                entry = selected;
            }
            else if (choice == "text" && path.StartsWith("vba/", StringComparison.Ordinal) && !path.EndsWith(".frx", StringComparison.Ordinal) && text != null && text.Length < 1024 * 1024)
                entry = "100644 blob " + Encoding.UTF8.GetString(Run(new[] { "hash-object", "-w", "--stdin" }, VbaGitSnapshot.Utf8.GetBytes(text)).Bytes).Trim() + "\t" + path.Split('/').Last();
            else throw new ArgumentException(UiText.Get("Choose ours, theirs or text for a VBA text source."));
            plan.Tree = ReplacePath(plan.Tree, path.Split('/'), 0, entry);
            plan.Conflicts = plan.Conflicts.Where(x => x != path).ToArray(); SaveMerge(plan); return plan;
        }
        /// <summary>Résout l’entrée Git située au chemin de fichier demandé.</summary>
        /// <param name="tree">Identifiant de commit ou arbre Git à lire.</param>
        /// <param name="path">Chemin du fichier dans le paquet VBA.</param>
        /// <returns>Entrée Git du chemin, ou null si elle est absente.</returns>
        private string TreeAtPath(string tree, string path)
        {
            string entry = null;
            foreach (string name in path.Split('/'))
            {
                entry = Tree(tree).FirstOrDefault(x => x.EndsWith("\t" + name, StringComparison.Ordinal));
                if (entry == null) return null;
                tree = entry.Split(' ', '\t')[2];
            }
            return entry;
        }
        /// <summary>Remplace ou supprime récursivement une entrée dans un arbre Git.</summary>
        /// <param name="tree">Identifiant de commit ou arbre Git à lire.</param>
        /// <param name="path">Chemin du fichier dans le paquet VBA.</param>
        /// <param name="index">Position du segment courant dans le chemin.</param>
        /// <param name="replacement">Entrée de remplacement, ou null pour supprimer le chemin.</param>
        /// <returns>Identifiant de l’arbre reconstruit avec le chemin remplacé.</returns>
        private string ReplacePath(string tree, string[] path, int index, string replacement)
        {
            var entries = tree == null ? new List<string>() : Tree(tree);
            string old = entries.FirstOrDefault(x => x.EndsWith("\t" + path[index], StringComparison.Ordinal));
            if (old != null) entries.Remove(old);
            if (index == path.Length - 1) { if (replacement != null) entries.Add(replacement); }
            else
            {
                if (old != null && !old.StartsWith("040000 tree ", StringComparison.Ordinal)) throw new InvalidOperationException(UiText.Get("File/directory conflict: choose the parent directory first."));
                string child = ReplacePath(old?.Split(' ', '\t')[2], path, index + 1, replacement);
                entries.Add("040000 tree " + child + "\t" + path[index]);
            }
            return entries.Count == 0 ? Encoding.UTF8.GetString(Run(new[] { "mktree" }, new byte[0]).Bytes).Trim() : MakeTree(entries);
        }
        /// <summary>Crée un commit de fusion après résolution et validation du paquet VBA.</summary>
        /// <param name="plan">Plan de fusion à vérifier ou finaliser.</param>
        /// <param name="message">Message du commit à créer.</param>
        /// <returns>Identifiant du commit de fusion créé.</returns>
        internal string MergeCommit(GitMergePlan plan, string message)
        {
            if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException(UiText.Get("A merge commit message is required."));
            AssertMerge(plan);
            if (plan.Conflicts.Length != 0) throw new InvalidOperationException(UiText.Get("Resolve all conflicts before completing the merge."));
            Read(plan.Tree); // Validate the complete VBA package before creating the merge commit.
            return Encoding.UTF8.GetString(Run(new[] { "commit-tree", plan.Tree, "-p", plan.Ours, "-p", plan.Theirs }, VbaGitSnapshot.Utf8.GetBytes(message + "\n")).Bytes).Trim();
        }
        /// <summary>Supprime l’état persistant de la fusion en cours.</summary>
        internal void AbortMerge() { if (File.Exists(MergeFile)) File.Delete(MergeFile); }
    }
}
