using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class GitCheckpoint
    {
        public string Id { get; set; }
        public string Label { get; set; }
        public string Commit { get; set; }
        public override string ToString() { return Label + " · " + Id; }
    }
    internal sealed class GitMergePlan
    {
        public string Branch { get; set; }
        public string Ours { get; set; }
        public string Theirs { get; set; }
        public string Tree { get; set; }
        public string[] Conflicts { get; set; }
    }
    internal sealed class GitConflictContent
    {
        public string Path { get; set; }
        public string Ours { get; set; }
        public string Theirs { get; set; }
    }
    internal sealed partial class MacroGitRepository
    {
        internal static void ValidateBranch(string name) { new MacroGitRepository(Path.GetTempPath(), name); }
        internal string[] Branches() { return Text("for-each-ref", "--format=%(refname:strip=2)", "refs/heads/").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); }
        internal void CreateBranch(string name)
        {
            ValidateBranch(name);
            string head = Resolve(Head) ?? throw new InvalidOperationException("Créez d’abord un commit.");
            Text("update-ref", "refs/heads/" + name, head, new string('0', 40));
        }
        internal string BranchCommit(string name)
        {
            ValidateBranch(name);
            return Resolve("refs/heads/" + name) ?? throw new InvalidOperationException("Branche locale inconnue. Créez-la ou récupérez-la depuis le dépôt distant.");
        }
        internal void SelectBranch(string name)
        {
            BranchCommit(name);
            Text("config", "codex.activeBranch", name);
            Branch = name;
            string previous = Resolve("refs/remotes/origin/selected");
            if (previous != null) Text("update-ref", "-d", "refs/remotes/origin/selected", previous);
        }
        internal string[] RemoteBranches()
        {
            return Text("ls-remote", "--heads", "origin").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Substring(x.IndexOf("refs/heads/", StringComparison.Ordinal) + 11)).ToArray();
        }
        internal void TrackBranch(string name)
        {
            ValidateBranch(name);
            if (Resolve("refs/heads/" + name) != null) throw new InvalidOperationException("Cette branche locale existe déjà.");
            Text("fetch", "--no-tags", "origin", "refs/heads/" + name + ":refs/heads/" + name);
        }
        internal GitCheckpoint Checkpoint(VbaGitSnapshot live, string label)
        {
            if (string.IsNullOrWhiteSpace(label) || label.Length > 160 || label.Any(char.IsControl))
                throw new ArgumentException("Un nom de checkpoint de 1 à 160 caractères est requis.");
            string id = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            string commit = Commit(live, null, label);
            SetRef("refs/codex/checkpoints/" + id, commit);
            return new GitCheckpoint { Id = id, Commit = commit, Label = label };
        }
        internal GitCheckpoint[] Checkpoints()
        {
            return Text("for-each-ref", "--sort=-refname", "--format=%(refname:strip=3)%09%(objectname)%09%(subject)", "refs/codex/checkpoints/")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(line => {
                    var parts = line.Split(new[] { '\t' }, 3); return new GitCheckpoint { Id = parts[0], Commit = parts[1], Label = parts[2] };
                }).ToArray();
        }
        internal string CheckpointCommit(string id)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(id ?? "", @"^\d{17}-[0-9a-f]{8}$")) throw new ArgumentException("Identifiant de checkpoint invalide.");
            return Resolve("refs/codex/checkpoints/" + id) ?? throw new InvalidOperationException("Checkpoint absent.");
        }
        private string MergeFile { get { return Path.Combine(directory, "codex-merge.json"); } }
        internal GitMergePlan PendingMerge { get { return File.Exists(MergeFile) ? new JavaScriptSerializer().Deserialize<GitMergePlan>(File.ReadAllText(MergeFile)) : null; } }
        private void SaveMerge(GitMergePlan plan)
        {
            string temporary = MergeFile + ".pending";
            File.WriteAllText(temporary, new JavaScriptSerializer().Serialize(plan), VbaGitSnapshot.Utf8);
            if (File.Exists(MergeFile)) File.Replace(temporary, MergeFile, null); else File.Move(temporary, MergeFile);
            // Keep generated merge trees reachable during cache maintenance.
            SetRef("refs/codex/merge-tree", plan.Tree);
        }
        internal GitMergePlan BeginMerge(string branch)
        {
            if (PendingMerge != null) throw new InvalidOperationException("Une fusion est déjà en cours.");
            string ours = Resolve(Head) ?? throw new InvalidOperationException("Aucun commit local.");
            string theirs = BranchCommit(branch);
            var output = Run(new[] { "merge-tree", "--write-tree", "--name-only", "-z", ours, theirs }, null, true, true);
            if (output.ExitCode != 0 && output.ExitCode != 1) throw new InvalidOperationException("Fusion impossible : vérifiez les ancêtres communs et Git 2.38 ou ultérieur.");
            var records = Encoding.UTF8.GetString(output.Bytes).Split('\0');
            var plan = new GitMergePlan { Branch = Branch, Ours = ours, Theirs = theirs, Tree = records[0].Trim(),
                Conflicts = output.ExitCode == 0 ? new string[0] : records.Skip(1).TakeWhile(x => x.Length != 0).ToArray() };
            if (!System.Text.RegularExpressions.Regex.IsMatch(plan.Tree, "^[0-9a-f]{40}$")) throw new InvalidOperationException("Résultat de fusion Git invalide.");
            SaveMerge(plan); return plan;
        }
        internal void AssertMerge(GitMergePlan plan)
        {
            if (plan == null || plan.Branch != Branch || plan.Ours != Resolve(Head)) throw new InvalidOperationException("La branche a changé depuis le début de la fusion.");
        }
        internal GitConflictContent ConflictContent(string path)
        {
            var plan = PendingMerge; AssertMerge(plan);
            if (!plan.Conflicts.Contains(path)) throw new ArgumentException("Chemin absent des conflits.");
            return new GitConflictContent { Path = path, Ours = ConflictText(plan.Ours, path), Theirs = ConflictText(plan.Theirs, path) };
        }
        private string ConflictText(string tree, string path)
        {
            string entry = TreeAtPath(tree, path);
            if (entry == null) return "";
            var parts = entry.Split(' ', '\t');
            if (parts[1] != "blob" || path.EndsWith(".frx", StringComparison.OrdinalIgnoreCase)) return "[Ressource binaire ou dossier : choisir local ou entrant]";
            if (long.Parse(Text("cat-file", "-s", parts[2])) > 65536) return "[Aperçu indisponible au-delà de 64 Kio : choisir local ou entrant]";
            try { return VbaGitSnapshot.Utf8.GetString(Run(new[] { "cat-file", "blob", parts[2] }).Bytes); }
            catch (DecoderFallbackException) { return "[Contenu non UTF-8 : choisir local ou entrant]"; }
        }
        internal GitMergePlan ResolveConflict(string path, string choice, string text = null)
        {
            var plan = PendingMerge; AssertMerge(plan);
            if (!plan.Conflicts.Contains(path)) throw new InvalidOperationException("Ce fichier ne figure pas dans les conflits.");
            string entry;
            if (choice == "ours" || choice == "theirs")
            {
                var selected = TreeAtPath(choice == "ours" ? plan.Ours : plan.Theirs, path);
                entry = selected;
            }
            else if (choice == "text" && path.StartsWith("vba/", StringComparison.Ordinal) && !path.EndsWith(".frx", StringComparison.Ordinal) && text != null && text.Length < 1024 * 1024)
                entry = "100644 blob " + Encoding.UTF8.GetString(Run(new[] { "hash-object", "-w", "--stdin" }, VbaGitSnapshot.Utf8.GetBytes(text)).Bytes).Trim() + "\t" + path.Split('/').Last();
            else throw new ArgumentException("Choisissez ours, theirs ou text pour une source VBA textuelle.");
            plan.Tree = ReplacePath(plan.Tree, path.Split('/'), 0, entry);
            plan.Conflicts = plan.Conflicts.Where(x => x != path).ToArray(); SaveMerge(plan); return plan;
        }
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
        private string ReplacePath(string tree, string[] path, int index, string replacement)
        {
            var entries = tree == null ? new List<string>() : Tree(tree);
            string old = entries.FirstOrDefault(x => x.EndsWith("\t" + path[index], StringComparison.Ordinal));
            if (old != null) entries.Remove(old);
            if (index == path.Length - 1) { if (replacement != null) entries.Add(replacement); }
            else
            {
                if (old != null && !old.StartsWith("040000 tree ", StringComparison.Ordinal)) throw new InvalidOperationException("Conflit fichier/dossier : choisissez d’abord le dossier parent.");
                string child = ReplacePath(old?.Split(' ', '\t')[2], path, index + 1, replacement);
                entries.Add("040000 tree " + child + "\t" + path[index]);
            }
            return entries.Count == 0 ? Encoding.UTF8.GetString(Run(new[] { "mktree" }, new byte[0]).Bytes).Trim() : MakeTree(entries);
        }
        internal string MergeCommit(GitMergePlan plan, string message)
        {
            if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("Un message de commit de fusion est requis.");
            AssertMerge(plan);
            if (plan.Conflicts.Length != 0) throw new InvalidOperationException("Résolvez tous les conflits avant de terminer la fusion.");
            Read(plan.Tree); // Validate the complete VBA package before creating the merge commit.
            return Encoding.UTF8.GetString(Run(new[] { "commit-tree", plan.Tree, "-p", plan.Ours, "-p", plan.Theirs }, VbaGitSnapshot.Utf8.GetBytes(message + "\n")).Bytes).Trim();
        }
        internal void AbortMerge() { if (File.Exists(MergeFile)) File.Delete(MergeFile); }
    }
}
