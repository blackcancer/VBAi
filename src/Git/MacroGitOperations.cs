using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class MacroGitOperations : IDisposable
    {
        internal readonly MacroGitRepository Repository;
        private readonly VbaGitProject project;
        private FileStream cacheLock;
        internal MacroGitOperations(VbaGitProject project, MacroGitRepository repository) { this.project = project; Repository = repository; }

        internal static MacroGitOperations Open(VbaGitProject project, string scope, string account)
        {
            string cache = MacroGitRepository.ScopeDirectory(scope);
            if (!File.Exists(Path.Combine(cache, "binding.json"))) throw new InvalidOperationException("Liez d’abord ce document à GitHub depuis l’interface. L’agent ne choisit pas de dépôt à votre place.");
            var held = new FileStream(Path.Combine(cache, "session.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            try
            {
                var binding = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(cache, "binding.json")));
                string url = MacroGitRepository.ValidateRemote(binding["Remote"]), branch = binding["Branch"];
                string id;
                using (var sha = SHA256.Create()) id = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(url + "\n" + branch))).Replace("-", "");
                var repository = new MacroGitRepository(Path.Combine(cache, id + ".git"), branch, account);
                repository.Initialize(url);
                return new MacroGitOperations(project, repository) { cacheLock = held };
            }
            catch { held.Dispose(); throw; }
        }
        public void Dispose() { cacheLock?.Dispose(); }

        internal string Revision(VbaGitSnapshot snapshot)
        {
            using (var hash = SHA256.Create())
            {
                foreach (var file in snapshot.Serialize())
                {
                    byte[] key = Encoding.UTF8.GetBytes(file.Key + "\0" + file.Value.Length + "\0");
                    hash.TransformBlock(key, 0, key.Length, null, 0);
                    hash.TransformBlock(file.Value, 0, file.Value.Length, null, 0);
                }
                byte[] state = Encoding.UTF8.GetBytes(Repository.Branch + "\0" + Repository.Resolve(Repository.Head) + "\0" +
                    Repository.Resolve(MacroGitRepository.Baseline) + "\0" + new JavaScriptSerializer().Serialize(Repository.PendingMerge) + "\0" + Repository.RecoveryPending);
                hash.TransformFinalBlock(state, 0, state.Length);
                return BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant();
            }
        }
        internal async Task<object> StatusAsync()
        {
            var live = project.Capture();
            return await Task.Run<object>(() => new {
                State = Revision(live), Branch = Repository.Branch, Head = Repository.Resolve(Repository.Head),
                ChangedFiles = live.Changes(Repository.Read(Repository.Resolve(MacroGitRepository.Baseline))),
                Branches = Repository.Branches(), Checkpoints = Repository.Checkpoints(), History = Repository.History(),
                Merge = Repository.PendingMerge, RecoveryPending = Repository.RecoveryPending,
                Synchronization = Repository.SynchronizationStatus()
            });
        }
        internal Task<GitConflictContent> ConflictAsync(string path) { return Task.Run(() => Repository.ConflictContent(path)); }
        private void Ready(bool allowMerge = false)
        {
            if (Repository.RecoveryPending) throw new InvalidOperationException("Restaurez l’import interrompu avant de poursuivre.");
            if (!allowMerge && Repository.PendingMerge != null) throw new InvalidOperationException("Terminez ou annulez la fusion en cours.");
        }
        private void Clean(VbaGitSnapshot live)
        {
            var baseline = Repository.Read(Repository.Resolve(MacroGitRepository.Baseline));
            if (baseline == null || !live.SameAs(baseline)) throw new InvalidOperationException("Des modifications locales existent. Commitez-les ou créez un checkpoint puis restaurez un état commité avant de changer de branche/fusionner.");
        }
        internal async Task<object> ExecuteAsync(string action, string expectedState = null, string name = null, string text = null, string choice = null, string path = null)
        {
            var live = project.Capture();
            if (expectedState != null && expectedState != await Task.Run(() => Revision(live))) throw new InvalidOperationException("L’état Git/VBA a changé. Relisez git_status avant de modifier.");
            if (action != "rollback") Ready(action.StartsWith("merge_", StringComparison.Ordinal));
            switch (action)
            {
                case "checkpoint_create": return await Task.Run(() => Repository.Checkpoint(live, name));
                case "checkpoint_restore":
                    var saved = await Task.Run(() => Repository.Read(Repository.CheckpointCommit(name)));
                    await ImportAsync(saved, live); return await StatusAsync();
                case "branch_create": await Task.Run(() => Repository.CreateBranch(name)); break;
                case "branch_track": await Task.Run(() => Repository.TrackBranch(name)); break;
                case "remote_branches": return await Task.Run(() => Repository.RemoteBranches());
                case "branch_switch":
                    string branchCommit = await Task.Run(() => { Clean(live); return Repository.BranchCommit(name); });
                    var target = await Task.Run(() => Repository.Read(branchCommit));
                    await ImportAsync(target, live);
                    await Task.Run(() => { Repository.SelectBranch(name); Repository.SetRef(MacroGitRepository.Baseline, branchCommit); }); break;
                case "commit":
                    string commit = await Task.Run(() => {
                        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Un message de commit est requis.");
                        string parent = Repository.Resolve(Repository.Head) ?? Repository.Resolve("refs/remotes/origin/selected");
                        var previous = Repository.Read(parent);
                        if (Repository.Resolve(Repository.Head) == null && previous != null && !live.SameAs(previous)) throw new InvalidOperationException("Importez d’abord l’état distant avec Pull.");
                        string next = live.SameAs(previous) ? parent : Repository.Commit(live, parent, text);
                        Repository.SetRef(Repository.Head, next); Repository.SetRef(MacroGitRepository.Baseline, next); return next;
                    }); return new { Commit = commit, Published = false };
                case "fetch": await Task.Run(() => Repository.Fetch()); break;
                case "push":
                    await Task.Run(() => {
                        string local = Repository.Resolve(Repository.Head) ?? throw new InvalidOperationException("Aucun commit à publier.");
                        string remote = Repository.Fetch(); if (remote != null) Repository.RequireFastForward(remote, local);
                        Repository.Push(local);
                    }); break;
                case "pull":
                    string incoming = await Task.Run(() => {
                        var previous = Repository.Read(Repository.Resolve(MacroGitRepository.Baseline));
                        if (previous != null && !live.SameAs(previous)) throw new InvalidOperationException("Le VBA contient des modifications locales non commitées.");
                        string remote = Repository.Fetch() ?? throw new InvalidOperationException("Branche distante absente.");
                        Repository.RequireFastForward(Repository.Resolve(Repository.Head), remote); return remote;
                    });
                    await ImportAsync(await Task.Run(() => Repository.Read(incoming)), live);
                    await Task.Run(() => { Repository.SetRef(Repository.Head, incoming); Repository.SetRef(MacroGitRepository.Baseline, incoming); }); break;
                case "merge_begin": return await Task.Run(() => { Clean(live); return Repository.BeginMerge(name); });
                case "merge_resolve": return await Task.Run(() => Repository.ResolveConflict(path, choice, text));
                case "merge_abort": await Task.Run(() => Repository.AbortMerge()); break;
                case "merge_complete":
                    string merged = await Task.Run(() => { Clean(live); return Repository.MergeCommit(Repository.PendingMerge, text ?? "Fusion VBA"); });
                    await ImportAsync(await Task.Run(() => Repository.Read(merged)), live);
                    await Task.Run(() => { Repository.SetRef(Repository.Head, merged); Repository.SetRef(MacroGitRepository.Baseline, merged); Repository.AbortMerge(); }); break;
                case "rollback":
                    var rollback = await Task.Run(() => {
                        var after = Repository.Read(Repository.Resolve(MacroGitRepository.AfterImport));
                        if (after == null || !live.SameAs(after)) throw new InvalidOperationException("Le code a changé depuis l’import ; restauration automatique refusée.");
                        return Repository.Read(Repository.Resolve(MacroGitRepository.Backup)) ?? throw new InvalidOperationException("Aucune sauvegarde.");
                    });
                    await ImportAsync(rollback, live, true); break;
                default: throw new ArgumentException("Action Git inconnue.");
            }
            return await StatusAsync();
        }
        private async Task ImportAsync(VbaGitSnapshot target, VbaGitSnapshot expected, bool rollback = false)
        {
            if (target == null) throw new InvalidOperationException("La cible ne contient pas de sources VBA.");
            if (!project.Capture().SameAs(expected)) throw new InvalidOperationException("Le VBA a changé pendant l’opération.");
            if (expected.SameAs(target)) { if (rollback) Repository.CompleteRecovery(); return; }
            if (!rollback) await Task.Run(() => { Repository.Checkpoint(expected, "Avant import · " + DateTime.Now.ToString("s")); Repository.PrepareRecovery(expected); });
            else File.WriteAllText(Repository.RecoveryFile, Repository.Resolve(MacroGitRepository.Backup));
            bool started = false;
            try { project.Apply(target, expected, () => started = true); }
            finally
            {
                if (started) { var actual = project.Capture(); await Task.Run(() => Repository.RecordImportedState(actual)); }
                else if (!rollback) Repository.CompleteRecovery();
            }
            Repository.CompleteRecovery();
        }
    }
}
