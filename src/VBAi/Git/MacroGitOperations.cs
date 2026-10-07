using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Coordonne les opérations Git avec l’état courant du projet VBA et son dépôt local.</summary>
    internal sealed class MacroGitOperations : IDisposable
    {

        /// <summary>Dépôt Git associé à la portée du document.</summary>
        internal readonly MacroGitRepository Repository;

        /// <summary>Accès au projet VBE utilisé pour capturer et appliquer les sources.</summary>
        private readonly VbaGitProject project;

        /// <summary>Verrou exclusif conservé pendant la session d’opérations.</summary>
        private FileStream cacheLock;

        /// <summary>Callback facultatif appelé avant l’import pour présenter son résumé.</summary>
        internal Action<string> ImportPreview;

        /// <summary>Optional read-only modal-owner admission immediately before the existing native mutation boundary.</summary>
        internal Action ImportOwnerPreflight;

        /// <summary>Optional read-only owner validator after a native persistence continuation.</summary>
        internal Action ImportOwnerReadback;

        /// <summary>Optional qualification-only guard on the commit returned by the one production pull fetch.</summary>
        internal string ExpectedIncomingCommit;

        /// <summary>Résout le répertoire de cache de la portée, notamment pour isoler les tests.</summary>
        internal static Func<string, string> CacheDirectory = MacroGitRepository.ResolveScopeDirectory;

        /// <summary>Crée un coordinateur pour le projet et le dépôt fournis.</summary>
        /// <param name="project">Projet VBA à lire ou modifier.</param>
        /// <param name="repository">Dépôt associé à ce projet.</param>
        internal MacroGitOperations(VbaGitProject project, MacroGitRepository repository) { this.project = project; Repository = repository; }

        /// <summary>Ouvre le dépôt lié à la portée et conserve son verrou de session.</summary>
        /// <param name="project">Projet VBA concerné.</param>
        /// <param name="scope">Identifiant de portée du document.</param>
        /// <param name="account">Compte GitHub à utiliser pour les opérations distantes.</param>
        /// <returns>Coordinateur prêt à exécuter des opérations sur le dépôt lié.</returns>
        /// <exception cref="InvalidOperationException">Aucun lien de dépôt valide n’existe pour cette portée.</exception>
        /// <exception cref="IOException">Le verrou de session ne peut pas être acquis.</exception>
        internal static MacroGitOperations Open(VbaGitProject project, string scope, string account)
        {
            string cache = CacheDirectory(scope);
            if (!File.Exists(Path.Combine(cache, "binding.json"))) throw new InvalidOperationException(UiText.Get("Link this document to GitHub through the interface first. The agent does not choose a repository for you."));
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

        /// <summary>Libère le verrou de session détenu par ce coordinateur.</summary>
        public void Dispose() { cacheLock?.Dispose(); }

        /// <summary>Calcule une empreinte combinant les sources capturées et les références Git en cours.</summary>
        /// <param name="snapshot">État des sources du projet utilisé pour calculer la révision.</param>
        /// <returns>Empreinte SHA-256 hexadécimale de l’état combiné.</returns>
        internal string Revision(VbaGitSnapshot snapshot) { return Revision(snapshot, ReadRecoveryState()); }

        /// <summary>Hashes the exact recovery identity together with the reviewed VBA and Git state.</summary>
        private string Revision(VbaGitSnapshot snapshot, RecoveryState recovery)
        {
            var observer = ObserveRevision;
            List<RevisionFileDiagnostic> observedFiles = null;
            bool observationFailed = false;
            int observedFileCount = 0;
            if (observer != null)
            {
                try { observedFiles = new List<RevisionFileDiagnostic>(); }
                catch { observationFailed = true; }
            }
            using (var hash = SHA256.Create())
            {
                foreach (var file in snapshot.ComparisonFiles())
                {
                    byte[] key = Encoding.UTF8.GetBytes(file.Key + "\0" + file.Value.Length + "\0");
                    hash.TransformBlock(key, 0, key.Length, null, 0);
                    hash.TransformBlock(file.Value, 0, file.Value.Length, null, 0);
                    if (observer != null)
                    {
                        observedFileCount++;
                        if (!observationFailed && observedFiles.Count < RevisionDiagnostic.MaximumFiles)
                        {
                            try { observedFiles.Add(new RevisionFileDiagnostic(key, file.Value)); }
                            catch { observationFailed = true; }
                        }
                    }
                }
                // Keep the original operand evaluation order and observe these exact values only.
                string branch = Repository.Branch;
                string head = Repository.Resolve(Repository.Head);
                string baseline = Repository.Resolve(MacroGitRepository.Baseline);
                string merge = new JavaScriptSerializer().Serialize(Repository.PendingMerge);
                byte[] state = Encoding.UTF8.GetBytes(branch + "\0" + head + "\0" + baseline + "\0" + merge + "\0" +
                    recovery.Backup + "\0" + recovery.After + "\0" + (recovery.Marker == null ? "absent" : "present:" + recovery.Marker));
                hash.TransformFinalBlock(state, 0, state.Length);
                string revision = BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant();
                if (observer != null && !observationFailed)
                {
                    try
                    {
                        observer(new RevisionDiagnostic(observedFiles.ToArray(), observedFileCount, branch, head, baseline,
                            merge, recovery.Backup, recovery.After, recovery.Marker, state, revision));
                    }
                    catch { /* Diagnostic allocation or callbacks must never change the admission verdict. */ }
                }
                return revision;
            }
        }

        /// <summary>Optional instance-only diagnostics; null performs no diagnostic hashing, allocation, I/O or environment lookup.</summary>
        internal Action<RevisionDiagnostic> ObserveRevision;

        /// <summary>Bounded hashes of operands actually used by one revision; never contains source or raw arguments.</summary>
        internal sealed class RevisionDiagnostic
        {
            internal const int MaximumFiles = 128;
            public readonly RevisionFileDiagnostic[] Files;
            public readonly int TotalFiles;
            public readonly bool FilesTruncated;
            public readonly RevisionValueDiagnostic Branch, Head, Baseline, PendingMerge, Backup, After, Marker;
            public readonly string StateSha256, Revision;
            internal RevisionDiagnostic(RevisionFileDiagnostic[] files, int total, string branch, string head,
                string baseline, string merge, string backup, string after, string marker, byte[] state, string revision)
            {
                Files = files; TotalFiles = total; FilesTruncated = total > files.Length;
                Branch = new RevisionValueDiagnostic(branch); Head = new RevisionValueDiagnostic(head);
                Baseline = new RevisionValueDiagnostic(baseline); PendingMerge = new RevisionValueDiagnostic(merge);
                Backup = new RevisionValueDiagnostic(backup); After = new RevisionValueDiagnostic(after);
                Marker = new RevisionValueDiagnostic(marker); StateSha256 = DiagnosticSha(state); Revision = revision;
            }
        }

        /// <summary>Hashes the exact length-prefixed file key and content already passed to the revision hash.</summary>
        internal sealed class RevisionFileDiagnostic
        {
            public readonly string KeySha256, ContentSha256;
            public readonly int KeyBytes, ContentBytes;
            internal RevisionFileDiagnostic(byte[] key, byte[] content)
            { KeySha256 = DiagnosticSha(key); ContentSha256 = DiagnosticSha(content); KeyBytes = key.Length; ContentBytes = content.Length; }
        }

        /// <summary>Separates null from empty without retaining an operand's text.</summary>
        internal sealed class RevisionValueDiagnostic
        {
            public readonly string Sha256;
            public readonly int Bytes;
            public readonly bool IsNull;
            internal RevisionValueDiagnostic(string value)
            {
                IsNull = value == null;
                byte[] bytes = Encoding.UTF8.GetBytes(value ?? "");
                Bytes = bytes.Length; Sha256 = DiagnosticSha(bytes);
            }
        }

        /// <summary>Called only while opt-in diagnostic evidence is collected.</summary>
        private static string DiagnosticSha(byte[] bytes)
        {
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        /// <summary>Immutable commit and marker identities used throughout one import or rollback.</summary>
        private sealed class RecoveryState
        {
            /// <summary>Exact backup and measured-after commits, plus the bounded marker content or null for absence.</summary>
            internal readonly string Backup, After, Marker;
            /// <summary>Freezes recovery authority before any reviewed rollback target is resolved.</summary>
            internal RecoveryState(string backup, string after, string marker) { Backup = backup; After = after; Marker = marker; }
            /// <summary>Compares every authority-bearing recovery value without resolving another target.</summary>
            internal bool SameAs(RecoveryState other) { return Backup == other.Backup && After == other.After && Marker == other.Marker; }
        }

        /// <summary>Resolves recovery references once, before a rollback target is read by immutable commit ID.</summary>
        private RecoveryState ReadRecoveryState()
        {
            string marker = Repository.RecoveryMarkerIdentity();
            return new RecoveryState(Repository.Resolve(MacroGitRepository.Backup), Repository.Resolve(MacroGitRepository.AfterImport), marker);
        }

        /// <summary>Refuses changed recovery authority without replacing the frozen target or replaying a mutation.</summary>
        private void RequireRecoveryState(RecoveryState expected)
        {
            if (!expected.SameAs(ReadRecoveryState()))
                throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
        }

        /// <summary>Capture le projet et renvoie son état, ses changements, branches, points de contrôle et fusion.</summary>
        /// <returns>Objet de statut destiné au protocole de l’agent.</returns>
        internal async Task<object> StatusAsync()
        {
            var live = project.Capture();
            return await Task.Run<object>(() => new
            {
                State = Revision(live),
                Repository.Branch,
                Head = Repository.Resolve(Repository.Head),
                ChangedFiles = live.Changes(Repository.Read(Repository.Resolve(MacroGitRepository.Baseline))),
                Branches = Repository.Branches(),
                Checkpoints = Repository.Checkpoints(),
                History = Repository.History(),
                Merge = Repository.PendingMerge,
                Repository.RecoveryPending,
                Synchronization = Repository.SynchronizationStatus()
            });
        }

        /// <summary>Récupère le contenu des deux côtés d’un conflit pour le chemin indiqué.</summary>
        /// <param name="path">Chemin du fichier en conflit.</param>
        /// <returns>Tâche produisant le contenu du conflit.</returns>
        internal Task<GitConflictContent> ConflictAsync(string path) { return Task.Run(() => Repository.ConflictContent(path)); }

        /// <summary>Refuse les opérations incompatibles avec une récupération ou fusion en attente.</summary>
        /// <param name="allowMerge"><see langword="true"/> pour autoriser les actions de progression de fusion.</param>
        /// <exception cref="InvalidOperationException">Une récupération interrompue doit être restaurée ou une fusion doit être terminée.</exception>
        private void Ready(bool allowMerge = false)
        {
            if (Repository.RecoveryPending) throw new InvalidOperationException(UiText.Get("Restore the interrupted import before continuing."));
            if (!allowMerge && Repository.PendingMerge != null) throw new InvalidOperationException(UiText.Get("Complete or abort the current merge."));
        }

        /// <summary>Vérifie que les sources locales correspondent à la référence de base du dépôt.</summary>
        /// <param name="live">Instantané courant du projet.</param>
        /// <exception cref="InvalidOperationException">Le dépôt n’a pas de base ou des changements locaux sont présents.</exception>
        private void Clean(VbaGitSnapshot live)
        {
            var baseline = Repository.Read(Repository.Resolve(MacroGitRepository.Baseline));
            if (baseline == null || !live.SameAs(baseline)) throw new InvalidOperationException(UiText.Get("Local changes exist. Commit them, or create a checkpoint and restore a committed state before switching branches or merging."));
        }

        /// <summary>Exécute une action Git après vérification de l’état attendu du projet.</summary>
        /// <param name="action">Identifiant de l’action à exécuter.</param>
        /// <param name="expectedState">Empreinte d’état attendue, ou <see langword="null"/> si aucune n’est fournie.</param>
        /// <param name="name">Nom de branche, point de contrôle ou message de l’action.</param>
        /// <param name="text">Texte complémentaire, tel qu’un message de commit ou une résolution.</param>
        /// <param name="choice">Choix de résolution fourni pour une action de fusion.</param>
        /// <param name="path">Chemin du fichier concerné, le cas échéant.</param>
        /// <param name="modules">Modules sélectionnés pour un commit partiel.</param>
        /// <param name="references"><see langword="true"/> pour inclure les références du projet dans un commit partiel.</param>
        /// <returns>Statut actualisé du dépôt ou résultat spécifique à l’action.</returns>
        /// <exception cref="ArgumentException">L’action ou ses arguments obligatoires sont invalides.</exception>
        /// <exception cref="InvalidOperationException">L’état a changé ou les préconditions de l’action ne sont pas satisfaites.</exception>
        internal async Task<object> ExecuteAsync(string action, string expectedState = null, string name = null, string text = null, string choice = null, string path = null, string[] modules = null, bool references = false)
        {
            var live = project.Capture();
            RecoveryState recovery = expectedState != null || action == "rollback" ? await Task.Run(() => ReadRecoveryState()) : null;
            if (expectedState != null && expectedState != await Task.Run(() => Revision(live, recovery))) throw new InvalidOperationException(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
            if (action != "rollback") Ready(action.StartsWith("merge_", StringComparison.Ordinal));
            // Rollback's frozen recovery read has already validated the marker type and contents.
            switch (action)
            {
                case "pr_prepare": await Task.Run(() => Repository.SavePullDraft(name, text, choice)); break;
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
                case "module_restore":
                    var revision = await Task.Run(() => Repository.Read(Repository.VerifiedCommit(name)));
                    await ImportAsync(VbaGitSnapshot.Select(live, revision, new[] { path }), live);
                    break;
                case "commit_selected":
                case "commit":
                    string commit = await Task.Run(() =>
                    {
                        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException(UiText.Get("A commit message is required."));
                        string parent = Repository.Resolve(Repository.Head) ?? Repository.Resolve("refs/remotes/origin/selected");
                        var previous = Repository.Read(parent);
                        if (Repository.Resolve(Repository.Head) == null && previous != null && !live.SameAs(previous)) throw new InvalidOperationException(UiText.Get("Import the remote state with Pull first."));
                        var staged = action == "commit_selected" ? VbaGitSnapshot.Select(previous, live, modules, references) : live;
                        if (action == "commit_selected" && (modules == null || modules.Length == 0) && !references)
                            throw new InvalidOperationException(UiText.Get("Select at least one module or the references."));
                        string next = staged.SameAs(previous) ? parent : Repository.Commit(staged, parent, text);
                        Repository.SetRef(Repository.Head, next); Repository.SetRef(MacroGitRepository.Baseline, next); return next;
                    }); return new { Commit = commit, Published = false };
                case "fetch": await Task.Run(() => Repository.Fetch()); break;
                case "push":
                    await Task.Run(() =>
                    {
                        string local = Repository.Resolve(Repository.Head) ?? throw new InvalidOperationException(UiText.Get("No commit to publish."));
                        string remote = Repository.Fetch(); if (remote != null) Repository.RequireFastForward(remote, local);
                        Repository.Push(local);
                    }); break;
                case "pull":
                    string incoming = await Task.Run(() =>
                    {
                        var previous = Repository.Read(Repository.Resolve(MacroGitRepository.Baseline));
                        if (previous != null && !live.SameAs(previous)) throw new InvalidOperationException(UiText.Get("VBA contains uncommitted local changes."));
                        string remote = Repository.Fetch() ?? throw new InvalidOperationException("Branche distante absente.");
                        if (ExpectedIncomingCommit != null && !string.Equals(remote, ExpectedIncomingCommit, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("The frozen incoming Git commit changed before import.");
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
                    var rollback = await Task.Run(() =>
                    {
                        var after = Repository.Read(recovery.After);
                        if (after == null || !live.SameAs(after)) throw new InvalidOperationException(UiText.Get("Code changed since the import; automatic restore refused."));
                        return Repository.Read(recovery.Backup) ?? throw new InvalidOperationException("Aucune sauvegarde.");
                    });
                    await ImportPreparedAsync(rollback, live, true, recovery); break;
                default: throw new ArgumentException("Action Git inconnue.");
            }
            return await StatusAsync();
        }

        /// <summary>Applique un instantané au projet après sauvegarde et vérifie l’état importé.</summary>
        /// <param name="target">Instantané à importer.</param>
        /// <param name="expected">Instantané courant attendu avant mutation.</param>
        /// <param name="rollback"><see langword="true"/> si l’import restaure une récupération préparée.</param>
        /// <returns>Tâche terminée après l’application et l’enregistrement de l’état relu.</returns>
        /// <exception cref="InvalidOperationException">La cible est vide ou le projet ne correspond plus à l’état attendu.</exception>
        private Task ImportAsync(VbaGitSnapshot target, VbaGitSnapshot expected, bool rollback = false)
        { return ImportPreparedAsync(target, expected, rollback, rollback ? ReadRecoveryState() : null); }

        /// <summary>Imports one frozen target, retaining the exact recovery authority through admission and completion.</summary>
        private async Task ImportPreparedAsync(VbaGitSnapshot target, VbaGitSnapshot expected, bool rollback, RecoveryState recovery)
        {
            if (target == null) throw new InvalidOperationException(UiText.Get("The target contains no VBA sources."));
            if (!project.Capture().SameAs(expected)) throw new InvalidOperationException(UiText.Get("VBA changed during the operation."));
            if (rollback) RequireRecoveryState(recovery);
            if (expected.SameAs(target)) { if (rollback) { RequireRecoveryState(recovery); Repository.CompleteRecovery(); } return; }
            ImportPreview?.Invoke(target.ImportSummary(expected));
            if (!rollback) recovery = await Task.Run(() =>
            {
                Repository.Checkpoint(expected, UiText.Get("Before import · ") + DateTime.Now.ToString("s"));
                Repository.PrepareRecovery(expected);
                // Freeze the newly prepared state, after PrepareRecovery has cleared
                // any previous after-state and written this import's exact marker.
                return ReadRecoveryState();
            });
            else
            {
                RequireRecoveryState(recovery);
                File.WriteAllText(Repository.RecoveryFile, recovery.Backup, Encoding.ASCII);
                recovery = new RecoveryState(recovery.Backup, recovery.After, Convert.ToBase64String(Encoding.ASCII.GetBytes(recovery.Backup)));
            }
            bool started = false;
            Exception importFailure = null;
            try
            {
                await project.ApplyAsync(target, expected, () =>
                {
                    RequireRecoveryState(recovery);
                    ImportOwnerPreflight?.Invoke();
                    RequireRecoveryState(recovery);
                    started = true;
                }, RecoveryReadbackGuard(recovery));
            }
            catch (Exception error) { importFailure = error; throw; }
            finally
            {
                if (started)
                {
                    try
                    {
                        var actual = project.Capture();
                        string observedCommit = await Task.Run(() =>
                        {
                            RequireRecoveryState(recovery);
                            return Repository.RecordImportedState(actual);
                        });
                        recovery = new RecoveryState(recovery.Backup, observedCommit, recovery.Marker);
                    }
                    catch (Exception recoveryFailure) when (importFailure != null)
                    {
                        // Keep both failures and the pending recovery backup. Without
                        // an observed post-import state, automatic rollback stays refused.
                        throw new AggregateException(UiText.Get("VBA import failed and its resulting state could not be recorded. Recovery remains pending; inspect the retained backup and live project before restoring."),
                            importFailure, recoveryFailure);
                    }
                }
                else if (!rollback)
                {
                    try { RequireRecoveryState(recovery); Repository.CompleteRecovery(); }
                    catch (Exception recoveryFailure) when (importFailure != null)
                    {
                        throw new AggregateException("VBA import was refused before its mutation boundary and recovery completion also failed; both errors are retained.",
                            importFailure, recoveryFailure);
                    }
                }
            }
            RequireRecoveryState(recovery);
            Repository.CompleteRecovery();
        }

        /// <summary>Preserves frozen recovery authority around a read-only owner continuation guard.</summary>
        private Action RecoveryReadbackGuard(RecoveryState recovery)
        {
            return () =>
            {
                RequireRecoveryState(recovery);
                ImportOwnerReadback?.Invoke();
                RequireRecoveryState(recovery);
            };
        }
    }
}
