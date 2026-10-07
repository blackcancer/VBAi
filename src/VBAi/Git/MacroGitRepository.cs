using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBAi
{
    // Bare repository: no working tree, no checkout of remote files and no Git hooks.
    /// <summary>Gère un dépôt Git bare par document VBA sans arbre de travail.</summary>
    internal sealed partial class MacroGitRepository
    {

        /// <summary>Répertoire du dépôt bare utilisé comme cache du document.</summary>
        private readonly string directory;

        /// <summary>Compte GitHub configuré pour l’authentification distante.</summary>
        private readonly string account;

        /// <summary>Frontière d’exécution optionnelle des commandes ; null utilise le processus Git natif.</summary>
        internal Func<string[], byte[], bool, bool, Result> CommandOverride;

        /// <summary>Démarre le processus natif sans préambule sur son entrée standard.</summary>
        internal Func<Process, bool> StartProcess = ProcessInput.StartWithoutPreamble;

        /// <summary>Attend la fin du processus dans le délai fourni.</summary>
        internal Func<Process, int, bool> WaitForExit = (process, timeout) => process.WaitForExit(timeout);

        /// <summary>Interrompt le processus Git possédé par cette opération.</summary>
        internal Action<Process> StopProcess = process => process.Kill();

        /// <summary>Reads only the recovery marker metadata; failures remain observable rather than meaning absence.</summary>
        internal Func<string, FileAttributes> RecoveryAttributes = File.GetAttributes;

        /// <summary>Deletes the confirmed regular recovery marker once, without retry after an uncertain outcome.</summary>
        internal Action<string> DeleteRecoveryMarker = File.Delete;

        /// <summary>Branche locale active.</summary>
        /// <value>Branche locale active.</value>
        internal string Branch { get; private set; }

        /// <summary>Référence privée du dernier état servant de base à une synchronisation.</summary>
        internal const string Baseline = "refs/codex/baseline";

        /// <summary>Référence privée des sauvegardes avant import.</summary>
        internal const string Backup = "refs/codex/backup";

        /// <summary>Référence privée de l’état VBA après import.</summary>
        internal const string AfterImport = "refs/codex/after-import";

        /// <summary>Nom complet de la référence HEAD de la branche active.</summary>
        /// <value>Nom complet de la référence HEAD de la branche active.</value>
        internal string Head { get { return "refs/heads/" + Branch; } }

        /// <summary>Chemin du marqueur de récupération d’import.</summary>
        /// <value>Chemin du marqueur de récupération d’import.</value>
        internal string RecoveryFile { get { return Path.Combine(directory, "codex-recovery"); } }

        /// <summary>Indique si un marqueur de récupération est présent.</summary>
        /// <value>Indique si un marqueur de récupération est présent.</value>
        internal bool RecoveryPending { get { return ObserveRecoveryMarker().HasValue; } }

        /// <summary>Only a definite missing marker is absence; any access, I/O or other metadata error propagates.</summary>
        /// <returns>file attributes produced by the operation for observe recovery marker on macro git repository.</returns>
        private FileAttributes? ObserveRecoveryMarker()
        {
            try { return RecoveryAttributes(RecoveryFile); }
            catch (FileNotFoundException) { return null; }
            catch (DirectoryNotFoundException) { return null; }
        }

        /// <summary>Allows an absent or regular marker for rollback, never a directory or filesystem link.</summary>
        internal void RequireValidRecoveryMarker()
        {
            var attributes = ObserveRecoveryMarker();
            if (attributes.HasValue) RequireRegularRecoveryMarker(attributes.Value);
        }

        /// <summary>Rejects invalid marker types before any native import or marker deletion.</summary>
        /// <param name="attributes">Attributes read from the recovery marker before rollback or deletion.</param>
        private static void RequireRegularRecoveryMarker(FileAttributes attributes)
        {
            if ((attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
                throw new IOException("The recovery marker is not a regular file; recovery cannot continue automatically.");
        }

        /// <summary>Crée ou configure le dépôt bare pour la branche et le compte donnés.</summary>
        /// <param name="directory">Répertoire du dépôt bare.</param>
        /// <param name="branch">Nom de la branche source ou cible.</param>
        /// <param name="account">Compte GitHub facultatif pour les identifiants.</param>
        internal MacroGitRepository(string directory, string branch, string account = null)
        {
            if (!Regex.IsMatch(branch ?? "", @"^[A-Za-z0-9][A-Za-z0-9_./-]{0,127}$") || branch.Contains("..") ||
                branch.Contains("//") || branch.EndsWith("/") || branch.Split('/').Any(x => x.StartsWith(".") || x.EndsWith(".") || x.EndsWith(".lock")))
                throw new ArgumentException("Nom de branche Git invalide.");
            if (!string.IsNullOrEmpty(account) && !GitHubAccountService.ValidAccount(account)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            this.account = account;
            this.directory = Path.GetFullPath(directory); Branch = branch;
        }

        /// <summary>Calcule le répertoire local isolé pour la portée fournie.</summary>
        /// <param name="scope">Portée du document hôte servant à isoler son cache.</param>
        /// <returns>Chemin du répertoire isolé.</returns>
        internal static string ScopeDirectory(string scope)
        {
            using (var sha = SHA256.Create())
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "Git",
                    BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(scope))).Replace("-", "").ToLowerInvariant());
        }

        /// <summary>Valide et normalise une URL HTTPS GitHub sans identifiants.</summary>
        /// <param name="remote">URL HTTPS du dépôt GitHub.</param>
        /// <returns>URL HTTPS normalisée.</returns>
        internal static string ValidateRemote(string remote)
        {
            remote = (remote ?? "").Trim();
            if (!Regex.IsMatch(remote, @"^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+(?:\.git)?/?$", RegexOptions.IgnoreCase))
                throw new ArgumentException(UiText.Get("Enter the repository's HTTPS GitHub URL without a token or password."));
            return remote.TrimEnd('/');
        }

        /// <summary>Initialise le dépôt bare et son remote, ou vérifie leur cohérence.</summary>
        /// <param name="remote">URL HTTPS du dépôt GitHub.</param>
        internal void Initialize(string remote)
        {
            Directory.CreateDirectory(directory);
            if (!File.Exists(Path.Combine(directory, "HEAD")))
            {
                Run(new[] { "init", "--bare", directory }, null, false);
                Run(new[] { "remote", "add", "origin", remote });
            }
            if (Text("remote", "get-url", "origin") != remote)
                throw new InvalidOperationException(UiText.Get("This cache is already linked to another repository. Restore the previous URL."));
            var active = Run(new[] { "config", "--get", "codex.activeBranch" }, null, true, true);
            if (active.ExitCode == 0) { string name = Encoding.UTF8.GetString(active.Bytes).Trim(); ValidateBranch(name); Branch = name; }
        }

        /// <summary>Résout une référence Git ou retourne null si elle est absente.</summary>
        /// <param name="reference">Référence Git à résoudre ou mettre à jour.</param>
        /// <returns>Identifiant résolu ou null.</returns>
        internal string Resolve(string reference)
        {
            var result = Run(new[] { "rev-parse", "--verify", "--quiet", reference }, null, true, true);
            return result.ExitCode == 0 ? Encoding.UTF8.GetString(result.Bytes).Trim() : null;
        }

        /// <summary>Récupère la branche active distante dans la référence de suivi privée.</summary>
        /// <returns>Commit distant suivi, ou null si la branche est absente.</returns>
        internal string Fetch()
        {
            string target = "refs/heads/" + Branch;
            string listing = Text("ls-remote", "--heads", "origin", target);
            if (listing.Length == 0)
            {
                string previous = Resolve("refs/remotes/origin/selected");
                if (previous != null) Text("update-ref", "-d", "refs/remotes/origin/selected", previous);
                return null;
            }
            Text("fetch", "--no-tags", "origin", "+" + target + ":refs/remotes/origin/selected");
            return Resolve("refs/remotes/origin/selected");
        }

        /// <summary>Retourne jusqu’aux quarante derniers commits de la branche active.</summary>
        /// <returns>Lignes de résumé des commits récents.</returns>
        internal string[] History()
        {
            string head = Resolve(Head);
            return head == null ? new string[0] : Text("log", "-40", "--date=short", "--format=%h  %ad  %s", head)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        /// <summary>Résume l’écart entre les commits locaux et distants connus.</summary>
        /// <returns>Texte lisible du nombre de commits entrants et sortants.</returns>
        internal string SynchronizationStatus()
        {
            string head = Resolve(Head), remote = Resolve("refs/remotes/origin/selected");
            if (head == null) return remote == null ? UiText.Get("No local commit · use Fetch to inspect the remote repository") : "Branche distante disponible · premier pull requis";
            if (remote == null) return UiText.Get("Local branch · remote state unknown (Fetch)");
            var counts = Text("rev-list", "--left-right", "--count", head + "..." + remote).Split(new[] { '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return "↑ " + counts[0] + " sortant(s)   ↓ " + counts[1] + " entrant(s) · dernier Fetch";
        }

        /// <summary>Refuse une synchronisation qui n’est pas une avance rapide.</summary>
        /// <param name="before">Commit local avant synchronisation.</param>
        /// <param name="after">Commit distant proposé.</param>
        internal void RequireFastForward(string before, string after)
        {
            if (before == null || before == after) return;
            if (after == null || Run(new[] { "merge-base", "--is-ancestor", before, after }, null, true, true).ExitCode != 0)
                throw new InvalidOperationException(UiText.Get("Histories diverge. Reconcile the branch before synchronizing; no forced push is performed."));
        }

        /// <summary>Lit et valide le paquet VBA présent dans un commit.</summary>
        /// <param name="commit">Commit dont le paquet VBA est lu ou poussé.</param>
        /// <returns>Snapshot validé, ou null si le commit n’a pas de répertoire vba.</returns>
        internal VbaGitSnapshot Read(string commit)
        {
            if (commit == null) return null;
            var entries = Tree(commit);
            var folder = entries.FirstOrDefault(x => x.EndsWith("\tvba", StringComparison.Ordinal));
            if (folder == null) return null;
            if (!folder.StartsWith("040000 tree ", StringComparison.Ordinal)) throw new InvalidOperationException(UiText.Get("vba must be a Git directory."));
            string tree = folder.Split(' ', '\t')[2];
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            long total = 0;
            foreach (string entry in Tree(tree))
            {
                int tab = entry.IndexOf('\t');
                var parts = entry.Substring(0, tab).Split(' ');
                string name = entry.Substring(tab + 1);
                if (parts[0] != "100644" || parts[1] != "blob" || name.IndexOfAny(new[] { '/', '\\', ':', '\0' }) >= 0)
                    throw new InvalidOperationException(UiText.Get("The vba directory contains a link, subdirectory or unsupported file."));
                long size = long.Parse(Text("cat-file", "-s", parts[2]), System.Globalization.CultureInfo.InvariantCulture);
                total += size;
                if (total > VbaGitSnapshot.MaxBytes || files.Count > 2048) throw new InvalidOperationException(UiText.Get("The VBA repository exceeds 32 MB."));
                files.Add(name, Run(new[] { "cat-file", "blob", parts[2] }).Bytes);
            }
            return VbaGitSnapshot.Read(files);
        }

        /// <summary>Crée un commit contenant le snapshot VBA et le reste de l’arbre parent.</summary>
        /// <param name="snapshot">État sérialisé des fichiers VBA.</param>
        /// <param name="parent">Commit parent du nouvel état, ou null pour une racine.</param>
        /// <param name="message">Message du commit à créer.</param>
        /// <returns>Identifiant du commit créé.</returns>
        internal string Commit(VbaGitSnapshot snapshot, string parent, string message)
        {
            if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException(UiText.Get("A commit message is required."));
            var blobs = new List<string>();
            foreach (var file in snapshot.Serialize())
            {
                string hash = Encoding.UTF8.GetString(Run(new[] { "hash-object", "-w", "--stdin" }, file.Value).Bytes).Trim();
                blobs.Add("100644 blob " + hash + "\t" + file.Key);
            }
            string subtree = MakeTree(blobs);
            var root = parent == null ? new List<string>() : Tree(parent).Where(x => !x.EndsWith("\tvba", StringComparison.Ordinal)).ToList();
            root.Add("040000 tree " + subtree + "\tvba");
            string tree = MakeTree(root);
            var args = new List<string> { "commit-tree", tree };
            if (parent != null) { args.Add("-p"); args.Add(parent); }
            return Encoding.UTF8.GetString(Run(args.ToArray(), VbaGitSnapshot.Utf8.GetBytes(message + "\n")).Bytes).Trim();
        }

        /// <summary>Met à jour une référence Git avec contrôle de sa valeur précédente.</summary>
        /// <param name="reference">Référence Git à résoudre ou mettre à jour.</param>
        /// <param name="commit">Commit dont le paquet VBA est lu ou poussé.</param>
        internal void SetRef(string reference, string commit)
        {
            string previous = Resolve(reference);
            Text("update-ref", reference, commit, previous ?? new string('0', 40));
        }

        /// <summary>Pousse uniquement la branche active vers origin.</summary>
        /// <param name="commit">Commit dont le paquet VBA est lu ou poussé.</param>
        internal void Push(string commit)
        {
            // Push only the selected branch; private backups never leave the cache.
            Text("push", "--porcelain", "origin", commit + ":refs/heads/" + Branch);
            SetRef("refs/remotes/origin/selected", commit);
        }

        /// <summary>Enregistre une sauvegarde et un marqueur avant l’import VBA.</summary>
        /// <param name="snapshot">État sérialisé des fichiers VBA.</param>
        internal void PrepareRecovery(VbaGitSnapshot snapshot)
        {
            if (RecoveryPending) throw new InvalidOperationException(UiText.Get("Restore the interrupted import before continuing."));
            string backup = Commit(snapshot, Resolve(Backup), UiText.Get("VBA backup before import"));
            SetRef(Backup, backup);
            string previousAfter = Resolve(AfterImport);
            if (previousAfter != null) Text("update-ref", "-d", AfterImport, previousAfter);
            File.WriteAllText(RecoveryFile, backup, Encoding.ASCII);
        }

        /// <summary>Enregistre l’état obtenu après import dans une référence privée.</summary>
        /// <param name="state">État VBA à enregistrer après import.</param>
        internal void RecordImportedState(VbaGitSnapshot state)
        {
            SetRef(AfterImport, Commit(state, null, UiText.Get("VBA state after import")));
        }

        /// <summary>Supprime le marqueur après une récupération terminée.</summary>
        internal void CompleteRecovery()
        {
            var attributes = ObserveRecoveryMarker();
            if (!attributes.HasValue) return;
            RequireRegularRecoveryMarker(attributes.Value);
            DeleteRecoveryMarker(RecoveryFile);
            if (ObserveRecoveryMarker().HasValue)
                throw new IOException("The recovery marker remains after its single deletion request; completion is unverified. Do not retry automatically.");
        }

        /// <summary>Retourne les entrées directes de l’arbre Git.</summary>
        /// <param name="tree">Identifiant de commit ou arbre Git à lire.</param>
        /// <returns>Entrées directes de l’arbre.</returns>
        private List<string> Tree(string tree)
        {
            return Encoding.UTF8.GetString(Run(new[] { "ls-tree", "-z", tree }).Bytes).Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        /// <summary>Crée un arbre Git à partir des entrées fournies.</summary>
        /// <param name="entries">Entrées de l’arbre à créer.</param>
        /// <returns>Identifiant de l’arbre Git créé.</returns>
        private string MakeTree(IEnumerable<string> entries)
        {
            return Encoding.UTF8.GetString(Run(new[] { "mktree", "-z" }, VbaGitSnapshot.Utf8.GetBytes(string.Join("\0", entries) + "\0")).Bytes).Trim();
        }

        /// <summary>Exécute Git et décode la sortie standard en texte UTF-8.</summary>
        /// <returns>Sortie standard décodée et sans espaces terminaux.</returns>
        /// <param name="args">Arguments Git à transmettre.</param>
        private string Text(params string[] args) { return Encoding.UTF8.GetString(Run(args).Bytes).Trim(); }

        /// <summary>Sortie binaire et code de retour d’un processus Git.</summary>
        internal sealed class Result
        {

            /// <summary>Octets écrits sur la sortie standard du processus.</summary>
            internal byte[] Bytes;

            /// <summary>Code de sortie du processus Git.</summary>
            internal int ExitCode;
        }

        // Windows argv quoting. No shell, no command interpolation, no interactive terminal prompts.
        /// <summary>Échappe un argument selon les règles de la ligne de commande Windows.</summary>
        /// <param name="value">Argument à transmettre au processus Git.</param>
        /// <returns>Argument protégé pour la ligne de commande Windows.</returns>
        private static string Quote(string value)
        {
            return "\"" + Regex.Replace(value, "(\\\\*)\"", "$1$1\\\"").TrimEnd('\\') +
                new string('\\', value.Reverse().TakeWhile(x => x == '\\').Count() * 2) + "\"";
        }

        /// <summary>Exécute Git sans shell, avec annulation, progression et contrôle des erreurs.</summary>
        /// <param name="args">Arguments Git transmis sans interprétation par un shell.</param>
        /// <param name="input">Octets envoyés à l’entrée standard de Git, le cas échéant.</param>
        /// <param name="useRepository">Indique si Git reçoit le répertoire bare avec --git-dir.</param>
        /// <param name="allowFailure">Autorise le retour d’un code non nul au lieu de lever une exception.</param>
        /// <returns>Résultat contenant les octets de sortie et le code de sortie.</returns>
        private Result Run(string[] args, byte[] input = null, bool useRepository = true, bool allowFailure = false)
        {
            Cancellation.ThrowIfCancellationRequested();
            if (CommandOverride != null) return CommandOverride(args, input, useRepository, allowFailure);
            Progress?.Invoke(UiText.Get("Operation in progress…") + " · " + UiText.Get(args[0] == "push" ? "Push" :
                args[0] == "fetch" || args[0] == "ls-remote" ? "Fetch" : args[0] == "log" || args[0] == "rev-list" ? "History" : "Git changes"));
            var all = new List<string> { "-c", "core.hooksPath=" + Path.Combine(directory, "disabled-hooks"), "-c", "commit.gpgSign=false",
                "-c", "core.longpaths=true" };
            if (!string.IsNullOrEmpty(account))
                all.AddRange(new[] { "-c", "credential.helper=", "-c", "credential.helper=manager", "-c", "credential.https://github.com.username=" + account });
            // The child already starts inside this bare cache. Git for Windows rejects an
            // explicit absolute GIT_DIR over PATH_MAX - 40 even when core.longpaths is enabled.
            if (useRepository) all.Add("--git-dir=.");
            all.AddRange(args);
            var start = new ProcessStartInfo("git.exe", string.Join(" ", all.Select(Quote)))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                WorkingDirectory = directory
            };
            foreach (string key in start.EnvironmentVariables.Keys.Cast<string>().Where(x => x.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToArray())
                start.EnvironmentVariables.Remove(key);
            start.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            using (var process = new Process { StartInfo = start })
            using (var output = new MemoryStream())
            {
                if (!StartProcess(process)) throw new InvalidOperationException(UiText.Get("Unable to start Git."));
                Task read = process.StandardOutput.BaseStream.CopyToAsync(output);
                Task<string> error = process.StandardError.ReadToEndAsync();
                Task write = Task.Run(() =>
                {
                    try { if (input != null) process.StandardInput.BaseStream.Write(input, 0, input.Length); }
                    finally { process.StandardInput.Close(); }
                });
                using (Cancellation.Register(() => { try { if (!process.HasExited) StopProcess(process); } catch (InvalidOperationException) { } }))
                    if (!WaitForExit(process, 120000)) { try { StopProcess(process); } catch { } throw new TimeoutException(UiText.Get("Git did not respond within 120 seconds. Check the connection and GitHub authentication.")); }
                try { Task.WaitAll(read, error, write); }
                catch (Exception) when (Cancellation.IsCancellationRequested) { throw new OperationCanceledException(Cancellation); }
                Cancellation.ThrowIfCancellationRequested();
                if (process.ExitCode != 0 && !allowFailure)
                    throw new InvalidOperationException("Git " + args[0] + UiText.Get(" failed. ") + error.Result.Trim());
                return new Result { Bytes = output.ToArray(), ExitCode = process.ExitCode };
            }
        }
    }
}
