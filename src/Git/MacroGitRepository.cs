using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CodexVBE
{
    // Bare repository: no working tree, no checkout of remote files and no Git hooks.
    internal sealed partial class MacroGitRepository
    {
        private readonly string directory;
        private readonly string account;
        internal string Branch { get; private set; }
        internal const string Baseline = "refs/codex/baseline";
        internal const string Backup = "refs/codex/backup";
        internal const string AfterImport = "refs/codex/after-import";
        internal string Head { get { return "refs/heads/" + Branch; } }
        internal string RecoveryFile { get { return Path.Combine(directory, "codex-recovery"); } }
        internal bool RecoveryPending { get { return File.Exists(RecoveryFile); } }

        internal MacroGitRepository(string directory, string branch, string account = null)
        {
            if (!Regex.IsMatch(branch ?? "", @"^[A-Za-z0-9][A-Za-z0-9_./-]{0,127}$") || branch.Contains("..") ||
                branch.Contains("//") || branch.EndsWith("/") || branch.Split('/').Any(x => x.StartsWith(".") || x.EndsWith(".") || x.EndsWith(".lock")))
                throw new ArgumentException("Nom de branche Git invalide.");
            if (!string.IsNullOrEmpty(account) && !GitHubAccountService.ValidAccount(account)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            this.account = account;
            this.directory = Path.GetFullPath(directory); Branch = branch;
        }

        internal static string ScopeDirectory(string scope)
        {
            using (var sha = SHA256.Create())
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "Git",
                    BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(scope))).Replace("-", "").ToLowerInvariant());
        }

        internal static string ValidateRemote(string remote)
        {
            remote = (remote ?? "").Trim();
            if (!Regex.IsMatch(remote, @"^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+(?:\.git)?/?$", RegexOptions.IgnoreCase))
                throw new ArgumentException(UiText.Get("Enter the repository's HTTPS GitHub URL without a token or password."));
            return remote.TrimEnd('/');
        }

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

        internal string Resolve(string reference)
        {
            var result = Run(new[] { "rev-parse", "--verify", "--quiet", reference }, null, true, true);
            return result.ExitCode == 0 ? Encoding.UTF8.GetString(result.Bytes).Trim() : null;
        }

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

        internal string[] History()
        {
            string head = Resolve(Head);
            return head == null ? new string[0] : Text("log", "-40", "--date=short", "--format=%h  %ad  %s", head)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        internal string SynchronizationStatus()
        {
            string head = Resolve(Head), remote = Resolve("refs/remotes/origin/selected");
            if (head == null) return remote == null ? UiText.Get("No local commit · use Fetch to inspect the remote repository") : "Branche distante disponible · premier pull requis";
            if (remote == null) return UiText.Get("Local branch · remote state unknown (Fetch)");
            var counts = Text("rev-list", "--left-right", "--count", head + "..." + remote).Split(new[] { '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return "↑ " + counts[0] + " sortant(s)   ↓ " + counts[1] + " entrant(s) · dernier Fetch";
        }

        internal void RequireFastForward(string before, string after)
        {
            if (before == null || before == after) return;
            if (after == null || Run(new[] { "merge-base", "--is-ancestor", before, after }, null, true, true).ExitCode != 0)
                throw new InvalidOperationException(UiText.Get("Histories diverge. Reconcile the branch before synchronizing; no forced push is performed."));
        }

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

        internal void SetRef(string reference, string commit)
        {
            string previous = Resolve(reference);
            Text("update-ref", reference, commit, previous ?? new string('0', 40));
        }

        internal void Push(string commit)
        {
            // Push only the selected branch; private backups never leave the cache.
            Text("push", "--porcelain", "origin", commit + ":refs/heads/" + Branch);
            SetRef("refs/remotes/origin/selected", commit);
        }

        internal void PrepareRecovery(VbaGitSnapshot snapshot)
        {
            string backup = Commit(snapshot, Resolve(Backup), UiText.Get("VBA backup before import"));
            SetRef(Backup, backup);
            string previousAfter = Resolve(AfterImport);
            if (previousAfter != null) Text("update-ref", "-d", AfterImport, previousAfter);
            File.WriteAllText(RecoveryFile, backup, Encoding.ASCII);
        }

        internal void RecordImportedState(VbaGitSnapshot state)
        {
            SetRef(AfterImport, Commit(state, null, UiText.Get("VBA state after import")));
        }

        internal void CompleteRecovery() { if (File.Exists(RecoveryFile)) File.Delete(RecoveryFile); }

        private List<string> Tree(string tree)
        {
            return Encoding.UTF8.GetString(Run(new[] { "ls-tree", "-z", tree }).Bytes).Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }
        private string MakeTree(IEnumerable<string> entries)
        {
            return Encoding.UTF8.GetString(Run(new[] { "mktree", "-z" }, VbaGitSnapshot.Utf8.GetBytes(string.Join("\0", entries) + "\0")).Bytes).Trim();
        }
        private string Text(params string[] args) { return Encoding.UTF8.GetString(Run(args).Bytes).Trim(); }
        private sealed class Result { internal byte[] Bytes; internal int ExitCode; }

        // Windows argv quoting. No shell, no command interpolation, no interactive terminal prompts.
        private static string Quote(string value)
        {
            return "\"" + Regex.Replace(value, "(\\\\*)\"", "$1$1\\\"").TrimEnd('\\') +
                new string('\\', value.Reverse().TakeWhile(x => x == '\\').Count() * 2) + "\"";
        }
        private Result Run(string[] args, byte[] input = null, bool useRepository = true, bool allowFailure = false)
        {
            var all = new List<string> { "-c", "core.hooksPath=" + Path.Combine(directory, "disabled-hooks"), "-c", "commit.gpgSign=false" };
            if (!string.IsNullOrEmpty(account))
                all.AddRange(new[] { "-c", "credential.helper=", "-c", "credential.helper=manager", "-c", "credential.https://github.com.username=" + account });
            if (useRepository) all.Add("--git-dir=" + directory);
            all.AddRange(args);
            var start = new ProcessStartInfo("git.exe", string.Join(" ", all.Select(Quote))) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true,
                RedirectStandardError = true, RedirectStandardInput = true, WorkingDirectory = directory
            };
            foreach (string key in start.EnvironmentVariables.Keys.Cast<string>().Where(x => x.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToArray())
                start.EnvironmentVariables.Remove(key);
            start.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            using (var process = new Process { StartInfo = start })
            using (var output = new MemoryStream())
            {
                if (!ProcessInput.StartWithoutPreamble(process)) throw new InvalidOperationException(UiText.Get("Unable to start Git."));
                Task read = process.StandardOutput.BaseStream.CopyToAsync(output);
                Task<string> error = process.StandardError.ReadToEndAsync();
                Task write = Task.Run(() => { try { if (input != null) process.StandardInput.BaseStream.Write(input, 0, input.Length); }
                    finally { process.StandardInput.Close(); } });
                if (!process.WaitForExit(120000)) { try { process.Kill(); } catch { } throw new TimeoutException(UiText.Get("Git did not respond within 120 seconds. Check the connection and GitHub authentication.")); }
                Task.WaitAll(read, error, write);
                if (process.ExitCode != 0 && !allowFailure)
                    throw new InvalidOperationException("Git " + args[0] + UiText.Get(" failed. ") + error.Result.Trim());
                return new Result { Bytes = output.ToArray(), ExitCode = process.ExitCode };
            }
        }
    }
}
